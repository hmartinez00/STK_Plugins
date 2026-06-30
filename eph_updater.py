import comtypes.client
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from comtypes.gen import STKObjects

# ==========================
# CONFIGURACIÓN
# ==========================
RUTA_SATELITE = "*/Satellite/VRSS-2"
STEP_PROPAGACION = 60  # segundos


def seleccionar_archivo_eph():
    root_tk = tk.Tk()
    root_tk.withdraw()
    root_tk.attributes('-topmost', True)

    ruta_seleccionada = filedialog.askopenfilename(
        title="Selecciona el archivo de efemérides (.EPH)",
        filetypes=[("Archivos EPH", "*.EPH;*.eph"), ("Todos los archivos", "*.*")]
    )

    root_tk.destroy()
    return ruta_seleccionada


def leer_eph(ruta_eph):
    tree = ET.parse(ruta_eph)
    root_xml = tree.getroot()
    body = root_xml.find("FileBody")

    epoch_str = body.find("OrbitEpoch").text
    epoch_dt = datetime.strptime(epoch_str, "%Y-%m-%dT%H:%M:%S")
    epoch_stk = epoch_dt.strftime("%d %b %Y %H:%M:%S.000")

    return {
        "sma": float(body.find("OrbitRadius").text),
        "ecc": float(body.find("OrbitPartialityRatio").text),
        "inc": float(body.find("OrbitObliquity").text),
        "raan": float(body.find("AscendPoint").text),
        "arg_perigee": float(body.find("NearPointAngle").text),
        "mean_anomaly": float(body.find("BreadthAngle").text),
        "epoch": epoch_stk,
        "satellite_id": body.find("SatelliteID").text,
        "orbit_id": body.find("OrbitID").text,
    }


def cargar_orbita():
    try:
        ruta_eph = seleccionar_archivo_eph()
        if not ruta_eph:
            print("No se seleccionó ningún archivo. Operación cancelada.")
            return

        print(f"Archivo seleccionado: {ruta_eph}\n")

        elementos = leer_eph(ruta_eph)
        print(f"Satélite: {elementos['satellite_id']} | Órbita: {elementos['orbit_id']}")
        print(f"Época: {elementos['epoch']}")
        print(f"SMA: {elementos['sma']} m | Ecc: {elementos['ecc']}")
        print(f"Inc: {elementos['inc']}° | RAAN: {elementos['raan']}°")
        print(f"ArgPerigee: {elementos['arg_perigee']}° | MeanAnomaly: {elementos['mean_anomaly']}°\n")

        ui_application = comtypes.client.GetActiveObject("STK11.Application")
        root = ui_application.Personality2
        mi_escenario = root.CurrentScenario

        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        print(f"¡Conectado con éxito al escenario: '{mi_escenario.InstanceName}'!\n")

        scenario_iface = mi_escenario.QueryInterface(STKObjects.IAgScenario)
        start_time = scenario_iface.StartTime
        stop_time = scenario_iface.StopTime

        # Preservando HPOP: solo se actualiza la condición inicial,
        # el force model existente del satélite se mantiene intacto
        cmd_setstate = (
            f'SetState {RUTA_SATELITE} Classical HPOP '
            f'"{start_time}" "{stop_time}" {STEP_PROPAGACION} J2000 '
            f'"{elementos["epoch"]}" '
            f'{elementos["sma"]} {elementos["ecc"]} {elementos["inc"]} '
            f'{elementos["raan"]} {elementos["arg_perigee"]} {elementos["mean_anomaly"]}'
        )

        print(f"[CMD] {cmd_setstate}\n")
        root.ExecuteCommand(cmd_setstate)

        print(f"¡ÉXITO! Órbita de {elementos['satellite_id']} actualizada con HPOP, preservando el force model existente.")

    except Exception as e:
        print(f"\n[ERROR FATAL] {e}")


cargar_orbita()