import comtypes.client
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta
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
        filetypes=[
            ("Archivos EPH", "*.EPH;*.eph"),
            ("Todos los archivos", "*.*")
        ]
    )

    root_tk.destroy()
    return ruta_seleccionada


def leer_eph(ruta_eph):

    tree = ET.parse(ruta_eph)
    root_xml = tree.getroot()

    body = root_xml.find("FileBody")

    epoch_str = body.find("OrbitEpoch").text

    epoch_dt = datetime.strptime(
        epoch_str,
        "%Y-%m-%dT%H:%M:%S"
    )

    epoch_stk = epoch_dt.strftime(
        "%d %b %Y %H:%M:%S.000"
    )

    return {
        "sma": float(body.find("OrbitRadius").text),
        "ecc": float(body.find("OrbitPartialityRatio").text),
        "inc": float(body.find("OrbitObliquity").text),
        "raan": float(body.find("AscendPoint").text),
        "arg_perigee": float(body.find("NearPointAngle").text),
        "mean_anomaly": float(body.find("BreadthAngle").text),
        "epoch": epoch_stk,
        "epoch_dt": epoch_dt,
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

        print(
            f"Satélite: {elementos['satellite_id']} | "
            f"Órbita: {elementos['orbit_id']}"
        )

        print(f"Época orbital: {elementos['epoch']}")

        print(
            f"SMA: {elementos['sma']} m | "
            f"Ecc: {elementos['ecc']}"
        )

        print(
            f"Inc: {elementos['inc']}° | "
            f"RAAN: {elementos['raan']}°"
        )

        print(
            f"ArgPerigee: {elementos['arg_perigee']}° | "
            f"MeanAnomaly: {elementos['mean_anomaly']}°\n"
        )

        # =====================================================
        # GENERAR NUEVO INTERVALO DEL ESCENARIO
        # =====================================================

        epoch_dt = elementos["epoch_dt"]

        start_dt = epoch_dt.replace(
            hour=4,
            minute=40,
            second=0,
            microsecond=0
        )

        stop_dt = start_dt + timedelta(days=8)

        start_time = start_dt.strftime(
            "%d %b %Y %H:%M:%S.000"
        )

        stop_time = stop_dt.strftime(
            "%d %b %Y %H:%M:%S.000"
        )

        print("Intervalo generado automáticamente:")

        print(f"Start Time : {start_time}")
        print(f"Stop Time  : {stop_time}\n")

        # =====================================================
        # CONEXIÓN A STK
        # =====================================================

        ui_application = comtypes.client.GetActiveObject(
            "STK11.Application"
        )

        root = ui_application.Personality2

        mi_escenario = root.CurrentScenario

        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        print(
            f"¡Conectado con éxito al escenario: "
            f"'{mi_escenario.InstanceName}'!\n"
        )

        scenario_iface = mi_escenario.QueryInterface(
            STKObjects.IAgScenario
        )

        # =====================================================
        # ACTUALIZAR INTERVALO DEL ESCENARIO
        # =====================================================

        print("Actualizando intervalo temporal del escenario...")

        scenario_iface.SetTimePeriod(
            start_time,
            stop_time
        )

        root.Rewind()

        print("Intervalo actualizado correctamente.\n")

        # =====================================================
        # ACTUALIZAR ESTADO ORBITAL HPOP
        # =====================================================

        cmd_setstate = (
            f'SetState {RUTA_SATELITE} Classical HPOP '
            f'"{start_time}" "{stop_time}" '
            f'{STEP_PROPAGACION} '
            f'J2000 '
            f'"{elementos["epoch"]}" '
            f'{elementos["sma"]} '
            f'{elementos["ecc"]} '
            f'{elementos["inc"]} '
            f'{elementos["raan"]} '
            f'{elementos["arg_perigee"]} '
            f'{elementos["mean_anomaly"]}'
        )

        print("Comando enviado a STK:\n")
        print(cmd_setstate)
        print()

        root.ExecuteCommand(cmd_setstate)

        print(
            f"¡ÉXITO! Órbita de "
            f"{elementos['satellite_id']} "
            f"actualizada correctamente."
        )

    except Exception as e:

        print("\n[ERROR FATAL]")
        print(e)


if __name__ == "__main__":
    cargar_orbita()

