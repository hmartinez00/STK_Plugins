import comtypes.client
import csv
import tkinter as tk
from tkinter import filedialog
from comtypes.gen import STKObjects

# ==========================
# CONFIGURACIÓN
# ==========================
RUTA_SENSOR = "*/Satellite/VRSS-2/Sensor/V2-10"
ALTITUD_DEFAULT = 0.0  # km
PREFIJO_CHAIN = "Chain"

COLOR_TARGET = "yellow"
COLOR_CHAIN = "yellow"  # verde puro en RGB


def seleccionar_archivo_csv():
    """Abre un explorador de archivos para elegir el CSV de targets."""
    root_tk = tk.Tk()
    root_tk.withdraw()  # oculta la ventana principal vacía de tkinter
    root_tk.attributes('-topmost', True)  # asegura que el diálogo aparezca al frente

    ruta_seleccionada = filedialog.askopenfilename(
        title="Selecciona el archivo CSV de targets",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )

    root_tk.destroy()
    return ruta_seleccionada


def leer_targets_csv(ruta_csv):
    """Lee el CSV y devuelve una lista de dicts: {nombre, lat, lon}"""
    targets = []
    encodings_a_probar = ["utf-8-sig", "cp1252", "latin-1"]
    contenido = None

    for enc in encodings_a_probar:
        try:
            with open(ruta_csv, mode="r", encoding=enc, newline="") as f:
                contenido = f.read()
            print(f"[INFO] CSV leído correctamente con encoding: {enc}")
            break
        except UnicodeDecodeError:
            continue

    if contenido is None:
        raise ValueError("No se pudo leer el CSV con ningún encoding probado.")

    lector = csv.DictReader(contenido.splitlines())
    for fila in lector:
        try:
            nombre = fila["nombre"].strip()
            lat = float(fila["latitud"])
            lon = float(fila["longitud"])
            targets.append({"nombre": nombre, "lat": lat, "lon": lon})
        except (KeyError, ValueError) as e:
            print(f"[AVISO] Fila inválida ignorada: {fila} -> {e}")

    return targets


def limpiar_si_existe(scenario_children, nombre):
    try:
        obj = scenario_children.Item(nombre)
        scenario_children.Unload(obj)
        print(f"   -> Limpiado residuo previo: {nombre}")
    except Exception:
        pass


def crear_targets_y_chains():
    try:
        # === NUEVO: seleccionar el CSV mediante explorador de archivos ===
        ruta_csv = seleccionar_archivo_csv()

        if not ruta_csv:
            print("No se seleccionó ningún archivo. Operación cancelada.")
            return

        print(f"Archivo seleccionado: {ruta_csv}\n")
        # ===================================================================

        ui_application = comtypes.client.GetActiveObject("STK11.Application")
        root = ui_application.Personality2
        mi_escenario = root.CurrentScenario

        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        print(f"¡Conectado con éxito al escenario: '{mi_escenario.InstanceName}'!\n")

        scenario_children = mi_escenario.Children

        sensor_obj = root.GetObjectFromPath(RUTA_SENSOR)
        print(f"-> Sensor obtenido: {sensor_obj.ClassName} en {RUTA_SENSOR}\n")

        targets_data = leer_targets_csv(ruta_csv)
        if not targets_data:
            print("No se encontraron targets válidos en el CSV.")
            return

        print(f"Se encontraron {len(targets_data)} targets en el CSV.\n")

        exitosos = []
        fallidos = []

        for item in targets_data:
            nombre_target = item["nombre"]
            lat = item["lat"]
            lon = item["lon"]
            nombre_chain = f"{PREFIJO_CHAIN}_{nombre_target}"

            print(f"--- Procesando: {nombre_target} (lat={lat}, lon={lon}) ---")

            try:
                limpiar_si_existe(scenario_children, nombre_target)
                limpiar_si_existe(scenario_children, nombre_chain)

                target_obj = scenario_children.New(STKObjects.eTarget, nombre_target)
                target = target_obj.QueryInterface(STKObjects.IAgTarget)
                target.Position.AssignGeodetic(lat, lon, ALTITUD_DEFAULT)
                print(f"   -> Target '{nombre_target}' creado y posicionado.")

                ruta_target = f"*/Target/{nombre_target}"
                cmd_color_target = f"Graphics {ruta_target} SetColor {COLOR_TARGET}"
                root.ExecuteCommand(cmd_color_target)
                print(f"   -> Color del Target asignado: {COLOR_TARGET}")

                chain_raw = scenario_children.New(STKObjects.eChain, nombre_chain)
                chain = chain_raw.QueryInterface(STKObjects.IAgChain)

                chain.Objects.AddObject(sensor_obj)
                chain.Objects.AddObject(target_obj)
                print(f"   -> Chain '{nombre_chain}' creado, vinculando V2-10 y {nombre_target}.")

                ruta_chain = f"*/Chain/{nombre_chain}"
                cmd_color_chain = f"Graphics {ruta_chain} SetColor {COLOR_CHAIN}"
                root.ExecuteCommand(cmd_color_chain)
                print(f"   -> Color del Chain asignado: {COLOR_CHAIN}")

                chain.ComputeAccess()
                print(f"   -> Acceso calculado para '{nombre_chain}'.\n")

                exitosos.append(nombre_target)

            except Exception as e_item:
                print(f"   [ERROR] Falló el procesamiento de '{nombre_target}': {e_item}\n")
                fallidos.append(nombre_target)

        print("=" * 50)
        print(f"RESUMEN: {len(exitosos)} exitosos, {len(fallidos)} fallidos.")
        if exitosos:
            print(f"  Exitosos: {', '.join(exitosos)}")
        if fallidos:
            print(f"  Fallidos: {', '.join(fallidos)}")
        print("=" * 50)

    except Exception as e:
        print(f"\n[ERROR FATAL] {e}")


crear_targets_y_chains()