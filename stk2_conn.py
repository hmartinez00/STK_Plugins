import comtypes.client
import csv
from comtypes.gen import STKObjects

# ==========================
# CONFIGURACIÓN
# ==========================
RUTA_CSV = r"C:\Users\HP\Documents\REQ-87 - figuras_geograficas-afectacion  de terremotpo\targets.csv"
RUTA_SENSOR = "*/Satellite/VRSS-2/Sensor/V2-10"
ALTITUD_DEFAULT = 0.0  # km
PREFIJO_CHAIN = "Chain"


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
        # 1. Conexión al STK 11 activo
        ui_application = comtypes.client.GetActiveObject("STK11.Application")
        root = ui_application.Personality2
        mi_escenario = root.CurrentScenario

        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        print(f"¡Conectado con éxito al escenario: '{mi_escenario.InstanceName}'!\n")

        scenario_children = mi_escenario.Children

        # 2. Obtener la referencia REAL del sensor (una sola vez, fuera del loop)
        sensor_obj = root.GetObjectFromPath(RUTA_SENSOR)
        print(f"-> Sensor obtenido: {sensor_obj.ClassName} en {RUTA_SENSOR}\n")

        # 3. Leer el CSV
        targets_data = leer_targets_csv(RUTA_CSV)
        if not targets_data:
            print("No se encontraron targets válidos en el CSV.")
            return

        print(f"Se encontraron {len(targets_data)} targets en el CSV.\n")

        exitosos = []
        fallidos = []

        # 4. Procesar cada target
        for item in targets_data:
            nombre_target = item["nombre"]
            lat = item["lat"]
            lon = item["lon"]
            nombre_chain = f"{PREFIJO_CHAIN}_{nombre_target}"

            print(f"--- Procesando: {nombre_target} (lat={lat}, lon={lon}) ---")

            try:
                # 4a. Limpieza de residuos previos
                limpiar_si_existe(scenario_children, nombre_target)
                limpiar_si_existe(scenario_children, nombre_chain)

                # 4b. Crear y posicionar el Target
                target_obj = scenario_children.New(STKObjects.eTarget, nombre_target)
                target = target_obj.QueryInterface(STKObjects.IAgTarget)
                target.Position.AssignGeodetic(lat, lon, ALTITUD_DEFAULT)
                print(f"   -> Target '{nombre_target}' creado y posicionado.")

                # 4c. Crear el Chain y AGREGAR LOS OBJETOS REALES (no strings)
                chain_raw = scenario_children.New(STKObjects.eChain, nombre_chain)
                chain = chain_raw.QueryInterface(STKObjects.IAgChain)

                chain.Objects.AddObject(sensor_obj)
                chain.Objects.AddObject(target_obj)
                print(f"   -> Chain '{nombre_chain}' creado, vinculando V2-10 y {nombre_target}.")

                # 4d. Calcular acceso
                chain.ComputeAccess()
                print(f"   -> Acceso calculado para '{nombre_chain}'.\n")

                exitosos.append(nombre_target)

            except Exception as e_item:
                print(f"   [ERROR] Falló el procesamiento de '{nombre_target}': {e_item}\n")
                fallidos.append(nombre_target)

        # 5. Resumen final
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