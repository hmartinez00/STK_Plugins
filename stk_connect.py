import comtypes.client
from comtypes.gen import STKObjects

def crear_target_object_model():
    try:
        ui_application = comtypes.client.GetActiveObject("STK11.Application")
        root = ui_application.Personality2
        mi_escenario = root.CurrentScenario

        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        nombre_target = "caracas_Test-2"
        scenario_children = mi_escenario.Children

        # 1. LIMPIEZA: eliminar si ya existe
        try:
            existente = scenario_children.Item(nombre_target)
            scenario_children.Unload(existente)
            print("-> Limpiado residuo previo.")
        except Exception:
            print("-> No había residuo previo.")

        # 2. CREACIÓN del Target
        nuevo_obj = scenario_children.New(STKObjects.eTarget, nombre_target)
        target = nuevo_obj.QueryInterface(STKObjects.IAgTarget)
        print(f"-> Target '{nombre_target}' creado.")

        # 3. POSICIONAMIENTO: aquí van lat/lon/alt
        latitud = 10.6000
        longitud = -66.8000
        altitud = 0.0  # en km

        # target.Position es un IAgPosition
        posicion = target.Position
        posicion.AssignGeodetic(latitud, longitud, altitud)

        print(f"-> Posición asignada: lat={latitud}, lon={longitud}, alt={altitud} km")
        print(f"\n¡ÉXITO TOTAL! Target '{nombre_target}' creado y ubicado en Caracas.")

    except Exception as e:
        print(f"\n[ERROR FATAL] {e}")

crear_target_object_model()