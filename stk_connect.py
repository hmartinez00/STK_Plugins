import comtypes.client

def crear_target_perfecto_connect():
    try:
        # 1. Conexión al STK 11 activo
        ui_application = comtypes.client.GetActiveObject("STK11.Application")
        root = ui_application.Personality2
        mi_escenario = root.CurrentScenario
        
        if mi_escenario is None:
            print("No hay escenario abierto en STK 11.")
            return

        escenario_nombre = mi_escenario.InstanceName
        nombre_target = "caracas_Test"
        
        print(f"¡Conectado con éxito al escenario: '{escenario_nombre}'!")

        # 2. LIMPIEZA: Usamos el nombre EXPLÍCITO del escenario en lugar de */
        for clase_obj in ["Target", "Facility", "Sensor"]:
            try:
                # Ruta explícita: NombreEscenario/Clase/NombreObjeto
                root.ExecuteCommand(f"Unload / {escenario_nombre}/{clase_obj}/{nombre_target}")
                print(f"-> Limpiado residuo de tipo '{clase_obj}'.")
            except:
                # Si no existía, STK lanza error y lo ignoramos
                pass

        # 3. CREACIÓN: Usamos la ruta explícita
        # Sintaxis correcta: NewObj / <RutaEscenario>/Target <Nombre>
        print(f"Creando el Target '{nombre_target}' vía Connect...")
        root.ExecuteCommand(f"NewObj / {escenario_nombre}/Target {nombre_target}")

        # 4. POSICIONAMIENTO: Coordenadas geodésicas de Caracas
        latitud = 10.5000
        longitud = -66.9000
        altitud = 0.0  # En kilómetros
        
        # Ruta explícita para SetPosition (OJO: SetPosition no lleva barra / después del nombre del comando)
        comando_posicion = f"SetPosition {escenario_nombre}/Target/{nombre_target} Geodetic {latitud} {longitud} {altitud}"
        
        print(f"Configurando coordenadas geodésicas...")
        root.ExecuteCommand(comando_posicion)
        
        print(f"¡Target '{nombre_target}' creado y ubicado exitosamente en Caracas!")
        
    except Exception as e:
        print(f"Error en la operación: {e}")

# Ejecutar el script
crear_target_perfecto_connect()