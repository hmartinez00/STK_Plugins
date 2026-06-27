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
        
        # Ruta del objeto (para Unload, SetPosition, etc.)
        ruta_stk_objeto = f"*/Target/{nombre_target}"
        
        print(f"¡Conectado con éxito al escenario: '{escenario_nombre}'!\n")

        # 2. LIMPIEZA: Eliminamos residuos previos si existen
        cmd_unload = f"Unload / {ruta_stk_objeto}"
        print(f"[CMD] {cmd_unload}")
        try:
            root.ExecuteCommand(cmd_unload)
            print("-> Limpiado residuo previo.\n")
        except:
            print("-> No había residuo previo.\n")

        # 3. CREACIÓN: sintaxis correcta -> New / <Path_Padre> <Nombre>
        cmd_new = f"New / */Target {nombre_target}"
        print(f"[CMD] {cmd_new}")
        root.ExecuteCommand(cmd_new)
        print("-> Target creado exitosamente.\n")

        # 4. POSICIONAMIENTO
        latitud = 10.5000
        longitud = -66.9000
        altitud = 0.0  
        
        cmd_pos = f"SetPosition {ruta_stk_objeto} Geodetic {latitud} {longitud} {altitud}"
        print(f"[CMD] {cmd_pos}")
        root.ExecuteCommand(cmd_pos)
        
        print(f"\n¡ÉXITO TOTAL! Target '{nombre_target}' creado y ubicado en Caracas.")
        
    except Exception as e:
        print(f"\n[ERROR FATAL] {e}")

# Ejecutar el script
crear_target_perfecto_connect()