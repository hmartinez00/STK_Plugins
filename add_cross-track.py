import comtypes.client
import pandas as pd
from datetime import datetime
from comtypes.gen import STKObjects

# ========================= CONFIGURACIÓN =========================
RUTA_SENSOR = "*/Satellite/VRSS-2/Sensor/V2-10"
PREFIJO_CHAIN = "Chain"
step_time = 30.0

output_file = f"VRSS2_Roll_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
# ============================================================

ui_app = comtypes.client.GetActiveObject("STK11.Application")
root = ui_app.Personality2
scenario = root.CurrentScenario
print(f"Escenario: {scenario.InstanceName}\n")

# Usar el rango de tiempo real del escenario
scenario_iface = scenario.QueryInterface(STKObjects.IAgScenario)
start_time = scenario_iface.StartTime
stop_time = scenario_iface.StopTime
print(f"Rango de tiempo del escenario: {start_time} -> {stop_time}\n")

children = scenario.Children
chains = []

print("Buscando Chains...")
for i in range(children.Count):
    try:
        obj = children.Item(i)
        if obj.ClassName == "Chain" and obj.InstanceName.startswith(PREFIJO_CHAIN):
            chains.append(obj)
            print(f"  → {obj.InstanceName}")
    except:
        continue

print(f"\nTotal Chains encontrados: {len(chains)}\n")

all_data = []
sensor_obj = root.GetObjectFromPath(RUTA_SENSOR)  # una sola vez, fuera del loop

for chain_obj in chains:
    try:
        chain_name = chain_obj.InstanceName
        chain = chain_obj.QueryInterface(STKObjects.IAgChain)

        assigned = chain.Objects
        target_obj = None

        for j in range(assigned.Count):
            link = assigned.Item(j)
            item = link.LinkedObject
            if item.ClassName in ["Target", "Facility", "Place"] or "Target" in item.InstanceName:
                target_obj = item
                break

        if target_obj is None:
            print(f"   ⚠ No se encontró target en {chain_name}")
            continue

        target_name = target_obj.InstanceName
        print(f"Procesando → {target_name}")

        access = sensor_obj.GetAccessToObject(target_obj)
        access.ComputeAccess()

        # AER Data -> Default (Azimuth, Elevation, Range)
        dp_info_aer = access.DataProviders.Item("AER Data")
        dp_group_aer = dp_info_aer.QueryInterface(STKObjects.IAgDataProviderGroup)
        default_info = dp_group_aer.Group.Item("Default")
        dp_timevar_aer = default_info.QueryInterface(STKObjects.IAgDataPrvTimeVar)
        result_aer = dp_timevar_aer.Exec(start_time, stop_time, step_time)

        # Sensor Based Angles -> BodyFixed (Roll/Cross-Track real del sensor)
        dp_info_sba = access.DataProviders.Item("Sensor Based Angles")
        dp_group_sba = dp_info_sba.QueryInterface(STKObjects.IAgDataProviderGroup)
        bodyfixed_info = dp_group_sba.Group.Item("BodyFixed")
        dp_timevar_sba = bodyfixed_info.QueryInterface(STKObjects.IAgDataPrvTimeVar)
        result_sba = dp_timevar_sba.Exec(start_time, stop_time, step_time)

        # Construir DataFrame combinando ambos data providers
        df = pd.DataFrame({
            'Time_UTC': result_aer.DataSets.GetDataSetByName("Time").GetValues(),
            'Azimuth_deg': result_aer.DataSets.GetDataSetByName("Azimuth").GetValues(),
            'Elevation_deg': result_aer.DataSets.GetDataSetByName("Elevation").GetValues(),
            'Range_km': result_aer.DataSets.GetDataSetByName("Range").GetValues(),
            'AngleOffBoresight_deg': result_sba.DataSets.GetDataSetByName("Angle Off Boresight").GetValues(),
            'Roll_CrossTrack_deg': result_sba.DataSets.GetDataSetByName("Horizontal Angle").GetValues(),
            'Pitch_AlongTrack_deg': result_sba.DataSets.GetDataSetByName("Vertical Angle").GetValues(),
            'DistanceOffBoresight_km': result_sba.DataSets.GetDataSetByName("Distance Off Boresight").GetValues(),
            'Target': target_name,
            'Chain': chain_name
        })

        all_data.append(df)
        print(f"   ✓ {len(df)} registros\n")

    except Exception as e:
        print(f"   ✗ Error en {chain_name if 'chain_name' in locals() else 'desconocido'}: {e}\n")

# Exportar
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)

    summary = final_df.groupby('Target').agg({
        'Roll_CrossTrack_deg': ['min', 'max', 'mean'],
        'AngleOffBoresight_deg': ['min', 'max'],
        'Range_km': 'min',
        'Time_UTC': 'count'
    }).round(3)

    with pd.ExcelWriter(output_file) as writer:
        final_df.to_excel(writer, sheet_name="Detailed", index=False)
        summary.to_excel(writer, sheet_name="Summary")

    print(f"\n¡ÉXITO! Archivo creado: {output_file}")
else:
    print("No se generaron datos.")