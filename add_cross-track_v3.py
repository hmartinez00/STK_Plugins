import comtypes.client
import pandas as pd
from datetime import datetime
from comtypes.gen import STKObjects

# ========================= CONFIGURACIÓN =========================
RUTA_SENSOR = "*/Satellite/VRSS-2/Sensor/V2-10"
PREFIJO_CHAIN = "Chain"
step_time = 10.0
GAP_MINUTOS_NUEVA_PASADA = 30  # si el salto entre registros supera esto, es una pasada distinta

output_file = f"VRSS2_Roll_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
# ============================================================

ui_app = comtypes.client.GetActiveObject("STK11.Application")
root = ui_app.Personality2
scenario = root.CurrentScenario
print(f"Escenario: {scenario.InstanceName}\n")

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
sensor_obj = root.GetObjectFromPath(RUTA_SENSOR)


def parse_stk_time(t):
    """Convierte 'dd Mon yyyy HH:MM:SS.ffffff' a datetime de Python."""
    return datetime.strptime(t.split('.')[0], "%d %b %Y %H:%M:%S")


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

        # AER Data -> Default (Azimuth, Elevation)
        dp_info_aer = access.DataProviders.Item("AER Data")
        dp_group_aer = dp_info_aer.QueryInterface(STKObjects.IAgDataProviderGroup)
        default_info = dp_group_aer.Group.Item("Default")
        dp_timevar_aer = default_info.QueryInterface(STKObjects.IAgDataPrvTimeVar)
        result_aer = dp_timevar_aer.Exec(start_time, stop_time, step_time)

        # Sat Angles Data (Along Track, Cross Track = Roll real, Range, RangeRate, Path Delay)
        dp_info_sat = access.DataProviders.Item("Sat Angles Data")

        if dp_info_sat.IsGroup():
            dp_group_sat = dp_info_sat.QueryInterface(STKObjects.IAgDataProviderGroup)
            dp_info_sat = dp_group_sat.Group.Item("Default")

        dp_timevar_sat = dp_info_sat.QueryInterface(STKObjects.IAgDataPrvTimeVar)
        result_sat = dp_timevar_sat.Exec(start_time, stop_time, step_time)

        # Construir DataFrame combinando ambos data providers
        df = pd.DataFrame({
            'Time_UTC': result_aer.DataSets.GetDataSetByName("Time").GetValues(),
            'Azimuth_deg': result_aer.DataSets.GetDataSetByName("Azimuth").GetValues(),
            'Elevation_deg': result_aer.DataSets.GetDataSetByName("Elevation").GetValues(),
            'AlongTrack_deg': result_sat.DataSets.GetDataSetByName("Along Track").GetValues(),
            'CrossTrack_Roll_deg': result_sat.DataSets.GetDataSetByName("Cross Track").GetValues(),
            'Range_km': result_sat.DataSets.GetDataSetByName("Range").GetValues(),
            'RangeRate_kmps': result_sat.DataSets.GetDataSetByName("RangeRate").GetValues(),
            'PathDelay_sec': result_sat.DataSets.GetDataSetByName("Path Delay").GetValues(),
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

    # === NUEVO: detectar lotes (pasadas) por salto temporal, dentro de cada Target ===
    final_df['Time_dt'] = final_df['Time_UTC'].apply(parse_stk_time)
    final_df = final_df.sort_values(['Target', 'Time_dt']).reset_index(drop=True)

    # Diferencia de tiempo respecto al registro anterior, DENTRO de cada target
    gap = final_df.groupby('Target')['Time_dt'].diff()

    # Nueva pasada = primer registro de cada target, o salto mayor al umbral
    nueva_pasada = (gap.isna()) | (gap > pd.Timedelta(minutes=GAP_MINUTOS_NUEVA_PASADA))

    # Numerar pasadas de forma incremental, por target
    final_df['Pasada'] = nueva_pasada.groupby(final_df['Target']).cumsum()

    # Dentro de cada Target + Pasada, seleccionar el registro de mínimo Range
    idx_min_range = final_df.groupby(['Target', 'Pasada'])['Range_km'].idxmin()
    optimal_df = final_df.loc[idx_min_range].sort_values(['Target', 'Pasada']).reset_index(drop=True)
    optimal_df = optimal_df.drop(columns=['Time_dt'])  # columna auxiliar, no necesaria en el export
    # ===================================================================================

    final_df_export = final_df.drop(columns=['Time_dt'])

    summary = final_df.groupby('Target').agg({
        'CrossTrack_Roll_deg': ['min', 'max', 'mean'],
        'AlongTrack_deg': ['min', 'max'],
        'Range_km': 'min',
        'Time_UTC': 'count'
    }).round(3)

    with pd.ExcelWriter(output_file) as writer:
        final_df_export.to_excel(writer, sheet_name="Detailed", index=False)
        optimal_df.to_excel(writer, sheet_name="Optimal_per_Target_Pass", index=False)
        summary.to_excel(writer, sheet_name="Summary")

    print(f"\n¡ÉXITO! Archivo creado: {output_file}")
    print(f"Pasadas óptimas detectadas: {len(optimal_df)}")
else:
    print("No se generaron datos.")