import pandas as pd
import io
from datetime import datetime

# ================= CONFIGURACIÓN =================
archivo = r"Satellite-VRSS-2-To-Target-Venezuela_Aragua-Ocumare-de-la-Costa-Sismo_DAE Access + Sat Angles.csv"          # ← Cambia por el nombre exacto
output  = "Resumen_Access_Mejor_Roll.xlsx"
# ================================================

print("Leyendo archivo...")

with open(archivo, 'r', encoding='utf-8') as f:
    texto = f.read()

bloques = texto.split('\n\n')

df_access = None
df_angles = None

for bloque in bloques:
    if not bloque.strip():
        continue
    if "Start Time (UTCG)" in bloque and "Duration (sec)" in bloque:
        df_access = pd.read_csv(io.StringIO(bloque), on_bad_lines='skip')
    elif "Time (UTCG)" in bloque and "Along Track" in bloque:
        df_angles = pd.read_csv(io.StringIO(bloque), on_bad_lines='skip')

print(f"Access: {len(df_access) if df_access is not None else 0}")
print(f"Angles: {len(df_angles) if df_angles is not None else 0}")

if df_access is None or df_angles is None:
    print("No se detectaron ambas secciones")
    exit()

# Limpiar nombres
df_angles.columns = [col.split(' - ')[-1].strip() if ' - ' in col else col.strip() for col in df_angles.columns]

print("Columnas Angles:", list(df_angles.columns))

# Convertir tiempos
df_access['Start'] = pd.to_datetime(df_access.iloc[:,1], errors='coerce')
df_access['Stop']  = pd.to_datetime(df_access.iloc[:,2], errors='coerce')
df_angles['Time']  = pd.to_datetime(df_angles['Time (UTCG)'], errors='coerce')

resultados = []

for i, row in df_access.iterrows():
    start = row['Start']
    stop  = row['Stop']
    
    subset = df_angles[(df_angles['Time'] >= start) & (df_angles['Time'] <= stop)]
    
    if len(subset) == 0:
        continue
        
    best = subset.loc[subset['Range (km)'].idxmin()]
    
    resultados.append({
        'Access': i+1,
        'Target': str(row.iloc[-1]),
        'Start_Time': start,
        'Stop_Time': stop,
        'Duration_sec': row.iloc[3],
        'Best_Time': best['Time'],
        'Along_Track_deg': best.get('Along Track (deg)', None),
        'Cross_Track_deg': best.get('Cross Track (deg)', None),
        'Range_km': best['Range (km)']
    })

df_final = pd.DataFrame(resultados)

print("\n=== RESUMEN ===")
print(df_final[['Target', 'Duration_sec', 'Cross_Track_deg', 'Range_km']].round(3))

df_final.to_excel(output, index=False)
print(f"\n¡Archivo guardado!: {output}")