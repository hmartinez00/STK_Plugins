import pandas as pd
import io
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm

# ================== SELECTOR DE ARCHIVO ==================
print("Abriendo selector de archivos...")

root = Tk()
root.withdraw()  # Ocultar ventana principal
root.attributes('-topmost', True)  # Poner al frente

archivo = askopenfilename(
    title="Selecciona el archivo CSV de simulaciones",
    filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
)

root.destroy()

if not archivo:
    print("No se seleccionó ningún archivo. Saliendo...")
    exit()

print(f"Archivo seleccionado: {archivo}\n")

# ================== PROCESAMIENTO ==================
output = f"Resumen_Access_Mejor_Roll_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

print("Leyendo archivo...")

# Intentar diferentes encodings
encodings = ['cp1252', 'utf-8', 'utf-8-sig', 'latin-1']
texto = None

for enc in encodings:
    try:
        with open(archivo, 'r', encoding=enc) as f:
            texto = f.read()
        print(f"✓ Archivo leído con encoding: {enc}")
        break
    except UnicodeDecodeError:
        continue

if texto is None:
    raise ValueError("No se pudo leer el archivo con ningún encoding")

bloques = [b.strip() for b in texto.split('\n\n') if b.strip()]

df_access_list = []
df_angles_list = []

for bloque in bloques:
    if "Start Time (UTCG)" in bloque and "Duration (sec)" in bloque:
        df_access_list.append(pd.read_csv(io.StringIO(bloque), on_bad_lines='skip'))
    elif "Time (UTCG)" in bloque and "Along Track" in bloque:
        df_angles_list.append(pd.read_csv(io.StringIO(bloque), on_bad_lines='skip'))

print(f"Se encontraron {len(df_access_list)} bloques Access y {len(df_angles_list)} bloques Angles\n")

resultados = []

for df_access in df_access_list:
    df_access.columns = [col.strip() for col in df_access.columns]
    
    for i, row in tqdm(df_access.iterrows(), total=len(df_access), desc="Procesando accesos"):
        start = pd.to_datetime(row.iloc[1], errors='coerce')
        stop  = pd.to_datetime(row.iloc[2], errors='coerce')
        target_name = str(row.iloc[-1])
        
        for df_angles in df_angles_list:
            df_angles.columns = [col.split(' - ')[-1].strip() if ' - ' in col else col.strip() for col in df_angles.columns]
            
            time_col = [c for c in df_angles.columns if "Time" in c][0]
            df_angles['Time'] = pd.to_datetime(df_angles[time_col], errors='coerce')
            
            subset = df_angles[(df_angles['Time'] >= start) & (df_angles['Time'] <= stop)]
            
            if len(subset) > 5:   # mínimo de puntos para considerar válido
                best = subset.loc[subset['Range (km)'].idxmin()]
                
                resultados.append({
                    'Access': i+1,
                    'Target': target_name,
                    'Start_Time': start,
                    'Stop_Time': stop,
                    'Duration_sec': row.iloc[3],
                    'Best_Time': best['Time'],
                    'Along_Track_deg': best.get('Along Track (deg)', None),
                    'Cross_Track_deg': best.get('Cross Track (deg)', None),
                    'Range_km': best['Range (km)']
                })
                break

df_final = pd.DataFrame(resultados)

print(f"\nTotal de accesos procesados: {len(df_final)}")
if not df_final.empty:
    print(df_final.groupby('Target').size())

df_final.to_excel(output, index=False)
print(f"\n¡Archivo guardado!: {output}")