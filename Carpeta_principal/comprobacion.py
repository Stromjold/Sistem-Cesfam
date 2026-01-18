import pandas as pd
import os
import sys

# --- 1. CONFIGURACIÓN ROBUSTA DE RUTAS ---
# Detectamos dónde está este archivo script para construir las rutas desde ahí
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RUTA_CRUCE = os.path.join(BASE_DIR, "Archivos_Entrada", "cruce.xlsx")
RUTA_REPORTE = os.path.join(BASE_DIR, "Resultados", "Reporte_Final_Completo.xlsx")

print(f"--- INICIANDO COMPROBACIÓN ---")
print(f"📂 Directorio base detectado: {BASE_DIR}")

# --- 2. VERIFICACIÓN DE EXISTENCIA DE ARCHIVOS ---
if not os.path.exists(RUTA_CRUCE):
    print(f"\n❌ ERROR CRÍTICO: No encuentro el archivo de entrada.")
    print(f"   Buscando en: {RUTA_CRUCE}")
    print("   -> Asegúrate de que 'cruce.xlsx' esté dentro de la carpeta 'Archivos_Entrada'")
    sys.exit()

if not os.path.exists(RUTA_REPORTE):
    print(f"\n❌ ERROR CRÍTICO: No encuentro el reporte final.")
    print(f"   Buscando en: {RUTA_REPORTE}")
    print("   -> Asegúrate de haber ejecutado system_complete.py primero.")
    sys.exit()

# --- 3. CARGA DE DATOS ---
print("\n⏳ Leyendo archivos (esto puede tardar unos segundos)...")
try:
    df_cruce = pd.read_excel(RUTA_CRUCE)
    # Leemos todas las hojas del reporte para buscar por nombre
    xls_reporte = pd.ExcelFile(RUTA_REPORTE)
    dict_reporte = pd.read_excel(xls_reporte, sheet_name=None)
    print("   ✅ Archivos cargados correctamente.")
except Exception as e:
    print(f"   ❌ Error leyendo Excel: {e}")
    sys.exit()

# --- 4. FUNCIÓN DE LIMPIEZA ---
def limpiar(serie):
    """Deja solo números y K mayúscula para comparar peras con peras"""
    return serie.astype(str).str.upper().str.strip().str.replace(r'[^0-9K]', '', regex=True)

# --- 5. AUDITORÍA FONASA (PERCAPITA) ---
print(f"\n{'='*40}")
print("🔍 AUDITORÍA FONASA (PERCAPITA)")
print(f"{'='*40}")

try:
    # 1. Qué buscábamos (Columna B de Cruce)
    # Buscamos la columna, a veces se llama 'Rut Fonasa' o 'Fonasa (Ok)'
    col_fonasa_origen = [c for c in df_cruce.columns if 'fonasa' in c.lower()][0]
    ruts_buscados = set(limpiar(df_cruce[col_fonasa_origen].dropna()))
    
    # 2. Qué encontramos (Buscamos la hoja que tenga 'Fonasa' o 'Percapita' en el nombre)
    nombre_hoja_fonasa = [h for h in dict_reporte.keys() if 'fonasa' in h.lower() or 'percapita' in h.lower()][0]
    df_res_fonasa = dict_reporte[nombre_hoja_fonasa]
    
    # Unimos RUN y DV si existen, si no buscamos columna RUT
    if 'RUN' in df_res_fonasa.columns and 'DV' in df_res_fonasa.columns:
        ruts_encontrados = set(limpiar(df_res_fonasa['RUN']) + limpiar(df_res_fonasa['DV']))
    else:
        # Fallback por si la estructura es distinta
        col_res = df_res_fonasa.columns[0]
        ruts_encontrados = set(limpiar(df_res_fonasa[col_res]))

    coincidencias = ruts_buscados.intersection(ruts_encontrados)
    
    print(f"🔹 RUTs Solicitados (Limpios): {len(ruts_buscados)}")
    print(f"🔹 RUTs en Reporte Final:      {len(ruts_encontrados)}")
    print(f"✅ COINCIDENCIAS REALES:       {len(coincidencias)}")
    
    if len(coincidencias) == len(ruts_buscados):
        print("🏆 RESULTADO: ÉXITO TOTAL (100% encontrados)")
    else:
        print(f"⚠️ RESULTADO: Faltaron {len(ruts_buscados) - len(coincidencias)} registros.")

except Exception as e:
    print(f"⚠️ No se pudo auditar Fonasa: {e}")


# --- 6. AUDITORÍA RAYEN ---
print(f"\n{'='*40}")
print("🔍 AUDITORÍA RAYEN")
print(f"{'='*40}")

try:
    # 1. Qué buscábamos (Columna A de Cruce)
    col_rayen_origen = [c for c in df_cruce.columns if 'rayen' in c.lower()][0]
    ruts_buscados = set(limpiar(df_cruce[col_rayen_origen].dropna()))
    
    # 2. Qué encontramos
    nombre_hoja_rayen = [h for h in dict_reporte.keys() if 'rayen' in h.lower()][0]
    df_res_rayen = dict_reporte[nombre_hoja_rayen]
    
    # Buscamos la columna del RUT en el resultado
    col_rut_res = [c for c in df_res_rayen.columns if 'ident' in c.lower() or 'rut' in c.lower()][0]
    ruts_encontrados = set(limpiar(df_res_rayen[col_rut_res]))

    coincidencias = ruts_buscados.intersection(ruts_encontrados)
    
    print(f"🔹 RUTs Solicitados (Limpios): {len(ruts_buscados)}")
    print(f"🔹 RUTs en Reporte Final:      {len(ruts_encontrados)}")
    print(f"✅ COINCIDENCIAS REALES:       {len(coincidencias)}")

    
    porcentaje = (len(coincidencias) / len(ruts_buscados)) * 100
    print(f"📊 EFECTIVIDAD: {porcentaje:.1f}%")

except Exception as e:
    print(f"⚠️ No se pudo auditar Rayen: {e}")

input("\nPresiona Enter para cerrar...")