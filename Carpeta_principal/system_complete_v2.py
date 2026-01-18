"""
SISTEMA DE CRUCE V2 - ANTÍDOTO DE DECIMALES
--------------------------------------------------
Corrige el error donde 12345.0 no cruzaba con 12345.
Garantiza la búsqueda exacta en Rayen y Fonasa.
"""
import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import sys
import os
import re
import subprocess

# ==========================================
# 0. CONFIGURACIÓN DE RUTAS
# ==========================================
DIRECTORIO_PRINCIPAL = os.path.dirname(os.path.abspath(__file__))
DIR_CRUCE     = os.path.join(DIRECTORIO_PRINCIPAL, "Archivos_Entrada")
DIR_MAESTROS  = os.path.join(DIRECTORIO_PRINCIPAL, "Archivos_escanear")
DIR_RESULTADOS = os.path.join(DIRECTORIO_PRINCIPAL, "Resultados")

for carpeta in [DIR_CRUCE, DIR_MAESTROS, DIR_RESULTADOS]:
    if not os.path.exists(carpeta):
        os.makedirs(carpeta)

print(f"--- SISTEMA DE CRUCE V2 (CORRECCIÓN DE FORMATOS) ---")
print(f"📂 Trabajando en: {DIRECTORIO_PRINCIPAL}")

# ==========================================
# 1. LIMPIEZA INTELIGENTE (EL ANTÍDOTO)
# ==========================================

def limpiar_nombre_archivo(texto):
    if not texto: return "SinTitulo"
    return re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()

def limpiar_rut_robusto(serie):
    """
    Limpieza extrema:
    1. Convierte a string.
    2. Elimina '.0' al final (error común de Excel).
    3. Elimina puntos, guiones y espacios.
    4. Deja solo números y K.
    """
    return (serie.astype(str)
            .str.upper()
            .str.strip()
            .str.replace(r'\.0$', '', regex=True) # Quita decimal .0
            .str.replace(r'[^0-9K]', '', regex=True)) # Deja solo 0-9 y K

def encontrar_columna(df, keywords):
    cols_lower = [str(c).lower().strip() for c in df.columns]
    for key in keywords:
        if key in cols_lower:
            return df.columns[cols_lower.index(key)]
    return None

def seleccionar_archivo(titulo, carpeta_inicial):
    """Selector de archivos con memoria de carpeta"""
    root = tk.Tk()
    root.attributes('-topmost', True) 
    root.withdraw()
    print(f"\n📂 BUSCANDO: {titulo}...")
    
    # Intento de autoselección si solo hay 1 archivo
    archivos = [f for f in os.listdir(carpeta_inicial) if f.endswith(('.xlsx', '.xls', '.csv'))]
    if "Rayen" in titulo and any("Rayen" in f for f in archivos):
        mejor_match = [f for f in archivos if "Rayen" in f][0]
        ruta = os.path.join(carpeta_inicial, mejor_match)
        print(f"   ⚡ Autoseleccionado: {mejor_match}")
        return ruta
        
    ruta = filedialog.askopenfilename(title=titulo, initialdir=carpeta_inicial, filetypes=[("Excel Files", "*.xlsx *.xls *.csv")])
    if ruta: 
        print(f"   ✅ Seleccionado: {os.path.basename(ruta)}")
        return ruta
    return None

# ==========================================
# 2. MOTOR DE BÚSQUEDA
# ==========================================

def buscar_en_maestro(ruta_maestro, lista_ruts_buscar, nombre_entidad):
    if not ruta_maestro: return pd.DataFrame()
    
    print(f"\n   🤖 Procesando: {os.path.basename(ruta_maestro)}...")
    
    try:
        if ruta_maestro.lower().endswith('.csv'):
            df = pd.read_csv(ruta_maestro, dtype=str, encoding='latin-1')
        else:
            df = pd.read_excel(ruta_maestro, dtype=str)
        
        # DETECTAR COLUMNAS
        col_rut = encontrar_columna(df, ['rut', 'run', 'identificador', 'cedula', 'numero tipo identificacion', 'rut fonasa', 'rut rayen'])
        col_dv = encontrar_columna(df, ['dv', 'digito', 'digito verificador'])
        
        if not col_rut:
            col_rut = df.columns[0]
            print(f"      ⚠️ No encontré columna RUT obvia. Usaré: '{col_rut}'")
        else:
            print(f"      📍 Columna RUT: '{col_rut}'")

        # PREPARAR DATOS (LIMPIEZA ROBUSTA)
        set_buscados = set(limpiar_rut_robusto(pd.Series(lista_ruts_buscar)))
        
        # --- ESTRATEGIA 1: UNIÓN RUN + DV ---
        if col_rut and col_dv:
            rut_armado = limpiar_rut_robusto(df[col_rut]) + limpiar_rut_robusto(df[col_dv])
            matches = df[rut_armado.isin(set_buscados)]
            if not matches.empty:
                print(f"      ✅ [RUN+DV] Encontrados: {len(matches)}")
                return matches

        # --- ESTRATEGIA 2: CRUCE DIRECTO ---
        rut_maestro_limpio = limpiar_rut_robusto(df[col_rut])
        matches = df[rut_maestro_limpio.isin(set_buscados)]
        if not matches.empty:
            print(f"      ✅ [DIRECTO] Encontrados: {len(matches)}")
            return matches

        # --- ESTRATEGIA 3: RAÍZ (SIN DV) ---
        # Si falla, probamos quitando el último dígito al maestro (por si tiene guion-k y nosotros no)
        rut_maestro_raiz = rut_maestro_limpio.apply(lambda x: x[:-1] if len(x)>7 else x)
        matches = df[rut_maestro_raiz.isin(set_buscados)]
        if not matches.empty:
            print(f"      ✅ [RAÍZ] Encontrados: {len(matches)}")
            return matches

        print("      ❌ 0 Coincidencias. Revisa los formatos.")
        return pd.DataFrame()
        
    except Exception as e:
        print(f"      ❌ Error técnico: {e}")
        return pd.DataFrame()

# ==========================================
# 3. PROGRAMA PRINCIPAL
# ==========================================

def main():
    # --- PASO 1: EXTRAER RUTS LIMPIOS ---
    print("\n[PASO 1] Leyendo 'cruce.xlsx'...")
    ruta_cruce = seleccionar_archivo("Selecciona CRUCE", DIR_CRUCE)
    if not ruta_cruce: sys.exit()

    app = xw.App(visible=False)
    ruts_A, ruts_B = [], []
    titulo_A, titulo_B = "Columna_A", "Columna_B"

    try:
        wb = app.books.open(ruta_cruce)
        hoja = wb.sheets[0]
        titulo_A = str(hoja.range('A1').value).strip()
        titulo_B = str(hoja.range('B1').value).strip()
        color_rojo = hoja.range('A2').api.DisplayFormat.Interior.Color
        
        last = hoja.range('A' + str(hoja.cells.last_cell.row)).end('up').row
        print(f"      -> Escaneando {last} filas...")

        vals_A = hoja.range(f'A2:A{last}').value
        vals_B = hoja.range(f'B2:B{last}').value
        
        # Iteramos usando rango para velocidad, chequeando color celda por celda
        # (xlwings es lento leyendo colores, paciencia)
        for i in range(2, last + 1):
            if hoja.range(f'A{i}').api.DisplayFormat.Interior.Color != color_rojo:
                val = hoja.range(f'A{i}').value
                if val: ruts_A.append(str(val))
            
            if hoja.range(f'B{i}').api.DisplayFormat.Interior.Color != color_rojo:
                val = hoja.range(f'B{i}').value
                if val: ruts_B.append(str(val))
                
            if i % 1000 == 0: print(f"      Fila {i}...", end='\r')

        print(f"\n      ✅ Listos para buscar: {len(ruts_A)} de {titulo_A} | {len(ruts_B)} de {titulo_B}")
        
    finally:
        try: wb.close(); app.quit()
        except: pass

    # --- PASO 2: BUSCAR INFO ---
    print("\n[PASO 2] Cruzando datos...")
    
    # Rayen
    ruta_A = seleccionar_archivo(f"Base para {titulo_A}", DIR_MAESTROS)
    df_A = buscar_en_maestro(ruta_A, ruts_A, titulo_A)
    
    # Percapita
    ruta_B = seleccionar_archivo(f"Base para {titulo_B}", DIR_MAESTROS)
    df_B = buscar_en_maestro(ruta_B, ruts_B, titulo_B)

    # --- PASO 3: GUARDAR ---
    archivo_final = os.path.join(DIR_RESULTADOS, "Reporte_Final_Completo.xlsx")
    
    with pd.ExcelWriter(archivo_final) as writer:
        if not df_A.empty: df_A.to_excel(writer, sheet_name=limpiar_nombre_archivo(titulo_A)[:30], index=False)
        if not df_B.empty: df_B.to_excel(writer, sheet_name=limpiar_nombre_archivo(titulo_B)[:30], index=False)
            
    print(f"\n🏆 ¡REPORTE LISTO!\n    📂 {archivo_final}")
    print(f"    Resumen: {len(df_A)} encontrados en Rayen | {len(df_B)} encontrados en Fonasa")
    
    # --- PREGUNTAR SI DESEA ABRIR EL ARCHIVO ---
    print("\n" + "="*50)
    respuesta = input("¿Deseas abrir el reporte ahora? (Y/N): ").strip().upper()
    
    if respuesta == 'Y':
        try:
            if os.name == 'nt':  # Windows
                os.startfile(archivo_final)
            elif os.name == 'posix':  # macOS/Linux
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', archivo_final])
            print("    ✅ Abriendo archivo...")
        except Exception as e:
            print(f"    ❌ Error al abrir el archivo: {e}")
    else:
        print("    👋 ¡Hasta luego!")

if __name__ == "__main__":
    main()

