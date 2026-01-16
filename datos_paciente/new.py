import xlwings as xw
import pandas as pd
import sys

# --- CONFIGURACIÓN ---
ruta_archivo = r'C:\Users\Cirugia Menor\Downloads\cruce.xlsx'
archivo_rojos = 'Pacientes_Duplicados.xlsx'
archivo_limpios = 'Pacientes_Unicos.xlsx'

print("--- INICIANDO ESCÁNER CON RAYOS X (DisplayFormat) ---")
print("Abriendo Excel... (No toques el mouse)")

app = xw.App(visible=False) 

try:
    wb = app.books.open(ruta_archivo)
    hoja = wb.sheets[0]
    
    # 1. OBTENER EL COLOR REAL (LO QUE VEN TUS OJOS)
    # Usamos .api.DisplayFormat.Interior.Color para ver el formato condicional
    # La celda A2 sabemos que es roja.
    color_referencia = hoja.range('A2').api.DisplayFormat.Interior.Color
    
    print(f"\n[APRENDIZAJE] He mirado la celda A2.")
    print(f"El código interno del color detectado es: {color_referencia}")
    
    # Verificación: El blanco puro es 16777215. El gris es 15790320.
    if color_referencia == 16777215:
        print("ERROR: A2 se ve BLANCA. Asegúrate de que A2 sea un caso rojo.")
        sys.exit()
    elif color_referencia == 15790320: # Código del gris (240,240,240)
        print("ADVERTENCIA: He detectado GRIS, no ROJO. ¿Seguro que A2 es roja?")
        # Continuamos por si acaso tu gris es lo que quieres separar, 
        # pero es probable que debas cambiar la celda de referencia a una roja real.

    # 2. ESCANEO
    ultima_fila = hoja.range('A' + str(hoja.cells.last_cell.row)).end('up').row
    print(f"Procesando {ultima_fila} filas... (Paciencia)")

    datos_rojos = []
    datos_blancos = []

    # Iteramos fila por fila
    for i in range(1, ultima_fila + 1):
        celda = hoja.range(f'A{i}')
        
        # Leemos el color VISUAL (DisplayFormat)
        color_actual = celda.api.DisplayFormat.Interior.Color
        
        # Leemos los valores asegurando que sea una lista (ndim=1 evita el error de float)
        valores = celda.expand('right').options(ndim=1).value
        
        if color_actual == color_referencia:
            datos_rojos.append(valores)
        else:
            datos_blancos.append(valores)
            
        if i % 500 == 0:
            print(f"Escaneando fila {i}/{ultima_fila}...", end='\r')

    print(f"\n\n--- RESUMEN FINAL ---")
    print(f"Filas ROJAS (Duplicados): {len(datos_rojos)}")
    print(f"Filas NORMALES (Únicos): {len(datos_blancos)}")

    # 3. GUARDAR (Con manejo de errores para evitar cierres abruptos)
    if datos_rojos:
        df = pd.DataFrame(datos_rojos)
        df.to_excel(archivo_rojos, index=False, header=False)
        print(f"Guardado: {archivo_rojos}")
    
    if datos_blancos:
        df = pd.DataFrame(datos_blancos)
        df.to_excel(archivo_limpios, index=False, header=False)
        print(f"Guardado: {archivo_limpios}")

except Exception as e:
    print(f"\n!!! ERROR FATAL: {e}")
    import traceback
    traceback.print_exc()

finally:
    try:
        wb.close()
        app.quit()
    except:
        pass
    print("Excel cerrado.")