import openpyxl
import tkinter as tk
from tkinter import filedialog

# Seleccionar archivo
root = tk.Tk()
root.withdraw()
ruta = filedialog.askopenfilename(title="Selecciona tu Excel para inspeccionar")

if not ruta:
    print("No seleccionaste nada.")
else:
    wb = openpyxl.load_workbook(ruta, data_only=True)
    ws = wb.active

    print(f"\n--- INSPECCIONANDO LAS PRIMERAS 10 FILAS DE: {ruta} ---")
    print("Revisando la columna A (primera celda de cada fila)...")
    
    for i, row in enumerate(ws.iter_rows(max_row=10)):
        celda = row[0] # Mira la primera celda
        color = celda.fill.fgColor.rgb
        valor = celda.value
        
        # Traducir el resultado para que sea legible
        if color == "00000000" or color is None:
            estado_color = "TRANSPARENTE / SIN COLOR"
        else:
            estado_color = f"CÓDIGO: {color}"
            
        print(f"Fila {i+1}: Valor='{valor}' -> Python ve: {estado_color}")

    print("\n--- DIAGNÓSTICO ---")
    print("1. Si ves 'TRANSPARENTE' en filas que tú ves rojas -> Es Formato Condicional.")
    print("2. Si ves un código raro (ej: FFFFCCC) -> Ese es el código que debes copiar para el script anterior.")