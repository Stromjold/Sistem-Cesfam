import openpyxl
import tkinter as tk
from tkinter import filedialog
import sys

# Selector de archivo
root = tk.Tk()
root.withdraw()
print("Selecciona tu archivo 'cruce.xlsx'...")
ruta = filedialog.askopenfilename()

if not ruta: sys.exit()

wb = openpyxl.load_workbook(ruta, data_only=True)
ws = wb.active

print("Escanendo archivo en busca de colores ÚNICOS (esto puede tardar unos segundos)...")

colores_encontrados = set()

# Escanear todo el archivo
for fila in ws.iter_rows():
    celda = fila[0] # Miramos la columna A
    color = celda.fill.fgColor.rgb
    
    # Si encontramos un color nuevo, lo mostramos
    if color not in colores_encontrados:
        colores_encontrados.add(color)
        print(f"Nuevo color detectado en Fila {celda.row}: {color}")

print("\n--- RESULTADOS ---")
print("Copia el código que parezca Rojo (generalmente empieza con FF y NO es FFFFFFFF ni FFF0F0F0).")