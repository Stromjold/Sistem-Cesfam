import random
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

def generar_rut():
    """Genera un RUT chileno válido con dígito verificador"""
    numero = random.randint(5000000, 25000000)
    rut_str = str(numero)
    
    suma = 0
    multiplo = 2
    
    for i in range(len(rut_str) - 1, -1, -1):
        suma += int(rut_str[i]) * multiplo
        if multiplo == 7:
            multiplo = 2
        else:
            multiplo += 1
    
    dv = 11 - (suma % 11)
    if dv == 11:
        digito_verificador = '0'
    elif dv == 10:
        digito_verificador = 'K'
    else:
        digito_verificador = str(dv)
    
    rut_formateado = f"{rut_str[:-3]}.{rut_str[-3:]}-{digito_verificador}"
    return rut_formateado


def format_rut(rut_str: str) -> str:
    """
    Formatea un RUT chileno al formato XX.XXX.XXX-X
    Ej: 163456789 -> 16.345.678-9
    Ej: 15811.479-8 -> 15.811.479-8
    """
    if not rut_str or rut_str == '':
        return ''
    
    rut_str = str(rut_str).strip()
    
    # Extraer solo dígitos y K (dígito verificador puede ser K)
    rut_limpio = ''
    for c in rut_str:
        if c.isdigit():
            rut_limpio += c
        elif c.upper() == 'K':
            rut_limpio += c.upper()
    
    if len(rut_limpio) < 2:
        return rut_str
    
    # Separar cuerpo y dígito verificador
    # El dígito verificador está al final (puede ser número o K)
    digito = rut_limpio[-1]
    cuerpo = rut_limpio[:-1]
    
    # Formatear el cuerpo con puntos cada 3 dígitos de derecha a izquierda
    cuerpo_formateado = ''
    for i, digit in enumerate(reversed(cuerpo)):
        if i > 0 and i % 3 == 0:
            cuerpo_formateado = '.' + cuerpo_formateado
        cuerpo_formateado = digit + cuerpo_formateado
    
    return f"{cuerpo_formateado}-{digito}"

def format_dataframe_excel(df: pd.DataFrame, xlsx_path: str):
    """Aplica formato a un archivo Excel: encabezados en negrilla, bordes, celdas nulas en rojo"""
    try:
        # Guardar el DataFrame en Excel primero
        df.to_excel(xlsx_path, index=False, engine='openpyxl')
        
        # Cargar el archivo para aplicar formato
        wb = openpyxl.load_workbook(xlsx_path)
        
        # Validar que la hoja activa existe
        if wb.active is None:
            print(f"  ⚠ No se pudo acceder a la hoja de trabajo")
            wb.close()
            return
        
        ws = wb.active
        
        # Definir bordes
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Formato para la primera fila (encabezado)
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        
        # Formato para datos nulos (rojo)
        null_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        null_font = Font(color="FFFFFF")
        
        # Aplicar formato a todas las celdas
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                # Aplicar bordes
                cell.border = thin_border
                
                # Centro alineado
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                # Formato especial para la primera fila
                if cell.row == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                else:
                    # Aplicar formato rojo para celdas nulas/vacías
                    if cell.value is None or str(cell.value).strip() == '':
                        cell.fill = null_fill
                        cell.font = null_font
        
        # Ajustar ancho de columnas automáticamente
        for col_num in range(1, ws.max_column + 1):
            max_length = 0
            column_letter = get_column_letter(col_num)
            
            for row_num in range(1, ws.max_row + 1):
                cell = ws.cell(row=row_num, column=col_num)
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Guardar con formato aplicado
        wb.save(xlsx_path)
        wb.close()
        print(f"  ✓ Archivo guardado con formato: {xlsx_path}")
    except Exception as e:
        print(f"  ⚠ Error aplicando formato: {e}")


def generar_fecha_nacimiento():
    """Genera una fecha de nacimiento aleatoria"""
    año = random.randint(1945, 2005)
    mes = random.randint(1, 12)
    dia = random.randint(1, 28)
    return f"{dia:02d}/{mes:02d}/{año}"

def generar_usuarios(cantidad, id_inicial=1):
    """Genera una lista de usuarios con datos hipotéticos"""
    
    nombres = ['Juan', 'María', 'Pedro', 'Ana', 'Luis', 'Carmen', 'Diego', 'Sofia', 
               'Carlos', 'Valentina', 'Miguel', 'Francisca', 'José', 'Isabel', 'Andrés', 
               'Catalina', 'Felipe', 'Daniela', 'Jorge', 'Constanza', 'Roberto', 'Javiera', 
               'Sebastián', 'Martina', 'Alejandro', 'Antonia', 'Ricardo', 'Camila', 
               'Fernando', 'Victoria']
    
    apellidos = ['González', 'Rodríguez', 'Muñoz', 'López', 'Pérez', 'García', 
                 'Martínez', 'Fernández', 'Silva', 'Torres', 'Rojas', 'Díaz', 
                 'Soto', 'Contreras', 'Vargas', 'Castro', 'Ramírez', 'Morales', 
                 'Reyes', 'Fuentes', 'Castillo', 'Valdés', 'Hernández', 'Sepúlveda', 
                 'Espinoza', 'Núñez', 'Tapia', 'Gutiérrez', 'Carrasco', 'Vera']
    
    nacionalidades = ['Chilena'] * 6 + ['Argentina', 'Peruana', 'Boliviana', 
                                         'Venezolana', 'Colombiana']
    
    tipos_sangre = ['O+', 'O-', 'A+', 'A-', 'B+', 'B-', 'AB+', 'AB-']
    
    estados_civiles = ['Soltero', 'Casado', 'Divorciado', 'Viudo', 'Separado']
    
    calles = ['Avenida Libertador', 'Calle Providencia', 'Pasaje Los Álamos', 
              'Avenida Apoquindo', 'Calle Moneda', 'Avenida Brasil', 
              'Calle Huérfanos', 'Avenida Las Condes', 'Calle Agustinas', 
              'Avenida Vicuña Mackenna', 'Calle Bandera', 'Paseo Bulnes', 
              'Calle Estado', 'Avenida Grecia', 'Calle San Antonio']
    
    comunas = ['Santiago', 'Providencia', 'Las Condes', 'Ñuñoa', 'La Florida', 
               'Maipú', 'Puente Alto', 'Peñalolén', 'Estación Central', 'Recoleta']
    
    usuarios = []
    
    for i in range(cantidad):
        usuario = {
            'ID': id_inicial + i,
            'Nombre': random.choice(nombres),
            'Apellido': random.choice(apellidos),
            'RUT': generar_rut(),
            'Fecha_Nacimiento': generar_fecha_nacimiento() if random.random() > 0.08 else None,
            'Nacionalidad': random.choice(nacionalidades) if random.random() > 0.08 else None,
            'Tipo_Sangre': random.choice(tipos_sangre) if random.random() > 0.08 else None,
            'Direccion': f"{random.choice(calles)} {random.randint(1000, 9999)}, {random.choice(comunas)}" if random.random() > 0.08 else None,
            'Estado_Civil': random.choice(estados_civiles) if random.random() > 0.08 else None
        }
        usuarios.append(usuario)
    
    return usuarios

def main(output_dir='.'):
    # Ajustes de visualización para terminal
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)

    # Generar usuarios base
    print("Generando usuarios...")

    # Pedir nombres de archivo antes de generar
    nombre_a = input("\nNombre para el archivo A (Enter para 'tabla_a_usuarios.xlsx'): ").strip() or "tabla_a_usuarios.xlsx"
    nombre_b = input("Nombre para el archivo B (Enter para 'tabla_b_usuarios.xlsx'): ").strip() or "tabla_b_usuarios.xlsx"

    if not nombre_a.lower().endswith('.xlsx'):
        nombre_a += '.xlsx'
    if not nombre_b.lower().endswith('.xlsx'):
        nombre_b += '.xlsx'
    
    # Generar 100 usuarios comunes (en ambas tablas)
    usuarios_comunes = generar_usuarios(100, id_inicial=1)
    
    # Generar 80 usuarios solo para tabla A
    usuarios_solo_a = generar_usuarios(80, id_inicial=101)
    
    # Generar 50 usuarios solo para tabla B
    usuarios_solo_b = generar_usuarios(50, id_inicial=181)
    
    # Crear Tabla A: comunes + solo_a
    usuarios_tabla_a = usuarios_comunes + usuarios_solo_a
    
    # Agregar duplicados en Tabla A (duplicar 10 usuarios aleatorios)
    usuarios_duplicados_a = random.sample(usuarios_tabla_a, 10)
    usuarios_tabla_a.extend(usuarios_duplicados_a)
    
    random.shuffle(usuarios_tabla_a)  # Mezclar aleatoriamente
    df_tabla_a = pd.DataFrame(usuarios_tabla_a)
    
    # Crear Tabla B: comunes + solo_b
    usuarios_tabla_b = usuarios_comunes + usuarios_solo_b
    
    # Agregar duplicados en Tabla B (duplicar 8 usuarios aleatorios)
    usuarios_duplicados_b = random.sample(usuarios_tabla_b, 8)
    usuarios_tabla_b.extend(usuarios_duplicados_b)
    
    random.shuffle(usuarios_tabla_b)  # Mezclar aleatoriamente
    df_tabla_b = pd.DataFrame(usuarios_tabla_b)

    # Guardar en archivos Excel
    ruta_a = f"{output_dir}/{nombre_a}"
    ruta_b = f"{output_dir}/{nombre_b}"

    df_tabla_a.to_excel(ruta_a, index=False, engine='openpyxl')
    df_tabla_b.to_excel(ruta_b, index=False, engine='openpyxl')
    
    # Aplicar formato a los archivos
    format_dataframe_excel(df_tabla_a, ruta_a)
    format_dataframe_excel(df_tabla_b, ruta_b)

    print("\n✓ Archivos generados exitosamente:")
    print(f"  - {nombre_a} (190 usuarios con 10 duplicados)")
    print(f"  - {nombre_b} (158 usuarios con 8 duplicados)")
    print(f"  - Usuarios únicos en ambas tablas: 100")
    print(f"  - Usuarios únicos solo en A: 80")
    print(f"  - Usuarios únicos solo en B: 50")

    # Verificar duplicados por RUT
    duplicados_a = df_tabla_a['RUT'].duplicated().sum()
    duplicados_b = df_tabla_b['RUT'].duplicated().sum()
    
    print(f"\n--- Verificación ---")
    print(f"Duplicados en Tabla A: {duplicados_a}")
    print(f"Duplicados en Tabla B: {duplicados_b}")

    # Mostrar primeros registros de cada tabla
    print("\n--- Primeros 5 usuarios de Tabla A ---")
    print(df_tabla_a.head().to_string(index=False))

    print("\n--- Primeros 5 usuarios de Tabla B ---")
    print(df_tabla_b.head().to_string(index=False))

    # Estadísticas
    print("\n--- Estadísticas ---")
    print(f"Total usuarios Tabla A: {len(df_tabla_a)}")
    print(f"Total usuarios Tabla B: {len(df_tabla_b)}")


if __name__ == '__main__':
    main()