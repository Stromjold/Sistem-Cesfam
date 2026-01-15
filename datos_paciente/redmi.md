# 📊 Sistema de Comparación y Análisis de Datos - Documentación Técnica

Sistema avanzado para comparar archivos Excel/CSV, detectar duplicados, registros faltantes e incompletos, con generación automática de reportes.

---

## 📋 Contenido

- [Características](#-características)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Funcionamiento Detallado](#-funcionamiento-detallado)
- [Tipos de Análisis](#-tipos-de-análisis)
- [Estructura de Reportes](#-estructura-de-reportes)
- [Uso](#-uso)
- [Ejemplos](#-ejemplos)

---

## ✨ Características

### 🔍 Comparador de Archivos (`separador_datos.py`)

- **Comparación inteligente** de archivos Excel (.xlsx, .xls) y CSV
- **Detección automática** de columnas clave (RUT, ID, documento)
- **Normalización de datos**: Ignora mayúsculas/minúsculas y corrige formatos numéricos (ej: 12345.0 -> 12345)
- **Lectura inteligente**: Detecta automáticamente encabezados aunque el archivo tenga títulos o filas vacías al inicio
- **Análisis selectivo**: Duplicados, Faltantes, Incompletos o Todos
- **Detección de duplicados por RUT** con estadísticas detalladas
- **Interfaz de menús interactivos** en terminal
- **Reportes Excel organizados** por tipo de análisis
- **Soporte multi-hoja** y múltiples archivos
- **Optimización para grandes volúmenes** (>8MB)
- **Formato visual mejorado** con colores y tablas en terminal
- **Guardado robusto**: Sistema "anti-bloqueo" que genera copias automáticas (con timestamp) si el archivo de reporte está abierto en Excel
- **Estadísticas de Precisión**: Cálculo exacto de porcentajes de pérdida y coincidencia entre bases de datos

---

## 📦 Requisitos

### Software Necesario

- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Dependencias

```bash
pandas >= 1.3.0
openpyxl >= 3.0.0
tkinter (incluido en Python estándar)
```

---

## 🔧 Funcionamiento Detallado

### 1️⃣ **Flujo Principal del Programa**

```
INICIO
  ↓
[Menú Principal]
  ├─ 1. Comparar archivos → [Selección de archivos]
  ├─ 2. Modo batch          ↓
  └─ 3. Salir          [Menú de análisis]
                             ↓
                  ¿Qué quieres hacer?
                    ├─ 1. Duplicados
                    ├─ 2. Faltantes
                    ├─ 3. Incompletos
                    └─ 4. Todos
                             ↓
                    [Selección de hojas]
                             ↓
                       [ANÁLISIS]
                             ↓
                    [Generación Excel]
                             ↓
                     ¿Abrir archivo?
                             ↓
                          FIN
```

### 2️⃣ **Carga y Detección Automática**

#### **Lectura Inteligente de Tablas**
- **Salto de Títulos**: Si el archivo Excel tiene títulos decorativos o filas vacías al inicio, el sistema analiza la "densidad de datos" de las primeras 20 filas para encontrar automáticamente dónde comienzan los encabezados reales.

#### **Detección de Columna Clave**
```python
Prioridad de búsqueda ampliada:
1. RUT, RUN, ID, DOCUMENTO, CEDULA, FICHA, FOLIO, CASO, N_SOLICITUD
2. Columnas con >80% valores únicos
3. Detección automática por tipo de dato
```

#### **Normalización de Datos (Advanced Cleaning)**
```python
Proceso de limpieza profunda:
1. Conversión a texto y Mayúsculas (ignora case sensitivity)
2. Eliminación de espacios (trim)
3. Corrección de decimales flotantes: "12345.0" → "12345"
4. Generación de hash interno comparison-safe
```

### 3️⃣ **Algoritmos de Análisis**

#### **A. Detección de Faltantes**

**Lógica:**
- `faltantes_en_B` = Registros en A que NO están en B
- `faltantes_en_A` = Registros en B que NO están en A

**Implementación:**
```python
set_a = set(df_a['__KEY__'].unique())
set_b = set(df_b['__KEY__'].unique())

faltantes_en_b = df_a[~df_a['__KEY__'].isin(set_b)]  # En A, no en B
faltantes_en_a = df_b[~df_b['__KEY__'].isin(set_a)]  # En B, no en A
```

**Ejemplo:**
```
Archivo A (Percapita): RUTs [1, 2, 3, 4, 5]
Archivo B (Rayen):     RUTs [3, 4, 5, 6, 7]

Faltantes en B (Rayen):     [1, 2]  → Están en Percapita, faltan en Rayen
Faltantes en A (Percapita): [6, 7]  → Están en Rayen, faltan en Percapita
TODOS los faltantes:        [1, 2, 6, 7]
```

#### **B. Detección de Duplicados**

**Lógica:**
- Busca RUTs que aparecen más de una vez en el MISMO archivo
- Ordena por RUT para agrupar duplicados

**Implementación:**
```python
duplicados_a = df_a[df_a[key_a].duplicated(keep=False)].sort_values(key_a)
duplicados_b = df_b[df_b[key_b].duplicated(keep=False)].sort_values(key_b)
```

**Ejemplo:**
```
Archivo A tiene:
  RUT 12345678: 3 registros
  RUT 23456789: 2 registros
  RUT 34567890: 1 registro  ← No es duplicado

Duplicados detectados: 5 registros (2 RUTs únicos)
Top RUTs duplicados:
  • 12.345.678-9: 3 registros
  • 23.456.789-0: 2 registros
```

#### **C. Detección de Incompletos**

**Lógica:**
- Registros con al menos un campo vacío/nulo
- Se excluyen columnas especiales (__KEY__, RUT)

**Implementación:**
```python
def mark_incomplete(df, exclude_cols):
    campos_evaluar = [c for c in df.columns if c not in exclude_cols]
    mask_incomplete = df[campos_evaluar].isnull().any(axis=1)
    return df[mask_incomplete]
```

### 4️⃣ **Generación de Reportes Excel**

#### **Estructura de Archivos Generados**

**Según análisis seleccionado:**
- `REPORTE_DUPLICADOS.xlsx` (si solo Duplicados)
- `REPORTE_FALTANTES.xlsx` (si solo Faltantes)
- `REPORTE_INCOMPLETOS.xlsx` (si solo Incompletos)
- `REPORTE_COMPLETO_COMPARACION.xlsx` (si Todos)

#### **Estructura Interna de Hojas**

**Para cada tipo de análisis:**
```
Hoja 1: TODOS - [Tipo]
  └─ Consolidado de ambos archivos

Hoja 2: [Tipo] en [Archivo A]
  └─ Solo datos del primer archivo

Hoja 3: [Tipo] en [Archivo B]
  └─ Solo datos del segundo archivo
```

**Ejemplo para Faltantes:**
```
📊 REPORTE_FALTANTES.xlsx
  ├─ TODOS - Faltantes (32,616 registros)
  │   └─ Todos los registros que faltan en algún archivo
  │
  ├─ Faltantes en Rayen (16,076 registros)
  │   └─ Registros que están en Percapita pero NO en Rayen
  │
  └─ Faltantes en Percapita (16,540 registros)
      └─ Registros que están en Rayen pero NO en Percapita
```

#### **Formato Visual**

**Encabezados:**
- Fondo azul (#366092)
- Texto blanco en negrita
- Bordes delgados

**Datos:**
- Celdas nulas/vacías: Fondo rojo con "-"
- RUTs formateados: XX.XXX.XXX-X
- Ajuste automático de ancho (máx 50 caracteres)
- Alineación centrada

### 5️⃣ **Optimizaciones**

#### **Grandes Volúmenes (>8MB)**
```python
- Uso de sets para comparaciones O(1)
- Lectura por chunks para CSVs grandes
- Tipos de datos category para columnas repetitivas
- Procesamiento vectorizado con pandas
```

#### **Memoria**
```python
- Advertencias si uso >500MB
- Liberación automática de DataFrames temporales
- Copia eficiente con .copy() solo cuando necesario
```

### 6️⃣ **Visualización en Terminal**

#### **Tablas Formateadas**
```
╔════════════════════════════════════════╗
║  RUT         │ NOMBRE      │ EDAD      ║
╠════════════════════════════════════════╣
║  12.345.67.. │ Juan Pérez  │ 35        ║
║  23.456.78.. │ María Lóp.. │ 28        ║
╚════════════════════════════════════════╝
```

**Características:**
- Trunca columnas largas con "..."
- Ancho máximo 50 caracteres por columna
- Permite scroll horizontal
- RUTs formateados automáticamente

#### **Barra de Progreso y Estadísticas**
```
Guardando: [████████████████████████████████████████████████████████████████████████████████] 100.0%
```

**Nuevo Visualizador de Precisión:**
Al finalizar, verás un resumen exacto del cruce de datos:

```text
📊 ESTADÍSTICAS DE FALTANTES (PRECISION):
   ❌ FALTAN EN BASE_B: 500 usuarios
      (Representa el 3.02% de los datos originales de BASE_A)
   ❌ FALTAN EN BASE_A: 37 usuarios
      (Representa el 0.23% de los datos originales de BASE_B)
   ✅ REGISTROS COMUNES: 16,040
      (Presentes en ambos archivos)
```

---

## 🎯 Tipos de Análisis

### 1. **DUPLICADOS**
- **¿Qué detecta?** RUTs que aparecen múltiples veces en el MISMO archivo
- **Usa esta opción para:** Limpiar bases de datos con registros repetidos
- **Salida:** Lista de todos los registros duplicados agrupados por RUT

### 2. **FALTANTES**
- **¿Qué detecta?** Registros que están en un archivo pero no en el otro
- **Usa esta opción para:** Sincronizar dos bases de datos
- **Salida:** Registros faltantes separados por archivo origen

### 3. **INCOMPLETOS**
- **¿Qué detecta?** Registros con campos vacíos o nulos
- **Usa esta opción para:** Validar completitud de datos
- **Salida:** Registros con al menos un campo vacío

### 4. **TODOS**
- **¿Qué incluye?** Los tres análisis anteriores
- **Usa esta opción para:** Análisis completo de calidad de datos
- **Salida:** Archivo con todas las categorías separadas por hojas

---

## 📊 Estructura de Reportes

### Contenido de Cada Hoja

**Columnas Incluidas:**
- ✅ TODAS las columnas originales del archivo fuente
- ✅ Valores formateados (RUT con puntos y guión)
- ✅ Celdas vacías resaltadas en rojo
- ❌ NO se agregan columnas sintéticas (como "Origen")

**Orden de Datos:**
- Duplicados: Ordenados por RUT
- Faltantes: Orden original del archivo
- Incompletos: Orden original del archivo

### Información de Debug (en terminal)

Durante la ejecución verás:
```
🔍 DEBUG - Análisis seleccionados: ['duplicados']
🔍 DEBUG - Claves en reportes_dict:
    - TODOS - Duplicados: 300 filas
    - Duplicados en A: 150 filas
    - Duplicados en B: 150 filas

✓ Creada hoja: TODOS - Duplicados (300 filas)
✓ Creada hoja: Duplicados en Percapita (150 filas)
✓ Creada hoja: Duplicados en Rayen (150 filas)
```

---

## 🚀 Instalación

### Paso 1: Clonar o Descargar

```bash
# Opción 1: Clonar repositorio
git clone [URL_DEL_REPOSITORIO]
cd datos_paciente

# Opción 2: Descargar ZIP y extraer
```

### Paso 2: Instalar Dependencias

```bash
pip install pandas openpyxl
```

### Paso 3: Verificar Instalación

```bash
python separador_datos.py
```

---

## 📖 Uso

### Ejecución Básica

```bash
python separador_datos.py
```

### Flujo de Trabajo

#### **1. Menú Principal**
```
╔════════════════════════════════════════════════╗
║     🔍 COMPARADOR DE ARCHIVOS EXCEL/CSV       ║
╚════════════════════════════════════════════════╝

  1. 📊 Comparar dos archivos
  2. 📁 Modo batch (múltiples archivos)
  3. ❌ Salir

Escribe tu opción (1, 2 o 3):
```

#### **2. Selección de Archivos**
- Se abre ventana de diálogo del sistema
- Puedes seleccionar múltiples archivos (Ctrl+Click)
- Formatos soportados: .xlsx, .xls, .csv

#### **3. Menú de Tipo de Análisis**
```
❓ ¿QUÉ QUIERES HACER?
══════════════════════════════════════════════════
  1. Duplicados
  2. Faltantes
  3. Incompletos
  4. Todos los anteriores
══════════════════════════════════════════════════

Escribe tu opción (1, 2, 3 o 4):
```

#### **4. Selección de Hojas** (si es Excel)
```
📋 Hojas disponibles en Percapita.xlsx:
  1. Hoja1
  2. Datos
  3. Resumen

Selecciona hoja para Percapita (número o 'ALL'): 1
```

#### **5. Procesamiento**
Verás información detallada en tiempo real del análisis completo.

#### **6. Resultados**
- Archivo Excel generado en el mismo directorio
- Opción de abrir automáticamente
- Volver al menú principal o salir

---

## 💡 Ejemplos Prácticos

### Ejemplo 1: Detectar Duplicados

**Caso de uso:** Verificar si hay RUTs repetidos en un archivo.

**Resultado esperado:**
```
REPORTE_DUPLICADOS.xlsx
  ├─ TODOS - Duplicados
  └─ Duplicados en [Archivo]
      • RUT 12.345.678-9: 3 registros
      • RUT 23.456.789-0: 2 registros
```

### Ejemplo 2: Sincronizar Bases de Datos

**Interpretación de resultados:**
- **"Faltantes en BaseDatos2"**: Registros que debes agregar a BaseDatos2
- **"Faltantes en BaseDatos1"**: Registros que debes agregar a BaseDatos1

### Ejemplo 3: Análisis Completo

Genera un reporte con 9 hojas organizadas por categoría.

---

## ❌ Solución de Problemas

### "No hay datos para generar el reporte"
- Verifica que haya diferencias entre los archivos
- Prueba con otro tipo de análisis

### "No se pudo guardar el archivo" / "Permission Denied"
**Solución Automática:**
El programa detecta si tienes el archivo Excel abierto.
- **NO se detendrá** ni mostrará error.
- Guardará automáticamente una copia con la hora actual (ej: `REPORTE_FALTANTES_17025.xlsx`)
- Te avisará el nombre del nuevo archivo generado.

---

## 📧 Soporte

Para reportar problemas o sugerencias, revisa los mensajes de debug en la terminal que proporcionan información detallada sobre el procesamiento.

---

## 📝 Notas Finales

- El programa preserva TODAS las columnas originales
- Los RUTs se formatean automáticamente
- Las celdas vacías se resaltan en rojo en Excel
- El análisis es optimizado para archivos grandes
```

### 2. Instalar dependencias

```bash
pip install pandas openpyxl
```

### 3. Verificar instalación

```bash
python separador_datos.py
```

---

## 💻 Uso

### Comparar Archivos

#### Modo Interactivo (2 archivos)

```bash
python separador_datos.py
```

Luego selecciona **opción 1** en el menú.

**Pasos:**
1. Se abre ventana para seleccionar **los archivos a utilizar**
3. (Opcional) Selecciona hoja si es Excel multi-hoja
4. Automáticamente detecta la columna clave (RUT, ID, etc.)
5. Genera reporte consolidado y lo guarda en el mismo directorio

**Reporte generado:**
- `REPORTE_COMPLETO_COMPARACION.xlsx` con hojas:
  - **FALTANTES**: tabla A y B lado a lado con nombres reales (ej: Catemu, Chagres)
  - **DUPLICADOS**: tabla A y B lado a lado
  - **INCOMPLETOS**: tabla A y B lado a lado
  - **TODOS - Faltantes/Duplicados/Incompletos**: tablas consolidadas
  - **Usuarios Faltantes A/B**: análisis de usuarios con datos nulos

_Nota: Los títulos de las hojas usan nombres reales de archivos automáticamente_

#### Modo Múltiple (3+ archivos)

```bash
python separador_datos.py
```

Selecciona **opción 2** en el menú para comparar múltiples archivos contra uno base.

---

## 📁 Estructura del Proyecto

```
datos_paciente/
│
├── separador_datos.py                      # Script principal de comparación
├── redmi.md                                # Documentación del proyecto
│
├── [tus_archivos].xlsx                     # Archivos Excel a comparar
│
├── REPORTE_COMPLETO_COMPARACION.xlsx       # 📊 Reporte consolidado generado
│   │
│   ├── FALTANTES                          # Registros presentes en un archivo pero no en el otro
│   ├── DUPLICADOS                         # Registros con RUT/ID duplicado dentro de cada archivo
│   ├── INCOMPLETOS                        # Registros con campos vacíos o nulos
│   │
│   ├── TODOS - Faltantes                  # Consolidado global de faltantes
│   ├── TODOS - Duplicados                 # Consolidado global de duplicados
│   ├── TODOS - Incompletos                # Consolidado global de incompletos
│   │
│   ├── Usuarios Faltantes [Nombre A]      # Análisis por usuario: campos nulos en archivo A
│   └── Usuarios Faltantes [Nombre B]      # Análisis por usuario: campos nulos en archivo B
│
└── __pycache__/                            # Cache de Python (auto-generado)

---

## 📊 Ejemplos

### Ejemplo 1: Comparar Datos

```bash
python separador_datos.py
# Seleccionar opción 1
# Elegir los archivos autilizar **Limite de archivos**
```

### Ejemplo 2: Analizar Resultados

Después de ejecutar la comparación, abre el archivo generado:

**`REPORTE_COMPLETO_COMPARACION.xlsx`**

Encontrarás hojas organizadas con:
- **FALTANTES**: Registros únicos de cada archivo lado a lado
- **DUPLICADOS**: Registros con RUT/ID repetido
- **INCOMPLETOS**: Registros con campos nulos
- **TODOS - [Categoría]**: Consolidados globales
- **Usuarios Faltantes [Archivo]**: Análisis detallado por usuario

_Los títulos usan nombres reales: "Faltantes en (Nombres_archivo)", "Duplicados en (Nombres_archivo)", etc._

### Ejemplo 3: Comparación Múltiple

```bash
python separador_datos.py
# Seleccionar opción 2
# Elegir múltiples archivos (3 o más)
# El sistema compara todos contra el primero seleccionado
```

---

## 🔑 Detección Automática de Columnas

El sistema detecta automáticamente columnas clave buscando estos nombres:

- `id_rut`, `rut`, `RUT`
- `id`, `ID`, `id_usuario`, `usuario_id`
- `documento`, `doc`, `cedula`

Si no encuentra ninguna, usa la primera columna con mayor unicidad.

---

## 📈 Características Avanzadas

### Análisis de Unicidad

El comparador evalúa la calidad de las columnas clave:
- % de valores únicos
- Cantidad de duplicados
- Valores nulos

### Manejo de Archivos Grandes

- Lectura eficiente con `pandas`
- Procesamiento por chunks cuando es necesario
- Modo solo lectura para Excel

### Formato de Salida

Todos los reportes incluyen:
- Formateo automático en Excel: encabezados azules, bordes, nulos en rojo
- Títulos dinámicos con nombres reales de archivos
- Análisis de nulidades y duplicados por columna

#### Reporte único consolidado

Al finalizar la comparación se genera el archivo `REPORTE_COMPLETO_COMPARACION.xlsx` en el mismo directorio, con las hojas:

**Comparativas (lado a lado):**
- **FALTANTES**: Registros únicos en cada archivo
- **DUPLICADOS**: Registros duplicados detectados
- **INCOMPLETOS**: Registros con datos nulos

**Consolidadas:**
- **TODOS - Faltantes**: Todos los registros faltantes juntos
- **TODOS - Duplicados**: Todos los duplicados juntos
- **TODOS - Incompletos**: Todos los incompletos juntos

**Análisis de usuarios:**
- **Usuarios Faltantes A/B**: Detalle de usuarios con campos nulos (ordenados por cantidad)

_Los títulos de las tablas muestran nombres reales: "Faltantes en Catemu", "Duplicados en Chagres", etc._

---

## ⚙️ Configuración

### Agregar nombres de columnas clave

En `separador_datos.py`, línea 13:

```python
COMMON_KEY_NAMES = ['id_rut', 'rut', 'RUT', 'id', 'id_usuario', 'usuario_id', 'ID', 'documento', 'doc', 'cedula']
```

Añade tus propios nombres de columnas identificadoras.

---

## 🐛 Solución de Problemas

### Error: "No module named 'openpyxl'"

```bash
pip install openpyxl
```

### Error: "No module named 'tkinter'"

**Windows/macOS:** Ya viene instalado con Python

**Linux:**
```bash
sudo apt-get install python3-tk
```

### Los archivos no aparecen en la ventana

- Verifica que estés en el directorio correcto
- Asegúrate que los archivos tengan extensión `.xlsx`, `.xls` o `.csv`

### Detección incorrecta de columna clave

- Usa la opción manual (opción 2) en el menú
- Verifica que la columna tenga valores únicos
- Revisa que el nombre esté escrito correctamente

---

## 📝 Notas

- Los archivos de salida se guardan en `reportes_comparacion/`
- Los archivos originales **nunca se modifican**
- Los reportes se organizan automáticamente por categoría
- Compatible con Excel 2007+ (.xlsx) y versiones antiguas (.xls)

---

## 🤝 Contribuciones

Para agregar funcionalidades:

1. Documenta cambios en este README
2. Mantén la compatibilidad con versiones anteriores
3. Actualiza ejemplos si es necesario

---

---
## import a utilizar

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import tkinter as tk
from tkinter import filedialog
import os
from pathlib import Path
```
---

## 🏷️ Versión

**Versión actual:** 1.0.0  
**Última actualización:** Enero 2026

---

## 📄 Licencia

Código de uso educativo y demostrativo.

---

**¡Listo para usar! 🚀**