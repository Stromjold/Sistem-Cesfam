# Sistema de Limpieza y Cruce de Datos

Sistema unificado para procesar archivos Excel con códigos de color, extraer RUTs y generar reportes.

## 📋 Requisitos Previos

- Python 3.8 o superior
- Sistema operativo Windows
- Microsoft Excel instalado (para xlwings)

## 🚀 Instalación Rápida

### En Windows:

1. **Primera vez** - Instalar dependencias:
   ```
   setup.bat
   ```
   Este script:
   - Crea un entorno virtual Python
   - Instala todas las dependencias necesarias
   - Configura el proyecto automáticamente

2. **Ejecutar el programa**:
   ```
   ejecutar.bat
   ```

### Instalación Manual (opcional):

Si prefieres instalar manualmente:

```bash
# Crear entorno virtual
python -m venv .venv

# Activar entorno virtual
.venv\Scripts\activate.bat

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar programa
python system_complete_fixed.py
```

## 📦 Dependencias

El proyecto usa las siguientes librerías:
- **xlwings**: Para interactuar con Excel y leer códigos de color
- **pandas**: Para procesamiento de datos
- **openpyxl**: Para leer/escribir archivos Excel
- **tkinter**: Para interfaz gráfica (incluido con Python)

## 📁 Estructura del Proyecto

```
Carpeta_principal/
├── system_complete_fixed.py  # Programa principal
├── requirements.txt           # Dependencias del proyecto
├── setup.bat                  # Instalador automático
├── ejecutar.bat               # Ejecutor del programa
├── README.md                  # Este archivo
├── Archivo_Entrada/           # Archivos de entrada
├── Archivos_escanear/         # Archivos a procesar
└── Resultados/                # Reportes generados
```

## 🔧 Funcionalidades

1. **Lee colores en Excel**: Separa registros según color de celda (rojos/blancos)
2. **Extrae RUTs limpios**: Normaliza y limpia RUTs automáticamente
3. **Búsqueda automática**: Procesa todos los archivos de la carpeta
4. **Reporte unificado**: Genera un único archivo con resultados

## 💡 Uso

1. Coloca el archivo principal en `Archivo_Entrada/`
2. Coloca los archivos a escanear en `Archivos_escanear/`
3. Ejecuta `ejecutar.bat`
4. Los resultados se guardarán en `Resultados/`

## 🌐 Portabilidad

Este proyecto está configurado para ser completamente portable:
- Todas las dependencias están especificadas en `requirements.txt`
- Usa entorno virtual local (`.venv`)
- Scripts de instalación automática incluidos
- Funciona en cualquier Windows con Python instalado

Para mover a otro dispositivo:
1. Copia toda la carpeta `Carpeta_principal`
2. Ejecuta `setup.bat` en el nuevo dispositivo
3. ¡Listo para usar!

## ⚠️ Notas Importantes

- **Excel debe estar instalado** en el sistema para que xlwings funcione correctamente
- El entorno virtual (`.venv`) puede ser grande. Si quieres reducir el tamaño para compartir, elimina la carpeta `.venv` y el usuario final ejecutará `setup.bat` para recrearla
- Los archivos de entrada/salida no se incluyen por defecto, solo la estructura de carpetas

## 🐛 Solución de Problemas

**Error: Python no encontrado**
- Instala Python desde https://www.python.org/downloads/
- Asegúrate de marcar "Add Python to PATH" durante la instalación

**Error: xlwings no funciona**
- Verifica que Microsoft Excel esté instalado
- Ejecuta: `pip install --upgrade xlwings`

**Error: Permisos denegados**
- Ejecuta los .bat como administrador (click derecho → Ejecutar como administrador)

## 📝 Licencia

Proyecto interno de uso privado.
