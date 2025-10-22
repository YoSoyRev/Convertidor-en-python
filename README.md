# 📄 Convertidor Universal de Documentos a PDF (Python)

Este es un script de Python que permite convertir por lotes una variedad de formatos de documentos (Word, Excel, PowerPoint, e Imágenes) a archivos PDF, utilizando una interfaz gráfica simple para seleccionar los archivos y la carpeta de destino.

## 🌟 Características

* **Soporte Multi-Formato:** Convierte archivos `.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt` y los formatos de imagen más comunes (`.jpg`, `.png`, `.tiff`) a PDF.
* **Selección Gráfica:** Utiliza ventanas de diálogo para seleccionar múltiples archivos de entrada y una carpeta de destino.
* **Procesamiento por Lotes:** Procesa varios archivos de una sola vez.
* **Barra de Progreso:** Muestra el progreso de la conversión en tiempo real a través de la consola (`tqdm`).
* **Automatización de Office:** En entornos Windows, utiliza la automatización COM para asegurar la máxima fidelidad en la conversión de documentos de Microsoft Office.

## ⚠️ Requisitos

Este script tiene dependencias específicas de software y sistema operativo:

1.  **Python:** Versión 3.x.
2.  **Sistema Operativo:** **Windows** (Necesario para la conversión de Excel y PowerPoint a través de `pywin32`).
3.  **Microsoft Office:** Se requiere tener instalado **Microsoft Word, Excel y PowerPoint** para que las conversiones de `.docx`, `.xlsx` y `.pptx` funcionen correctamente.

## 🛠️ Instalación y Uso

### 1. Clona el Repositorio

Primero, descarga o clona el código del proyecto a tu máquina local.

### 2. Instalación de Librerías

Abre tu terminal o Símbolo del sistema y ejecuta el siguiente comando para instalar todas las dependencias necesarias de Python:

```bash
pip install docx2pdf img2pdf Pillow tqdm pywin32
