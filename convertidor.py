import os
import img2pdf
from docx2pdf import convert as convert_docx
from PIL import Image
from tkinter import filedialog
import tkinter as tk
from tqdm import tqdm
import win32com.client 

def convertir_docx_a_pdf(ruta_entrada, directorio_salida):
    nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
    ruta_salida = os.path.join(directorio_salida, f"{nombre_base}.pdf")
    
    try:
        convert_docx(ruta_entrada, ruta_salida)
        return ruta_salida
    except Exception as e:
        print(f"\nERROR DOCX ({os.path.basename(ruta_entrada)}): Falló la conversión de Word. Verifique que no esté abierto.")
        print(f"Detalle: {e}")
        return None

def convertir_imagen_a_pdf(ruta_entrada, directorio_salida):
    nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
    ruta_salida = os.path.join(directorio_salida, f"{nombre_base}.pdf")
    
    try:
        with Image.open(ruta_entrada) as img:
            img.verify()
        
        with open(ruta_salida, "wb") as f:
            f.write(img2pdf.convert(ruta_entrada))
            
        return ruta_salida
    except Exception as e:
        print(f"\nERROR IMAGEN ({os.path.basename(ruta_entrada)}): No se pudo convertir la imagen.")
        print(f"Detalle: {e}")
        return None

def convertir_excel_a_pdf(ruta_entrada, directorio_salida):
    nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
    ruta_salida = os.path.join(directorio_salida, f"{nombre_base}.pdf")
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(os.path.abspath(ruta_entrada))
        
        # 0 = Formato PDF, Exporta a PDF todo el libro
        wb.ExportAsFixedFormat(0, os.path.abspath(ruta_salida))
        
        wb.Close(False)
        excel.Quit()
        return ruta_salida
    except Exception as e:
        print(f"\nERROR EXCEL ({os.path.basename(ruta_entrada)}): Falló la automatización de Excel.")
        print(f"Detalle: {e}")
        return None

def convertir_pptx_a_pdf(ruta_entrada, directorio_salida):
    nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
    ruta_salida = os.path.join(directorio_salida, f"{nombre_base}.pdf")
    
    try:
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = False
        
        presentation = powerpoint.Presentations.Open(os.path.abspath(ruta_entrada), WithWindow=False)
        
        # 32 = Formato PDF
        presentation.ExportAsFixedFormat(os.path.abspath(ruta_salida), 32)
        
        presentation.Close()
        powerpoint.Quit()
        return ruta_salida
    except Exception as e:
        print(f"\nERROR PPTX ({os.path.basename(ruta_entrada)}): Falló la automatización de PowerPoint.")
        print(f"Detalle: {e}")
        return None

def convertidor_universal_pdf(ruta_entrada, directorio_salida):
    extension = ruta_entrada.lower().split('.')[-1]
    ruta_pdf = None

    if extension in ['docx', 'doc']:
        ruta_pdf = convertir_docx_a_pdf(ruta_entrada, directorio_salida)
        
    elif extension in ['jpg', 'jpeg', 'png', 'tiff', 'webp']:
        ruta_pdf = convertir_imagen_a_pdf(ruta_entrada, directorio_salida)
        
    elif extension in ['xlsx', 'xls']:
        ruta_pdf = convertir_excel_a_pdf(ruta_entrada, directorio_salida) 
        
    elif extension in ['pptx', 'ppt']:
        ruta_pdf = convertir_pptx_a_pdf(ruta_entrada, directorio_salida) 
        
    elif extension == 'pdf':
        return None
        
    else:
        print(f"\nADVERTENCIA: El formato '.{extension}' no está soportado. Saltando.")
        return None

    return ruta_pdf


def seleccionar_archivos_y_convertir():
    root = tk.Tk()
    root.withdraw() 
    
    tipos_archivos = [
        ('Archivos de Word', ('*.docx', '*.doc')),
        ('Archivos de Excel', ('*.xlsx', '*.xls')),
        ('Archivos de PowerPoint', ('*.pptx', '*.ppt')),
        ('Imágenes', ('*.jpg', '*.jpeg', '*.png', '*.tiff', '*.webp')),
        ('Todos los soportados', ('*.docx', '*.doc', '*.xlsx', '*.xls', '*.pptx', '*.ppt', '*.jpg', '*.jpeg', '*.png', '*.tiff', '*.webp')),
        ('Todos los archivos', '*.*')
    ]
    
    rutas_archivos = filedialog.askopenfilenames(
        title="1. Selecciona los documentos para convertir a PDF",
        filetypes=tipos_archivos
    )
    
    if not rutas_archivos:
        print("CANCELADO: No se seleccionó ningún archivo. Saliendo.")
        return

    directorio_salida = filedialog.askdirectory(
        title="2. Selecciona la carpeta donde guardar los archivos PDF"
    )
    
    if not directorio_salida:
        print("CANCELADO: No se seleccionó carpeta de destino. Saliendo.")
        return
    
    print(f"\nCarpeta de destino: {directorio_salida}")
    print(f"Iniciando conversión de {len(rutas_archivos)} archivos...")
    print("=" * 60)
    
    archivos_fallidos = []

    for ruta in tqdm(rutas_archivos, desc="Progreso Total", unit=" archivo", dynamic_ncols=True):
        ruta_pdf = convertidor_universal_pdf(ruta, directorio_salida)
        
        if ruta_pdf is None and os.path.basename(ruta).lower().split('.')[-1] not in ['pdf']:
            archivos_fallidos.append(os.path.basename(ruta))
        
    print("=" * 60)
    print("PROCESO TERMINADO: Conversión por lotes finalizada.")
    
    if archivos_fallidos:
        print("\n--- RESUMEN DE FALLOS ---")
        print(f"Archivos que NO pudieron convertirse ({len(archivos_fallidos)}):")
        for nombre in archivos_fallidos:
            print(f"- {nombre}")
        print("Verifique los mensajes de error anteriores.")


if __name__ == "__main__":
    seleccionar_archivos_y_convertir()