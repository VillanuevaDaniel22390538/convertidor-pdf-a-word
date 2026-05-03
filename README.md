# Convertidor Profesional de Documentos Hailie v2.0

![Version](https://img.shields.io/badge/version-2.0-green)
![Status](https://img.shields.io/badge/status-stable-blue)
![License](https://img.shields.io/badge/license-MIT-orange)

**Version 2.0**  
Desarrollado por **Daniel Donaldo Villanueva Canul**

---

## Capturas de pantalla

### Pantalla principal
![Dashboard](assets/Convertidor.png)

### Instrucciones de uso
![Instrucciones](assets/Convertidor_Instrucciones.png)

### Acerca del software
![Acerca](assets/Convertidor_Licencia.png)

---

## Descripcion

**Convertidor Profesional de Documentos Hailie** es una herramienta de escritorio que permite convertir multiples archivos entre distintos formatos de documentos de forma masiva, con una interfaz profesional, robusta y a prueba de errores.

### Formatos soportados

- PDF a Word (Microsoft Word)
- Word a PDF
- Word a TXT (texto plano)
- TXT a Word
- Imagen (PNG, JPG, JPEG, BMP) a PDF (combina todas las imagenes de una carpeta en un solo PDF, orden numerico automatico)

### Caracteristicas especiales

- Deteccion automatica: si un PDF tiene texto, se usa `pdf2docx` manteniendo el formato; si es escaneado, se aplica OCR con Tesseract (idioma espanol).
- Interfaz grafica con pestañas para organizar las conversiones.
- Barra de progreso y etiqueta de estado en tiempo real.
- Validacion de carpetas y manejo de errores por archivo (una conversion fallida no detiene el proceso).
- Estilo profesional sin emojis, colores neutros y tipografia clara.

---

## Requisitos previos

- Python 3.8 o superior (recomendado 3.12)
- Microsoft Word instalado (para la conversion Word -> PDF en Windows)
- Tesseract OCR instalado y agregado al PATH (para OCR en PDF escaneados)
- Poppler (opcional, solo si se usara la conversion PDF a imagenes, la cual no esta incluida en esta version)

---

## Instalacion

1. Clona el repositorio:
   ```bash
   git clone https://github.com/VillanuevaDaniel22390538/convertidor-pdf-a-word.git
   cd convertidor-pdf-a-word

Crea un entorno virtual (recomendado):
python -m venv venv
venv\Scripts\activate   # Windows
Instala las dependencias:
pip install -r requirements.txt
Ejecuta el programa:
python convert.py

Uso
Selecciona la carpeta de entrada (donde estan los archivos a convertir).
Selecciona la carpeta de salida (donde se guardaran los resultados).
Elige la pestaña correspondiente a la conversion deseada.
Haz clic en el boton de conversion.
Espera a que finalice el proceso; la barra de progreso te indicara el avance.

Archivo ejecutable
Puedes descargar la version compilada .exe desde la seccion Releases de este repositorio.
No requiere tener Python instalado.

Licencia
Este proyecto esta bajo la licencia MIT. Consulta el archivo LICENSE para mas detalles.

Creditos
Hecho con dedicacion por Daniel Donaldo Villanueva Canul
Correo: danieviloficialcontacto@gmail.com

Roadmap (versiones futuras)
Soporte para mas idiomas en OCR.
Modo oscuro.
Instalador con asistente (setup.exe).
Portable para Linux/Mac.
Conversion de PDF a imagenes (requiere Poppler).