# Convertidor Profesional de Documentos Hailie v2.1

![Version](https://img.shields.io/badge/version-2.1-green)
![Status](https://img.shields.io/badge/status-stable-blue)
![License](https://img.shields.io/badge/license-MIT-orange)

**Versión 2.1**  
Desarrollado por **Daniel Donaldo Villanueva Canul**  
Contacto: danieviloficialcontacto@gmail.com

---

## Capturas de pantalla

### Pantalla principal
![Dashboard](assets/Convertidor.png)

### Instrucciones de uso
![Instrucciones](assets/Convertidor_Instrucciones.png)

### Acerca del software
![Acerca](assets/Convertidor_Licencia.png)

---

## Descripción

**Convertidor Profesional de Documentos Hailie** es una herramienta de escritorio que permite convertir múltiples archivos entre distintos formatos de documentos de forma masiva, con una interfaz profesional, robusta y a prueba de errores.

### Formatos soportados

- PDF a Word (Microsoft Word)
- Word a PDF (con respaldo automático a LibreOffice)
- Word a TXT (texto plano)
- TXT a Word
- Imagen (PNG, JPG, JPEG, BMP) a PDF (combina todas las imágenes de una carpeta en un solo PDF, orden numérico automático)

### Características especiales

- Detección automática: si un PDF tiene texto, se usa `pdf2docx` manteniendo el formato; si es escaneado, se aplica OCR con Tesseract (idioma español).
- Conversión Word a PDF inteligente: primero intenta con Microsoft Word, si falla o no está instalado, utiliza LibreOffice automáticamente.
- Interfaz gráfica con pestañas para organizar las conversiones.
- Barra de progreso y etiqueta de estado en tiempo real.
- Validación de carpetas y manejo de errores por archivo (una conversión fallida no detiene el proceso).
- Estilo profesional sin emojis, colores neutros y tipografía clara.
- Botón de donación para apoyar el desarrollo continuo.

---

## Requisitos previos

- Python 3.8 o superior (recomendado 3.12)
- Microsoft Word o LibreOffice (para la conversión Word -> PDF). Si tienes ambos, el programa usará Word primero.
- Tesseract OCR instalado y agregado al PATH (para OCR en PDF escaneados) - opcional.
- Poppler (opcional, solo si usas conversión PDF a imágenes, no incluida en esta versión).

---

## Instalación

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
Selecciona la carpeta de entrada (donde están los archivos a convertir).

Selecciona la carpeta de salida (donde se guardarán los resultados).

Elige la pestaña correspondiente a la conversión deseada.

Haz clic en el botón de conversión.

Espera a que finalice el proceso; la barra de progreso te indicará el avance.

Archivo ejecutable
Puedes descargar la versión compilada .exe desde la sección Releases de este repositorio.
No requiere tener Python instalado.

Donaciones
Si este software te ha sido útil y deseas apoyar su desarrollo continuo, puedes hacer una donación a través de PayPal: https://www.paypal.com/paypalme/danievilmathers?locale.x=es_XC&country.x=MX 
Donar con PayPal
Cualquier aportación es bienvenida y ayuda a mantener el proyecto activo.

Licencia
Este proyecto está bajo la licencia MIT. Consulta el archivo LICENSE para más detalles.

Créditos
Hecho con dedicación por Daniel Donaldo Villanueva Canul
Correo: danieviloficialcontacto@gmail.com

Roadmap (versiones futuras)
Soporte para más idiomas en OCR.
Modo oscuro.
Instalador con asistente (setup.exe).
Portable para Linux/Mac.
Conversión de PDF a imágenes (requiere Poppler).