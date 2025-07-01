OCR
===

Aplicación simple en Streamlit para extraer texto y tablas de facturas en formato imagen o PDF.

## Instalación

```bash
pip install -r requirements.txt
```

La aplicación necesita Tesseract OCR instalado en el sistema para procesar las imágenes.

## Configuración

Defina la clave de API de OpenAI en la variable de entorno `OPENAI_API_KEY` antes de ejecutar la aplicación:

```bash
export OPENAI_API_KEY=tu-clave
```

## Uso

Ejecute la aplicación con Streamlit:

```bash
streamlit run ocr.py
```

Se abrirá una interfaz web donde podrá subir archivos y descargar los resultados en Excel o analizar el texto mediante OpenAI.
