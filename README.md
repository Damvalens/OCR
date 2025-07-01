OCR
===

Aplicación simple en Streamlit para extraer texto y tablas de facturas en formato imagen o PDF.

## Instalación

```bash
pip install -r requirements.txt
```

La aplicación necesita Tesseract OCR instalado en el sistema para procesar las imágenes.
- **Linux**: `sudo apt-get install tesseract-ocr`
- **macOS**: `brew install tesseract`
- **Windows**: Descargue el instalador desde [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) y siga las instrucciones.

Asegúrese de que el ejecutable `tesseract` esté disponible en su variable de entorno `PATH`.

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
