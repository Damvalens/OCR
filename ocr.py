import os
import requests
import streamlit as st
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
import pdfplumber
import pandas as pd
import io
import re
import openai

# Estilos CSS para un diseño más moderno
MODERN_STYLE = """
<style>
/* Fondo de gradiente */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(120deg, #f6d365 0%, #fda085 100%);
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0);
}
.stButton > button {
    background-color: #007BFF;
    color: white;
    border-radius: 8px;
    padding: 0.5em 1em;
    transition: background-color 0.3s ease, transform 0.2s ease;
}
.stButton > button:hover {
    background-color: #0056b3;
    transform: scale(1.05);
}
</style>
"""

# Carga la clave de API de OpenAI desde la variable de entorno ``OPENAI_API_KEY``
openai_api_key = os.getenv("OPENAI_API_KEY")

def query_openai(texto: str) -> str:
    """Envía el texto a OpenAI y devuelve la respuesta generada."""
    if not openai_api_key:
        return "OPENAI_API_KEY no configurada."
    try:
        client = openai.OpenAI(api_key=openai_api_key)
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": texto}],
        )
        return completion.choices[0].message.content.strip()
    except Exception as exc:
        return f"Error consultando OpenAI: {exc}"

def make_unique_columns(columns):
    """Función para hacer que los nombres de las columnas sean únicos."""
    seen = {}
    for i, col in enumerate(columns):
        if col == '':
            col = 'Unnamed'
        if col in seen:
            seen[col] += 1
            columns[i] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    return columns

def extract_tables_from_pdf(pdf_bytes: bytes):
    """Extrae tablas de un PDF y devuelve el texto de cada página."""
    tablas = []
    contenido_texto = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total_paginas = len(pdf.pages)
            progreso = st.progress(0)
            for numero_pagina, pagina in enumerate(pdf.pages):
                progreso.progress((numero_pagina + 1) / total_paginas)
                st.write(f"Procesando página {numero_pagina + 1}/{total_paginas}")
                for tabla in pagina.extract_tables():
                    if tabla:
                        encabezados = tabla[0]
                        df = pd.DataFrame(tabla[1:], columns=encabezados)
                        # Asegurar que las columnas tengan nombres únicos
                        df.columns = make_unique_columns(df.columns.tolist())
                        df.index = range(1, len(df) + 1)  # Agregar índices numéricos a las filas
                        tablas.append(df)
                contenido_texto.append(pagina.extract_text() or "")
            progreso.empty()
    except Exception as e:
        st.error(f"Error extrayendo tablas del PDF: {e}")
    return tablas, contenido_texto

def ocr_pdf_to_text(pdf_bytes: bytes) -> str:
    """Realiza OCR sobre un PDF escaneado y devuelve el texto."""
    try:
        images = convert_from_bytes(pdf_bytes)
        return "\n".join(pytesseract.image_to_string(img) for img in images)
    except Exception as exc:
        return f"Error procesando PDF: {exc}"

def ocr_image(image_file) -> str:
    """Realiza OCR sobre una imagen."""
    try:
        image = Image.open(image_file)
        return pytesseract.image_to_string(image)
    except Exception as exc:
        return f"Error al procesar la imagen: {exc}"

def format_text_as_table(texto):
    """Función para convertir texto extraído en una tabla de pandas."""
    lineas = texto.split('\n')
    datos = []
    for linea in lineas:
        columnas = re.split(r'\s{2,}', linea.strip())  # Dividir por espacios en blanco de 2 o más
        datos.append(columnas)
    encabezados = datos[0]
    df = pd.DataFrame(datos[1:], columns=encabezados)
    df.columns = make_unique_columns(df.columns.tolist())  # Hacer los nombres de columnas únicos
    df.index = range(1, len(df) + 1)  # Agregar índices numéricos a las filas
    return df

def export_to_excel(df, sheet_name='Sheet1'):
    """Función para exportar el DataFrame a un archivo Excel."""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.set_page_config(page_title="OCR DE MARKETPLACE S.A.", page_icon="📄", layout="wide")
    st.markdown(MODERN_STYLE, unsafe_allow_html=True)
    st.markdown(
        """
        <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("OCR DE MARKETPLACE S.A.")

    st.sidebar.header("Carga de Documento")
    uploaded_file = st.sidebar.file_uploader(
        "Carga una imagen o PDF de factura", type=["pdf", "png", "jpg", "jpeg"]
    )
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        with st.spinner('Procesando archivo...'):
            if uploaded_file.type == "application/pdf":
                st.write("Archivo PDF subido correctamente.")
                tablas, contenido_texto = extract_tables_from_pdf(file_bytes)
                if tablas:
                    pestanas = st.tabs([f"Tabla {i+1}" for i in range(len(tablas))])
                    for i, table in enumerate(tablas):
                        with pestanas[i]:
                            st.dataframe(table, use_container_width=True)

                    if st.button("Guardar Tablas en Excel"):
                        with st.spinner('Exportando tablas a Excel...'):
                            output = io.BytesIO()
                            writer = pd.ExcelWriter(output, engine='xlsxwriter')
                            for i, table in enumerate(tablas):
                                table.to_excel(writer, index=False, sheet_name=f'Tabla_{i+1}')
                            writer.close()
                            processed_data = output.getvalue()
                            st.download_button(
                                label="Descargar Excel",
                                data=processed_data,
                                file_name='tablas_extraidas.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                        st.balloons()
                else:
                    st.write("No se encontraron tablas en el PDF. Aplicando OCR.")
                    texto = ocr_pdf_to_text(file_bytes)
                    st.text_area("Texto OCR", texto, height=300)
                    df = format_text_as_table(texto)
                    st.dataframe(df)

                    if st.button("Guardar en Excel"):
                        processed_data = export_to_excel(df)
                        st.download_button(
                            label="Descargar Excel",
                            data=processed_data,
                            file_name='resultados_factura.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        )
                        st.balloons()
                    if st.button("Analizar Texto con OpenAI"):
                        with st.spinner('Obteniendo respuesta de OpenAI...'):
                            result = query_openai(texto)
                            st.write("Respuesta de OpenAI:")
                            st.write(result)
                        st.snow()
            else:
                texto = ocr_image(uploaded_file)
                if texto.startswith("Error"):
                    st.error(texto)
                else:
                    st.write("Texto extraído de la factura:")
                    st.text_area("Texto OCR", texto, height=200)

                    # Organizar el texto extraído en columnas
                    df = format_text_as_table(texto)
                    st.dataframe(df)

                    # Botón para exportar los datos a Excel
                    if st.button("Guardar en Excel"):
                        processed_data = export_to_excel(df)
                        st.download_button(
                            label="Descargar Excel",
                            data=processed_data,
                            file_name='resultados_factura.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                        st.balloons()

                    # Enviar texto extraído a OpenAI
                    if st.button("Analizar Texto con OpenAI"):
                        with st.spinner('Obteniendo respuesta de OpenAI...'):
                            result = query_openai(texto)
                            st.write("Respuesta de OpenAI:")
                            st.write(result)
                        st.snow()
                        st.success("Datos guardados exitosamente en 'resultados_factura.xlsx'.")

if __name__ == "__main__":
    main()

