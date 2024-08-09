import requests
import streamlit as st
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
import pdfplumber
import pandas as pd
import io
import re

# Configura aquí tu clave de API de OpenAI
openai_api_key = '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'

def extract_tables_from_pdf(file):
    """Función para extraer tablas de un PDF y convertirlas en DataFrame de pandas."""
    tablas = []
    contenido_texto = []
    try:
        with pdfplumber.open(file) as pdf:
            for numero_pagina, pagina in enumerate(pdf.pages):
                st.write(f"Procesando página {numero_pagina + 1}")
                for tabla in pagina.extract_tables():
                    if tabla:
                        # Mantener encabezados originales
                        encabezados = tabla[0]
                        df = pd.DataFrame(tabla[1:], columns=encabezados)
                        # Crear la fila de numeración de columnas
                        numeracion_columnas = [str(i+1) for i in range(len(encabezados))]
                        df_numeracion = pd.DataFrame([numeracion_columnas], columns=encabezados)
                        # Concatenar la fila de numeración, luego los encabezados, y luego los datos
                        df = pd.concat([df_numeracion, pd.DataFrame([encabezados], columns=encabezados), df]).reset_index(drop=True)
                        # Agregar índices numéricos a las filas
                        df.index = range(1, len(df) + 1)
                        tablas.append(df)
                contenido_texto.append(pagina.extract_text())
    except Exception as e:
        st.error(f"Error extrayendo tablas del PDF: {e}")
    return tablas, contenido_texto


def ocr_image(uploaded_file):
    """Función para realizar OCR en la imagen cargada, maneja tanto imágenes como PDF."""
    try:
        if uploaded_file.type == "application/pdf":
            images = convert_from_bytes(uploaded_file.getvalue())
            text = ' '.join([pytesseract.image_to_string(img) for img in images])
        else:
            image = Image.open(uploaded_file)
            text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        return f"Error al procesar la imagen: {e}"

def format_text_as_table(texto):
    """Función para convertir texto extraído en una tabla de pandas."""
    lineas = texto.split('\n')
    datos = []
    for linea in lineas:
        columnas = linea.split()  # Dividir por espacios en blanco
        datos.append(columnas)
    encabezados = datos[0]
    df = pd.DataFrame(datos[1:], columns=encabezados)  # Primera línea como encabezado
    # Insertar fila de numeración de columnas en la primera fila
    numeracion_columnas = [str(i+1) for i in range(len(encabezados))]
    df_numeracion = pd.DataFrame([numeracion_columnas], columns=encabezados)
    # Concatenar la fila de numeración y luego los datos
    df = pd.concat([pd.DataFrame([encabezados], columns=encabezados), df_numeracion, df]).reset_index(drop=True)
    # Agregar índices numéricos a las filas
    df.index = range(1, len(df) + 1)
    return df

def query_openai(texto):
    """Función para enviar texto a la API de OpenAI y obtener respuesta."""
    try:
        response = requests.post(
            "https://api.openai.com/v1/completions",
            headers={"Authorization": f"Bearer {openai_api_key}"},
            json={
                "model": "gpt-4",
                "prompt": texto,
                "max_tokens": 150
            }
        )
        data = response.json()
        return data['choices'][0]['text']
    except Exception as e:
        return f"Error al obtener respuesta de OpenAI: {e}"

def export_to_excel(df, sheet_name='Sheet1'):
    """Función para exportar el DataFrame a un archivo Excel."""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()  # Cambia save() por close()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.markdown(
        """
        <style>
        .stApp {
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid black;  /* Líneas entre columnas */
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("OCR DE MARKETPLACE S.A.")

    uploaded_file = st.file_uploader("Carga una imagen de factura aquí", type=["pdf", "png", "jpg", "jpeg"])
    if uploaded_file is not None:
        with st.spinner('Procesando archivo...'):
            if uploaded_file.type == "application/pdf":
                st.write("Archivo PDF subido correctamente.")
                tablas, contenido_texto = extract_tables_from_pdf(uploaded_file)
                if tablas:
                    for i, table in enumerate(tablas):
                        st.write(f"Tabla {i+1}")
                        st.dataframe(table)

                    # Botón para exportar las tablas a Excel
                    if st.button("Guardar Tablas en Excel"):
                        with st.spinner('Exportando tablas a Excel...'):
                            output = io.BytesIO()
                            writer = pd.ExcelWriter(output, engine='xlsxwriter')
                            for i, table in enumerate(tablas):
                                table.to_excel(writer, index=False, sheet_name=f'Tabla_{i+1}')
                            writer.close()  # Cambia save() por close()
                            processed_data = output.getvalue()
                            st.download_button(
                                label="Descargar Excel",
                                data=processed_data,
                                file_name='tablas_extraidas.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                else:
                    st.write("No se encontraron tablas en el PDF. Extrayendo texto.")
                    for texto_pagina in contenido_texto:
                        st.text_area("Texto Extraído", texto_pagina, height=300)
            else:
                texto = ocr_image(uploaded_file)
                if texto.startswith("Error"):
                    st.error(texto)
                else:
                    st.write("Texto extraído de la factura:")
                    st.text_area("Texto OCR", texto, height=200)  # Mostrar el texto extraído para depuración

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

                    # Enviar texto extraído a OpenAI
                    if st.button("Analizar Texto con OpenAI"):
                        with st.spinner('Obteniendo respuesta de OpenAI...'):
                            result = query_openai(texto)
                            st.write("Respuesta de OpenAI:")
                            st.write(result)
                        st.success("Datos guardados exitosamente en 'resultados_factura.xlsx'.")

if __name__ == "__main__":
    main()
