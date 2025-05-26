import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import zipfile
import os

st.set_page_config(page_title="Generador de PDFs en ZIP por Empresa", layout="centered")
st.title("📄 Generador de PDFs agrupados por NCODIGOPJ y descargables en ZIP")

uploaded_file = st.file_uploader("Sube el archivo Excel", type=["xlsx"])
pdf_title = st.text_input("Título principal para los PDFs", value="Informe de Empresa")

def draw_row(pdf, col_widths, data, line_height=5):
    x_start = pdf.get_x()
    y_start = pdf.get_y()

    max_lines = 0
    text_lines = []

    pdf.set_font("Arial", '', 9)

    for i, text in enumerate(data):
        max_chars = int(col_widths[i] / 2.5)
        words = text.split()
        lines = []
        current_line = ""

        for word in words:
            if len(current_line + " " + word) <= max_chars:
                current_line += " " + word if current_line else word
            else:
                lines.append(current_line)
                current_line = word
        lines.append(current_line)
        text_lines.append(lines)
        max_lines = max(max_lines, len(lines))

    row_height = max_lines * line_height

    if pdf.get_y() + row_height > pdf.page_break_trigger:
        pdf.add_page()
        draw_titles(pdf, title=pdf_title, empresa=pdf.empresa_nombre)
        draw_header(pdf, col_widths, ["APELLIDOS Y NOMBRES", "EMAIL", "PERFIL", "CARGOS", "FECHA INICIAL"], line_height)

    for i, lines in enumerate(text_lines):
        x = pdf.get_x()
        y = pdf.get_y()
        cell_width = col_widths[i]
        pdf.rect(x, y, cell_width, row_height)

        for idx, line in enumerate(lines):
            pdf.set_xy(x, y + idx * line_height)
            pdf.cell(cell_width, line_height, line, ln=0)

        pdf.set_xy(x + cell_width, y)

    pdf.set_xy(x_start, y_start + row_height)

def draw_header(pdf, col_widths, headers, line_height=5):
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    row_height = 2 * line_height

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(46, 139, 87)
    pdf.set_draw_color(0, 0, 0)
    for i, header in enumerate(headers):
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.set_xy(x, y)
        pdf.cell(col_widths[i], row_height, header, border=1, align='C', fill=True)
        pdf.set_xy(x + col_widths[i], y)

    pdf.set_xy(x_start, y_start + row_height)

def draw_titles(pdf, title, empresa):
    pdf.set_font("Arial", 'B', 14)
    pdf.set_text_color(0, 0, 128)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, empresa, ln=True, align="L")
    pdf.ln(5)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    if "NCODIGOPJ" not in df.columns:
        st.error("El archivo no contiene la columna 'NCODIGOPJ'")
    else:
        grouped = df.groupby(['NCODIGOPJ', 'EMPRESA'])

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for (ncodigopj, empresa), grupo in grouped:
                pdf = FPDF(orientation='L')
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.empresa_nombre = str(empresa)

                draw_titles(pdf, title=pdf_title, empresa=empresa)

                headers = ["APELLIDOS Y NOMBRES", "EMAIL", "PERFIL", "CARGOS", "FECHA INICIAL"]
                col_widths = [50, 51, 44, 60, 35]

                draw_header(pdf, col_widths, headers, line_height=5)
                pdf.set_font("Arial", '', 7)

                for _, row in grupo.iterrows():
                    values = [
                        str(row["APELLIDOS Y NOMBRES"]),
                        str(row["EMAIL"]),
                        str(row["PERFIL"]),
                        str(row["CARGOS"]).replace("<BR>", " / "),
                        str(row["FECHA INICIAL"])
                    ]
                    draw_row(pdf, col_widths, values, line_height=5)

                pdf_bytes = BytesIO(pdf.output(dest='S').encode('latin1'))
                clean_name = str(int(ncodigopj)) if str(ncodigopj).endswith('.0') else str(ncodigopj)
                filename = f"{clean_name}.pdf"
                zip_file.writestr(filename, pdf_bytes.read())

        zip_buffer.seek(0)
        st.download_button(
            label="📥 Descargar ZIP con todos los PDFs",
            data=zip_buffer,
            file_name="PDFs_empresas.zip",
            mime="application/zip"
        )

