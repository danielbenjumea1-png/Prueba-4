import streamlit as st
import pandas as pd
import numpy as np
import pytesseract
import cv2
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from PIL import Image, ImageEnhance
import re
import os
import shutil
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# ============================
# OCR - PREPROCESAMIENTO
# ============================

# Preprocesar imagen → NumPy BGR
def preprocesar_imagen(img):
    # Mejoras
    img_gray = img.convert('L')
    img_enhanced = ImageEnhance.Contrast(img_gray).enhance(2.0)

    # Convertir PIL → NumPy
    arr = np.array(img_enhanced)

    # Convertir GRAY → BGR (necesario para cvtColor)
    arr_bgr = cv2.cvtColor(arr, cv2.COLOR_GRAY2BGR)
    return arr_bgr


# OCR usando Tesseract
def leer_texto(img_array):
    gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray)
    return [t.strip() for t in text.splitlines() if t.strip()]


# Detectar códigos Bxxxxxx
def detectar_codigos(textos):
    frases_prohibidas = [
        "sistemadeinformacionbibliografico",
        "sistemadeinformacion",
        "bibliografico",
        "biblioteca",
        "universidad",
        "cooperativa",
        "colombia"
    ]
    posibles = []

    for t in textos:
        limp = t.lower().replace(" ", "").replace("-", "").strip()
        if any(f in limp for f in frases_prohibidas):
            continue

        if re.fullmatch(r"b\d{6,8}", limp):
            posibles.append(limp.upper())
        elif limp.startswith("b") and len(limp) >= 7:
            posibles.append(limp.upper())

    return max(posibles, key=len) if posibles else None


# Validación
def validar_codigo(codigo, df):
    if not re.fullmatch(r"^B\d{6,8}$", codigo):
        return False, "Formato inválido (debe ser B seguido de 6-8 dígitos)."
    if codigo in df['codigo'].values:
        return False, "Código ya existe."
    return True, ""


# ============================
# EXCEL
# ============================

COLOR_VERDE = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
COLOR_MORADO = PatternFill(start_color="800080", end_color="800080", fill_type="solid")

EXCEL_PATH = "inventario.xlsx"
BACKUP_PATH = "inventario_backup.xlsx"


def actualizar_excel(codigo, wb, sheet, codigo_a_fila, df):
    try:
        if codigo in codigo_a_fila:
            fila = codigo_a_fila[codigo]
            celda = f"A{fila}"
            sheet[celda].fill = COLOR_VERDE
            sheet[celda].font = Font(bold=True)
            return f"✔ Código {codigo} encontrado y marcado en verde."
        else:
            if st.button(f"Confirmar agregar código nuevo: {codigo}"):
                nueva = sheet.max_row + 1
                sheet[f"A{nueva}"] = codigo
                sheet[f"A{nueva}"].fill = COLOR_MORADO
                sheet[f"A{nueva}"].font = Font(bold=True)
                nuevo_df = pd.concat([df, pd.DataFrame({'codigo': [codigo]})], ignore_index=True)
                st.session_state['df'] = nuevo_df
                return f"➕ Código nuevo agregado: {codigo}"
            return "Pendiente de confirmación."
    except Exception as e:
        return f"Error al actualizar Excel: {str(e)}"


def crear_backup():
    if os.path.exists(EXCEL_PATH):
        shutil.copy(EXCEL_PATH, BACKUP_PATH)


# ============================
# EXPORTAR PDF
# ============================

def exportar_pdf(df, filename="inventario.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    c.drawString(100, 750, "Inventario Biblioteca UCC")
    y = 720
    for index, row in df.iterrows():
        c.drawString(100, y, f"Código: {row['codigo']}")
        y -= 20
        if y < 50:
            c.showPage()
            y = 750
    c.save()


# ============================
# STREAMLIT - INICIO
# ============================

st.set_page_config(page_title="📚 Inventario Biblioteca UCC", page_icon="📚", layout="wide")

if 'df' not in st.session_state:
    st.session_state['df'] = None

if 'codigos_detectados' not in st.session_state:
    st.session_state['codigos_detectados'] = []

st.title("📚 Inventario Biblioteca UCC - Sede Medellín")

with st.expander("📖 Guía de Uso"):
    st.write("""
    - **Escaneo**: Toma una foto clara del código.
    - **Manual**: Puedes escribir el código.
    - **Batch**: Sube varias imágenes.
    - **Buscar**: Filtra el inventario.
    - Descarga Excel, CSV o PDF.
    """)

# ============================
# Cargar o pedir Excel
# ============================

if not os.path.exists(EXCEL_PATH):
    st.error("No se encontró 'inventario.xlsx'.")
    file = st.file_uploader("Sube tu inventario inicial", type=["xlsx"])
    if file:
        with open(EXCEL_PATH, "wb") as f:
            f.write(file.getbuffer())
        st.success("Inventario cargado. Recarga la página.")
    st.stop()

try:
    wb = load_workbook(EXCEL_PATH)
    sheet = wb.active
    if st.session_state['df'] is None:
        st.session_state['df'] = pd.read_excel(EXCEL_PATH)
    df = st.session_state['df']
except Exception as e:
    st.error(f"Error al cargar Excel: {str(e)}")
    st.stop()

# Buscar columna
codigo_columna = None
for col in df.columns:
    if "codigo" in col.lower():
        codigo_columna = col
        break

if not codigo_columna:
    st.error("No existe una columna 'codigo'.")
    st.stop()

codigo_a_fila = {str(row[codigo_columna]).strip(): idx + 2 for idx, row in df.iterrows()}

# ============================
# ESCANEO
# ============================

st.subheader("📷 Escanear código")
img_file = st.camera_input("Toma una foto")

if img_file:
    with st.spinner("Procesando imagen..."):
        img = Image.open(img_file)
        arr = preprocesar_imagen(img)
        textos = leer_texto(arr)
        codigo = detectar_codigos(textos)

    if codigo:
        st.success(f"Código detectado: **{codigo}**")
        valido, msg = validar_codigo(codigo, df)
        if not valido:
            st.warning(msg)
        else:
            resultado = actualizar_excel(codigo, wb, sheet, codigo_a_fila, df)
            st.info(resultado)
            if "agregado" in resultado or "marcado" in resultado:
                crear_backup()
                wb.save(EXCEL_PATH)
    else:
        st.warning("No se detectó un código válido.")

# ============================
# BATCH
# ============================

st.subheader("📂 Procesar varias imágenes")
files = st.file_uploader("Sube imágenes", accept_multiple_files=True, type=["jpg", "jpeg", "png"])

if files:
    for f in files:
        img = Image.open(f)
        arr = preprocesar_imagen(img)
        text = leer_texto(arr)
        codigo = detectar_codigos(text)
        if codigo:
            valido, _ = validar_codigo(codigo, df)
            if valido:
                actualizar_excel(codigo, wb, sheet, codigo_a_fila, df)

    crear_backup()
    wb.save(EXCEL_PATH)
    st.success("Batch procesado.")

# ============================
# INGRESO MANUAL
# ============================

st.subheader("✍ Ingreso manual")
codigo_manual = st.text_input("Escribe el código:")

if codigo_manual:
    codigo_manual = codigo_manual.strip().upper()
    valido, msg = validar_codigo(codigo_manual, df)
    if not valido:
        st.warning(msg)
    else:
        resultado = actualizar_excel(codigo_manual, wb, sheet, codigo_a_fila, df)
        st.info(resultado)
        if "agregado" in resultado or "marcado" in resultado:
            crear_backup()
            wb.save(EXCEL_PATH)

# ============================
# BUSCAR
# ============================

st.subheader("🔍 Buscar")
query = st.text_input("Buscar por código:")
if query:
    st.dataframe(df[df[codigo_columna].str.contains(query, case=False, na=False)])
else:
    st.dataframe(df)

# ============================
# DESCARGAS
# ============================

st.subheader("⬇ Descargas")
col1, col2, col3 = st.columns(3)

with col1:
    with open(EXCEL_PATH, "rb") as f:
        st.download_button("Excel", f, file_name="inventario_actualizado.xlsx")

with col2:
    csv = df.to_csv(index=False)
    st.download_button("CSV", csv, file_name="inventario.csv", mime="text/csv")

with col3:
    exportar_pdf(df)
    with open("inventario.pdf", "rb") as f:
        st.download_button("PDF", f, file_name="inventario.pdf")
