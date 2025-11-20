import streamlit as st
import pandas as pd
import numpy as np
import easyocr
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from PIL import Image, ImageEnhance
import re
import os
import shutil

# =====================================
# OCR - EasyOCR
# =====================================

@st.cache_resource
def cargar_ocr():
    return easyocr.Reader(['es', 'en'])

ocr = cargar_ocr()

# =====================================
# PREPROCESAR IMAGEN
# =====================================

def preprocesar_imagen(img):
    img_gray = img.convert("L")
    img_enhanced = ImageEnhance.Contrast(img_gray).enhance(2.0)
    arr = np.array(img_enhanced)
    return arr

# =====================================
# LEER TEXTO (OCR) - CON MANEJO DE ERRORES
# =====================================

def leer_texto(img_array):
    try:
        textos = ocr.readtext(img_array, detail=0)  # detail=0 para solo texto
        return textos
    except Exception as e:
        st.error(f"Error en OCR: {str(e)}")
        return []

# =====================================
# DETECTAR CÓDIGOS
# =====================================

def detectar_codigos(textos):
    excluidos = [
        "sistemadeinformacion", "bibliografico", "biblioteca",
        "universidad", "cooperativa", "colombia"
    ]

    posibles = []

    for t in textos:
        limp = t.lower().replace(" ", "").replace("-", "").strip()

        if any(x in limp for x in excluidos):
            continue

        if re.fullmatch(r"b\d{6,8}", limp):
            posibles.append(limp.upper())
        elif limp.startswith("b") and len(limp) >= 7:
            posibles.append(limp.upper())

    return max(posibles, key=len) if posibles else None

# =====================================
# VALIDAR CÓDIGO
# =====================================

def validar_codigo(codigo, df):
    if not re.fullmatch(r"^B\d{6,8}$", codigo):
        return False, "Formato incorrecto (B + 6-8 dígitos)."

    if codigo in df["codigo"].astype(str).values:
        return False, "Código ya existe."

    return True, ""

# =====================================
# EXCEL
# =====================================

COLOR_VERDE = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
COLOR_MORADO = PatternFill(start_color="800080", end_color="800080", fill_type="solid")

EXCEL_PATH = "inventario.xlsx"
BACKUP_PATH = "inventario_backup.xlsx"

def actualizar_excel(codigo, wb, sheet, df):
    # Buscar fila dinámicamente para evitar inconsistencias
    codigos_existentes = df["codigo"].astype(str).values
    if codigo in codigos_existentes:
        fila = df[df["codigo"] == codigo].index[0] + 2
        celda = f"A{fila}"
        sheet[celda].fill = COLOR_VERDE
        sheet[celda].font = Font(bold=True)
        return f"✔ Código {codigo} marcado en verde."
    else:
        # Usar session_state para manejar confirmación sin loops
        key_confirm = f"confirmar_{codigo}"
        if key_confirm not in st.session_state:
            st.session_state[key_confirm] = False
        
        if st.button(f"Agregar nuevo código: {codigo}", key=f"btn_{codigo}"):
            st.session_state[key_confirm] = True
        
        if st.session_state[key_confirm]:
            nueva = sheet.max_row + 1
            sheet[f"A{nueva}"] = codigo
            sheet[f"A{nueva}"].fill = COLOR_MORADO
            sheet[f"A{nueva}"].font = Font(bold=True)
            
            # Actualizar DataFrame en session_state
            nuevo_df = pd.concat([df, pd.DataFrame({"codigo": [codigo]})], ignore_index=True)
            st.session_state["df"] = nuevo_df
            
            # Limpiar estado para evitar re-agregados
            del st.session_state[key_confirm]
            
            return f"➕ Código agregado: {codigo}"
        
        return "Pendiente de confirmación (haz clic en el botón)."

def crear_backup():
    if os.path.exists(EXCEL_PATH):
        shutil.copy(EXCEL_PATH, BACKUP_PATH)

# =====================================
# STREAMLIT UI
# =====================================

st.set_page_config(page_title="Inventario UCC", page_icon="📚", layout="wide")

if "df" not in st.session_state:
    st.session_state["df"] = None

st.title("📚 Inventario Biblioteca UCC - Medellín")

# CARGAR EXCEL
if not os.path.exists(EXCEL_PATH):
    st.error("No existe inventario.xlsx. Cárgalo.")
    f = st.file_uploader("Sube inventario", type="xlsx")
    if f:
        with open(EXCEL_PATH, "wb") as f_file:
            f_file.write(f.getbuffer())
        st.success("Cargado. Recarga la app.")
    st.stop()

wb = load_workbook(EXCEL_PATH)
sheet = wb.active

if st.session_state["df"] is None:
    st.session_state["df"] = pd.read_excel(EXCEL_PATH)

df = st.session_state["df"]

# =====================================
# ESCANEO
# =====================================

st.subheader("📷 Escanear código")
img_file = st.camera_input("Toma una foto del código")

if img_file:
    with st.spinner("Procesando..."):
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
            r = actualizar_excel(codigo, wb, sheet, df)
            st.info(r)

            if "✔" in r or "➕" in r:
                crear_backup()
                wb.save(EXCEL_PATH)
    else:
        st.warning("No se detectó un código válido.")

# =====================================
# DESCARGAS - SOLO EXCEL Y CSV (SIN COL3)
# =====================================

st.subheader("⬇ Descargas")

col1, col2 = st.columns(2)

with col1:
    with open(EXCEL_PATH, "rb") as f:
        st.download_button("Descargar Excel", f, file_name="inventario.xlsx")

with col2:
    st.download_button("Descargar CSV", data=df.to_csv(index=False), file_name="inventario.csv", mime="text/csv")


with col3:
    exportar_pdf(df, "inventario.pdf")  # Pasar filename explícitamente
    with open("inventario.pdf", "rb") as f:
        st.download_button("Descargar PDF", f, file_name="inventario.pdf")
