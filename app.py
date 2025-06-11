import streamlit as st
import pandas as pd
import tempfile
import os
import time
from funciones_utils import (
    detectar_columna_imagenes,
    descargar_imagenes_temp,
    convertir_columnas_a_str,
    filtrar_df_por_imagenes
)
from plantilla_basica import generar_presentacion_basica
from plantilla_exhibiciones import generar_presentacion_exhibiciones
from pptx.util import Inches
from PIL import Image

st.set_page_config(page_title="Generador PPT", layout="centered")
st.title("📊 Generador de Presentaciones PowerPoint desde Excel con Imágenes")

# Estado de sesión inicial
if "generado" not in st.session_state:
    st.session_state.generado = False
if "archivos_generados" not in st.session_state:
    st.session_state.archivos_generados = []
if "archivo_excel" not in st.session_state:
    st.session_state.archivo_excel = None
if "en_proceso" not in st.session_state:
    st.session_state.en_proceso = False

# Formulario de carga y configuración
if not st.session_state.generado and not st.session_state.en_proceso:
    with st.container():
        with st.expander("📁 Paso 1: Sube tu archivo Excel", expanded=True):
            archivo_excel = st.file_uploader("Sube un archivo Excel con links a imágenes", type=[".xlsx", ".xls"])
        
        if archivo_excel:
            st.toast("✅ Archivo cargado con éxito")
            st.session_state.archivo_excel = archivo_excel
            df = pd.read_excel(archivo_excel)
            convertir_columnas_a_str(df)
            columna_img = detectar_columna_imagenes(df)
            
            if columna_img:
                df = filtrar_df_por_imagenes(df, columna_img)
                st.success(f"🔍 Columna de imágenes detectada automáticamente: `{columna_img}`")

                st.divider()
                st.subheader("🎨 Paso 2: Configura tu presentación")
                plantilla_opcion = st.radio("Selecciona plantilla", ["Básica", "Exhibiciones"], horizontal=True)
                st.session_state.plantilla_opcion = plantilla_opcion

                if plantilla_opcion == "Básica":
                    st.markdown("**Opciones para plantilla básica:**")
                    encabezados_seleccionados = st.multiselect("Selecciona columnas para mostrar como texto", df.columns.tolist())
                    fotos_por_slide = st.slider("Cantidad de fotos por diapositiva", min_value=1, max_value=4, value=2)
                    color_fuente = st.color_picker("Color de texto", "#000000")

                    st.session_state.encabezados_seleccionados = encabezados_seleccionados
                    st.session_state.fotos_por_slide = fotos_por_slide
                    st.session_state.color_fuente = color_fuente

                elif plantilla_opcion == "Exhibiciones":
                    st.markdown("**Opciones para plantilla de exhibiciones:**")
                    agrupador_1 = st.selectbox("Columna para título del slide", df.columns.tolist())
                    color_agrupador_1 = st.color_picker("Color del título", "#000000")
                    agrupador_2 = st.selectbox("Columna para subtítulo", df.columns.tolist())
                    color_agrupador_2 = st.color_picker("Color del subtítulo", "#000000")
                    columna_encabezado = st.selectbox("Columna para encabezado (opcional)", [""] + df.columns.tolist())
                    color_encabezado = st.color_picker("Color del encabezado", "#000000")

                    st.session_state.agrupador_1 = agrupador_1
                    st.session_state.color_agrupador_1 = color_agrupador_1
                    st.session_state.agrupador_2 = agrupador_2
                    st.session_state.color_agrupador_2 = color_agrupador_2
                    st.session_state.columna_encabezado = columna_encabezado if columna_encabezado else None
                    st.session_state.color_encabezado = color_encabezado

                tipo_fuente = st.selectbox("Tipo de fuente", ["Arial", "Calibri", "Times New Roman"])
                fondo_bytes = st.file_uploader("Fondo opcional (imagen)", type=[".jpg", ".jpeg", ".png"])

                st.session_state.tipo_fuente = tipo_fuente
                st.session_state.fondo_bytes = fondo_bytes

                st.divider()
                if st.button("🚀 Generar PowerPoint"):
                    st.session_state.en_proceso = True
                    st.rerun()

# Proceso de generación con feedback visual
if st.session_state.en_proceso:
    with st.spinner("⏳ Generando presentación, por favor espera..."):
        start_time = time.time()

        temp_dir = tempfile.mkdtemp()
        status_text = st.empty()
        progress_bar = st.progress(0)
        tiempo_texto = st.empty()

        df = pd.read_excel(st.session_state.archivo_excel)
        convertir_columnas_a_str(df)
        columna_img = detectar_columna_imagenes(df)
        df = filtrar_df_por_imagenes(df, columna_img)
        df["img_path"] = descargar_imagenes_temp(df, columna_img, temp_dir)
        df_filtrado = df[df["img_path"].notna()]

        plantilla_opcion = st.session_state.plantilla_opcion
        tipo_fuente = st.session_state.tipo_fuente
        fondo_bytes = st.session_state.fondo_bytes

        if plantilla_opcion == "Básica":
            resultado = generar_presentacion_basica(
                df_filtrado,
                "presentacion",
                st.session_state.fotos_por_slide,
                st.session_state.encabezados_seleccionados,
                fondo_bytes,
                tipo_fuente,
                st.session_state.color_fuente,
                temp_dir,
                tiempo_texto,
                start_time
            )
        else:
            resultado = generar_presentacion_exhibiciones(
                df_filtrado,
                st.session_state.agrupador_1,
                st.session_state.agrupador_2,
                fondo_bytes,
                tipo_fuente,
                status_text,
                progress_bar,
                temp_dir,
                'img_path',
                st.session_state.color_agrupador_1,
                st.session_state.color_agrupador_2,
                st.session_state.color_encabezado,
                st.session_state.columna_encabezado,
                tiempo_texto,
                start_time
            )

        st.session_state.archivos_generados = resultado if isinstance(resultado, list) else [("presentacion", resultado)]
        st.session_state.generado = True
        st.session_state.en_proceso = False
        st.session_state.tiempo_generacion = round(time.time() - start_time, 2)
        st.rerun()

# Mostrar resultado y descargas
if st.session_state.generado:
    st.balloons()
    st.success("✅ Presentación generada con éxito.")
    st.info(f"⏱ Tiempo total: {st.session_state.tiempo_generacion} segundos")

    for nombre, path in st.session_state.archivos_generados:
        with open(path, "rb") as f:
            st.download_button(f"📥 Descargar {nombre}.pptx", f, file_name=f"{nombre}.pptx")

    if st.button("🔄 Generar otra vez"):
        st.session_state.clear()
        st.rerun()

# Auxiliar para tamaños de imagen
def calcular_dimensiones_auto(img_path, max_width_in, max_height_in):
    try:
        with Image.open(img_path) as img:
            width_px, height_px = img.size
            width_in, height_in = width_px / 96, height_px / 96
            ratio = min(max_width_in / width_in, max_height_in / height_in)
            return Inches(width_in * ratio), Inches(height_in * ratio)
    except Exception as e:
        print(f"Error al calcular dimensiones: {e}")
        return Inches(max_width_in), Inches(max_height_in)
