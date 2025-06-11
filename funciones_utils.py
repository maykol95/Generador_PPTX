# funciones_utils.py
import os
import requests
import tempfile
import pandas as pd


def detectar_columna_imagenes(df):
    for col in df.columns:
        if df[col].dropna().astype(str).str.contains(r'(https?://|\.jpg|\.png|\.jpeg|^[\w-]{25,}$)', case=False).any():
            return col
    return None

def descargar_imagenes_temp(df, columna_img, temp_dir):
    """
    Descarga imágenes desde URLs en la columna columna_img de df
    y guarda los archivos temporalmente en temp_dir.
    Devuelve una serie con las rutas locales, alineada con df.index.
    """
    img_paths = []
    for url in df[columna_img]:
        try:
            if not isinstance(url, str) or url.strip() == "":
                img_paths.append(None)
                continue

            filename = os.path.basename(url).split("?")[0]
            local_path = os.path.join(temp_dir, filename)

            # Evitar descargar si ya existe
            if not os.path.exists(local_path):
                r = requests.get(url, stream=True)
                if r.status_code == 200:
                    with open(local_path, "wb") as f:
                        for chunk in r.iter_content(1024):
                            f.write(chunk)
                else:
                    local_path = None
            img_paths.append(local_path)
        except Exception as e:
            print(f"Error descargando {url}: {e}")
            img_paths.append(None)

    # Devuelve serie alineada con índice original
    return pd.Series(img_paths, index=df.index)

def convertir_columnas_a_str(df):
    return df.astype(str)

def filtrar_df_por_imagenes(df, columna_img):
    return df[df[columna_img].notna() & df[columna_img].astype(str).str.strip().ne("")]
