import io
import re
import warnings
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Consolidador de planillas", layout="wide")
st.title("Consolidador de planillas (Excel / CSV)")

st.write(
    "Sube uno o más archivos .xlsx/.xls/.csv. "
    "Se leerá la hoja **BBDD** (o la primera si no existe). "
    "Los encabezados se **normalizan suavemente** (trim + colapso de espacios) "
    "y se consolidan en un Excel final con los encabezados que usan tus planillas de entrada."
)

# ===== Encabezados canónicos (alineados a tus archivos de entrada) =====
HEADERS = [
    "NUMERO OBRA ICONSTRUYE",
    "MES (MMM-AA)",
    "APELLIDO PATERNO, APELLIDO MATERNO Y NOMBRES DEL TRABAJADOR",
    "RUT TRABAJADOR (SIN PUNTOS Y CON GUION)",
    "DIAS TRABAJADOS ",
    "NUMERO DE CONTRATO",
    "RAZON SOCIAL EMPRESA SUBCONTRATISTA ",
    "RUT EMPRESA SUBCONTRATISTA (SIN PUNTOS Y CON GUION)",
    "RAZON SOCIAL EMPRESA CONTRATISTA ",
    "RUT EMPRESA CONTRATISTA (SIN PUNTOS Y CON GUION)",
]
ORDERED_COLS = ["File name"] + HEADERS

# ===== Opciones en Sidebar =====
with st.sidebar:
    st.header("Opciones")
    prefer_sheet = st.text_input("Nombre de hoja preferida (Excel)", value="BBDD")
    allow_sheet_picker = st.checkbox("Permitir elegir otra hoja si no existe", value=False)

    st.subheader("Lectura de CSV")
    csv_sep = st.text_input("Separador CSV", value=",")
    csv_encoding = st.selectbox("Encoding CSV", ["utf-8", "latin-1", "utf-16"], index=0)

    st.subheader("Vista previa")
    show_preview_rows = st.slider("Filas a mostrar", min_value=5, max_value=200, value=50)

uploaded_files = st.file_uploader(
    "Selecciona tus archivos", type=["xlsx", "xls", "csv"], accept_multiple_files=True
)

# ===== Utilidades =====
def normalize_header(name: str) -> str:
    """Trim y colapso de espacios internos (sin cambiar el texto base)."""
    if not isinstance(name, str):
        return ""
    return " ".join(name.strip().split())

def trim_text_df(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]) or df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()
    return df

def excel_to_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.read()

def read_excel_sheet(file) -> pd.DataFrame:
    """Lee la hoja preferida 'BBDD' si existe; si no, la primera o una elegida (opcional)."""
    xls = pd.ExcelFile(file)
    sheet = prefer_sheet if prefer_sheet in xls.sheet_names else None
    if sheet is None:
        if allow_sheet_picker:
            sheet = st.selectbox(
                f"Elige hoja para **{getattr(file, 'name', 'archivo')}**",
                xls.sheet_names,
                key=f"sheet_{getattr(file, 'name', id(file))}",
            )
        else:
            sheet = xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet, dtype=str)

if uploaded_files:
    st.info("Procesando archivos…")
    warnings_log = []
    df_out = pd.DataFrame(columns=HEADERS)
    df_out["File name"] = ""

    progress = st.progress(0.0)
    status = st.empty()

    for i, up in enumerate(uploaded_files, start=1):
        name = up.name
        try:
            # 1) Leer
            if name.lower().endswith(".csv"):
                df = pd.read_csv(
                    up, dtype=str, skip_blank_lines=True, encoding=csv_encoding, sep=csv_sep, engine="python"
                )
            else:
                df = read_excel_sheet(up)

            # 2) Normalizar encabezados suaves
            norm_map = {c: normalize_header(c) for c in df.columns}
            df = df.rename(columns=norm_map)

            # 3) También normalizamos nuestras HEADERS para chequear presencia
            norm_headers = [normalize_header(h) for h in HEADERS]

            # 4) Seleccionar columnas presentes (según nombre normalizado)
            present_norm = [h for h in norm_headers if h in df.columns]
            # Backmap para volver al nombre “canónico” original de HEADERS
            backmap = {normalize_header(h): h for h in HEADERS}

            tmp = pd.DataFrame()
            for c in present_norm:
                tmp[backmap[c]] = df[c]

            # 5) Crear columnas faltantes vacías
            missing = [h for h in HEADERS if h not in tmp.columns]
            for m in missing:
                tmp[m] = ""

            # 6) Nombre de archivo + orden final
            tmp["File name"] = name
            tmp = tmp[ORDERED_COLS]

            # 7) Limpiar textos y filtrar filas sin NUMERO OBRA ICONSTRUYE
            tmp = trim_text_df(tmp)
            before = len(tmp)
            mask_valid = tmp["NUMERO OBRA ICONSTRUYE"].notna() & (tmp["NUMERO OBRA ICONSTRUYE"] != "")
            tmp = tmp[mask_valid]
            removed = before - len(tmp)
            if removed > 0:
                warnings_log.append(f"{name}: {removed} filas descartadas sin NUMERO OBRA ICONSTRUYE")

            # 8) Concatenar
            df_out = pd.concat([df_out, tmp], ignore_index=True)

            # 9) Advertencias de columnas faltantes
            if missing:
                warnings_log.append(f"{name}: faltan columnas -> {', '.join(missing)}")

        except Exception as exn:
            warnings_log.append(f"{name}: ERROR al procesar -> {exn}")

        progress.progress(i / len(uploaded_files))
        status.write(f"Procesado: **{i}/{len(uploaded_files)}**")

    st.success("¡Listo! Consolidación finalizada.")
    st.subheader("Vista previa")
    st.dataframe(df_out.head(show_preview_rows), use_container_width=True, height=420)

    # Descarga de Excel final
    st.download_button(
        "⤓ Descargar Excel consolidado",
        data=excel_to_bytes(df_out),
        file_name="unificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Descarga del log
    if warnings_log:
        st.warning("Se generaron advertencias/errores durante el proceso.")
        st.download_button(
            "⤓ Descargar log de advertencias",
            data=("\n".join(warnings_log)).encode("utf-8"),
            file_name="log_consolidacion.txt",
            mime="text/plain",
        )
else:
    st.info("Carga tus archivos para comenzar. Ajusta opciones en el panel izquierdo si es necesario.")
