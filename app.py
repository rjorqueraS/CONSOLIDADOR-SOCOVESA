import io
import warnings
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Unificador de planillas", layout="wide")

st.title("Unificador de planillas (Excel/CSV)")
st.write(
    "Sube uno o más archivos .xlsx / .xls / .csv. "
    "La app leerá la hoja **BBDD** (o la primera si no existe), "
    "mantendrá sólo las columnas requeridas y generará un Excel final."
)

# === Definición de columnas esperadas (como en tu script) ===
HEADERS = {
    "NUMERO OBRA ICONSTRUYE": str,
    "MES (MMM-AA)": str,
    "APELLIDO     PATERNO ": str,
    "APELLIDO MATERNO": str,
    "NOMBRES ": str,
    "RUT TRABAJADOR (SIN PUNTOS Y CON GUION)": str,
    "DIAS TRABAJADOS ": str,
    "NUMERO DE CONTRATO ICONSTRUYE": str,
    "NOMBRE SUBCONTRATO SERRES (RAZON SOCIAL EMPRESA)": str,
    "NOMBRE SUBCONTRATO ICONSTRUYE": str,
    "RUT SUBCONTRATO (SIN PUNTOS Y CON GUION)": str,
}
ORDERED_COLS = ["File name"] + list(HEADERS.keys())

uploaded_files = st.file_uploader(
    "Selecciona tus archivos", type=["xlsx", "xls", "csv"], accept_multiple_files=True
)

if uploaded_files:
    failed = []
    # DataFrame de salida
    df_out = pd.DataFrame(columns=list(HEADERS.keys()))
    df_out["File name"] = ""

    progress = st.progress(0.0)
    status = st.empty()

    for i, up in enumerate(uploaded_files, start=1):
        try:
            name = up.name.lower()
            if name.endswith(".csv"):
                # Leer CSV (todo como texto)
                file_df = pd.read_csv(
                    up, dtype=str, skip_blank_lines=True, encoding="utf-8", engine="python")
            else:
                # Leer Excel: primero intenta hoja 'BBDD', si no existe usa la primera
                try:
                    file_df = pd.read_excel(up, sheet_name="BBDD", dtype=str)
                except Exception:
                    file_df = pd.read_excel(up, sheet_name=0, dtype=str)

            # Mantener sólo columnas presentes en HEADERS
            cols_presentes = list(set(HEADERS.keys()) & set(file_df.columns))
            file_df = file_df[cols_presentes].copy()

            # Nombre de archivo
            file_df["File name"] = up.name

            # Concatenar
            df_out = pd.concat([df_out, file_df], ignore_index=True)

        except Exception as exn:
            failed.append(f"{up.name}: {exn}")

        # Actualizar progreso
        progress.progress(i / len(uploaded_files))
        status.write(f"Procesado: **{i}/{len(uploaded_files)}**")

    # Eliminar filas sin 'NUMERO OBRA ICONSTRUYE'
    if "NUMERO OBRA ICONSTRUYE" in df_out.columns:
        df_out = df_out[~df_out["NUMERO OBRA ICONSTRUYE"].isna()]

    # Reordenar columnas para salida
    for col in ORDERED_COLS:
        if col not in df_out.columns:
            df_out[col] = ""
    df_out = df_out[ORDERED_COLS]

    st.success("¡Listo! Archivos procesados.")
    st.subheader("Vista previa")
    st.dataframe(df_out, use_container_width=True, height=400)

    # === Botón para descargar Excel final ===
    def make_excel_bytes(df: pd.DataFrame) -> bytes:
        buffer = io.BytesIO()
        # Usa xlsxwriter si está disponible; pandas elegirá engine automáticamente si tienes openpyxl
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)
        return buffer.read()

    xlsx_bytes = make_excel_bytes(df_out)
    st.download_button(
        label="⤓ Descargar Excel combinado (.xlsx)",
        data=xlsx_bytes,
        file_name="unificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # === Log de errores (opcional) ===
    if failed:
        st.warning("Algunos archivos tuvieron errores. Puedes descargar el log.")
        log_text = "\n".join(failed).encode("utf-8")
        st.download_button(
            label="⤓ Descargar log de errores (.txt)",
            data=log_text,
            file_name="errores.txt",
            mime="text/plain",
        )

else:
    st.info("Carga tus archivos para comenzar.")
