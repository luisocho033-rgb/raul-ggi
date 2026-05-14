import sqlite3
import os
import re
from pathlib import Path
import groq
import pandas as pd
import streamlit as st

# --- CONFIGURACION ---
DB_PATH = Path("asistencia.db")
UPLOAD_DIR = Path("archivos_excel")
UPLOAD_DIR.mkdir(exist_ok=True)
TABLE_NAME = "asistencia_diaria"

def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            fecha TEXT,
            identificador TEXT,
            nombre_trabajador TEXT,
            departamento TEXT,
            estatus TEXT,
            PRIMARY KEY(fecha, identificador)
        )
    """)
    conn.commit()
    return conn

def procesar_flexible(file_path, conn):
    df = pd.read_excel(file_path, engine="openpyxl")
    df.columns = [c.lower().strip() for c in df.columns]

    def buscar(keys):
        for c in df.columns:
            if any(k in c for k in keys):
                return c
        return None

    col_id     = buscar(["ced", "ci", "p00", "identific", " id", "num"])
    col_nombre = buscar(["nom", "trabajador", "empleado", "personal"])
    col_fecha  = buscar(["fech", "dia", "date"])
    col_depto  = buscar(["dep", "area", "div", "unidad", "secc"])
    col_status = buscar(["estat", "situ", "motivo", "asist", "nov", "tipo"])

    df_final = pd.DataFrame()
    if col_fecha:  df_final["fecha"]            = pd.to_datetime(df[col_fecha], errors="coerce").dt.strftime("%Y-%m-%d")
    if col_id:     df_final["identificador"]    = df[col_id].astype(str).str.replace(".0","",regex=False).str.strip()
    if col_nombre: df_final["nombre_trabajador"]= df[col_nombre].astype(str).str.strip()
    if col_depto:  df_final["departamento"]     = df[col_depto].astype(str).str.strip()
    if col_status: df_final["estatus"]          = df[col_status].astype(str).str.strip()

    # Si no hay identificador, usar índice para evitar colisiones de PRIMARY KEY
    if "identificador" not in df_final.columns:
        df_final["identificador"] = [f"fila_{i}" for i in range(len(df_final))]
    if "fecha" not in df_final.columns:
        df_final["fecha"] = "sin-fecha"

    n_insertados = 0
    for _, row in df_final.iterrows():
        try:
            conn.execute(
                f"INSERT OR REPLACE INTO {TABLE_NAME} VALUES (?,?,?,?,?)",
                (
                    row.get("fecha", ""),
                    row.get("identificador", ""),
                    row.get("nombre_trabajador", ""),
                    row.get("departamento", ""),
                    row.get("estatus", ""),
                )
            )
            n_insertados += 1
        except Exception:
            pass
    conn.commit()
    return n_insertados


def get_context_data(conn):
    """Lee la DB y construye un resumen real para pasarle a la IA."""
    try:
        df = pd.read_sql(f"SELECT * FROM {TABLE_NAME}", conn)
    except Exception:
        return "No hay datos en la base de datos.", pd.DataFrame()

    if df.empty:
        return "La base de datos está vacía.", df

    total = len(df)

    # Conteo por estatus
    by_status = df["estatus"].value_counts().to_dict()
    status_txt = "\n".join(f"  - {k}: {v}" for k, v in by_status.items())

    # Ausentes
    ausentes = df[df["estatus"].str.contains("Ausent|Falt", case=False, na=False)]
    aus_txt = ausentes[["nombre_trabajador", "identificador", "departamento", "fecha"]].to_string(index=False) if not ausentes.empty else "Ninguno"

    # Teletrabajo
    tele = df[df["estatus"].str.contains("Telet|Telew", case=False, na=False)]
    tele_txt = tele[["nombre_trabajador", "identificador", "departamento", "fecha"]].to_string(index=False) if not tele.empty else "Ninguno"

    # Vacaciones
    vacas = df[df["estatus"].str.contains("Vacac", case=False, na=False)]
    vacas_txt = vacas[["nombre_trabajador", "identificador", "departamento", "fecha"]].to_string(index=False) if not vacas.empty else "Ninguno"

    # Muestra de datos completos (máx 80 filas)
    muestra = df.head(80).to_string(index=False)

    context = f"""
DATOS REALES DE ASISTENCIA GGI
================================
Total de registros en DB: {total}
Fechas disponibles: {sorted(df['fecha'].dropna().unique().tolist())}

CONTEO POR ESTATUS:
{status_txt}

AUSENTES / FALTAS ({len(ausentes)}):
{aus_txt}

TELETRABAJO ({len(tele)}):
{tele_txt}

VACACIONES ({len(vacas)}):
{vacas_txt}

MUESTRA COMPLETA (hasta 80 registros):
{muestra}

NOTA: identificador de 6 dígitos = ficha P00 (ej: 123456 → P00123456). Más de 6 dígitos = cédula (CI).
""".strip()

    return context, df


def main():
    st.set_page_config(page_title="Gestión GGI", layout="wide")
    st.title("📋 Gestión de Asistencia · GGI")

    conn = init_db()
    api_key = st.secrets.get("GROQ_API_KEY", "")
    client = groq.Client(api_key=api_key) if api_key else None

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("📂 Carga de Datos")
        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
        if uploaded_file:
            nombre_id = st.text_input("Nombre para identificar este archivo",
                                      value=uploaded_file.name.replace(".xlsx", ""))
            if st.button("✅ Guardar y Cargar a DB"):
                path = UPLOAD_DIR / f"{nombre_id}.xlsx"
                with open(path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    regs = procesar_flexible(path, conn)
                    st.success(f"¡Listo! {regs} registros cargados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.subheader("📁 Archivos Guardados")
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        if not archivos:
            st.caption("Ningún archivo cargado aún.")
        for arc in archivos:
            col_n, col_d = st.columns([4, 1])
            col_n.text(f"📄 {arc.name}")
            if col_d.button("✕", key=str(arc)):
                arc.unlink()
                st.rerun()

        st.markdown("---")
        if st.button("🗑️ Borrar TODA la base de datos", type="secondary"):
            conn.execute(f"DELETE FROM {TABLE_NAME}")
            conn.commit()
            st.success("Base de datos vaciada.")
            st.rerun()

    # ── VISTA GENERAL ─────────────────────────────────────────────────────────
    context_data, df_full = get_context_data(conn)

    if not df_full.empty:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total registros", len(df_full))
        col2.metric("Ausentes", len(df_full[df_full["estatus"].str.contains("Ausent|Falt", case=False, na=False)]))
        col3.metric("Teletrabajo", len(df_full[df_full["estatus"].str.contains("Telet|Telew", case=False, na=False)]))
        col4.metric("Vacaciones", len(df_full[df_full["estatus"].str.contains("Vacac", case=False, na=False)]))
        st.markdown("---")

    st.subheader("📊 Resumen de Novedades")
    if df_full.empty:
        st.info("No hay datos. Carga un Excel para comenzar.")
    else:
        novedades = df_full[
            df_full["estatus"].str.contains(
                "Vacac|Telet|Telew|Ausent|Falt|Permis|Licen", case=False, na=False
            )
        ].copy()

        novedades["tipo"] = novedades["identificador"].apply(
            lambda x: "P00" if str(x).replace("-","").isdigit() and len(str(x).replace("-","")) == 6 else "CI"
        )

        if novedades.empty:
            st.info("Sin novedades detectadas con los filtros actuales.")
        else:
            st.dataframe(
                novedades[["fecha", "tipo", "identificador", "nombre_trabajador", "departamento", "estatus"]],
                use_container_width=True,
                hide_index=True
            )

        with st.expander("Ver tabla completa"):
            st.dataframe(df_full, use_container_width=True, hide_index=True)

    # ── CHAT CON IA ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🤖 Consultar con IA")

    if not client:
        st.warning("Configura GROQ_API_KEY en los secrets de Streamlit para usar el chat.")
    else:
        if "messages" not in st.session_state:
            st.session_state.messages = []

        for m in st.session_state.messages:
            with st.chat_message(m["role"]):
                st.markdown(m["content"])

        if prompt := st.chat_input("Pregunta algo sobre los datos de asistencia..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            # Re-leer datos frescos antes de responder
            context_data, _ = get_context_data(conn)

            system_prompt = f"""Eres un analista de RRHH de la empresa GGI. Tienes acceso a los datos REALES de asistencia del personal.

REGLAS:
1. SIEMPRE responde basándote en los datos reales que se te proporcionan abajo. NUNCA digas que no tienes acceso a datos.
2. Si te preguntan por ausentes, lista los nombres reales con su ID, departamento y fecha.
3. Si te preguntan cuántos, da el número exacto según los datos.
4. Un identificador de 6 dígitos = ficha P00. Más de 6 dígitos = cédula (CI).
5. Si la columna de estatus usa palabras distintas (ej: "No asistió", "Falta"), igual identifícalos como ausentes.
6. Responde en español, de forma directa y útil. Usa listas cuando hay varios nombres.
7. Si los datos están incompletos o no hay registros para algo, dilo claramente.

DATOS ACTUALES DE LA BASE DE DATOS:
{context_data}"""

            with st.chat_message("assistant"):
                with st.spinner("Analizando datos..."):
                    try:
                        history = [
                            {"role": m["role"], "content": m["content"]}
                            for m in st.session_state.messages
                        ]
                        resp = client.chat.completions.create(
                            model="llama-3.3-70b-versatile",   # modelo mucho más capaz
                            messages=[
                                {"role": "system", "content": system_prompt},
                                *history
                            ],
                            temperature=0.1,
                            max_tokens=1500,
                        )
                        respuesta = resp.choices[0].message.content
                        st.markdown(respuesta)
                        st.session_state.messages.append({"role": "assistant", "content": respuesta})
                    except Exception as e:
                        st.error(f"Error al consultar la IA: {e}")

        if st.session_state.get("messages"):
            if st.button("🗑️ Limpiar chat"):
                st.session_state.messages = []
                st.rerun()

if __name__ == "__main__":
    main()
