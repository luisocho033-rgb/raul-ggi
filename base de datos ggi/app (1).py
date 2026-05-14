import sqlite3
import os
import re
from pathlib import Path
import groq
import pandas as pd
import streamlit as st

# --- CONFIGURACION DE RUTAS Y TABLAS ---
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

    posibles_ids     = [c for c in df.columns if any(x in c for x in ["ced", "ci", "p00", "id", "identific"])]
    posibles_nombres = [c for c in df.columns if any(x in c for x in ["nom", "trabajador", "empleado"])]
    posibles_fechas  = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos  = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status  = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()

    if posibles_fechas:
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors="coerce").dt.strftime("%Y-%m-%d")
    if posibles_ids:
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False).str.strip()
    if posibles_nombres:
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str).str.strip()
    if posibles_deptos:
        df_final["departamento"] = df[posibles_deptos[0]].astype(str).str.strip()
    if posibles_status:
        df_final["estatus"] = df[posibles_status[0]].astype(str).str.strip()

    if not df_final.empty:
        df_final.to_sql(TABLE_NAME, conn, if_exists="append", index=False)
        conn.commit()
        return len(df_final)
    return 0


def limpiar_sql(raw: str) -> str:
    cleaned = re.sub(r"```(?:sql)?", "", raw, flags=re.IGNORECASE)
    cleaned = cleaned.replace(";", "").strip()
    return cleaned


def obtener_muestra_db(conn) -> str:
    """Devuelve info del estado actual de la DB para incluir en el prompt."""
    try:
        total = pd.read_sql(f"SELECT COUNT(*) as total FROM {TABLE_NAME}", conn).iloc[0, 0]
        fechas = pd.read_sql(f"SELECT DISTINCT fecha FROM {TABLE_NAME} ORDER BY fecha DESC LIMIT 10", conn)
        estatuses = pd.read_sql(f"SELECT DISTINCT estatus FROM {TABLE_NAME} LIMIT 20", conn)
        return (
            f"La tabla '{TABLE_NAME}' tiene {total} registros. "
            f"Fechas disponibles (ultimas 10): {fechas['fecha'].tolist()}. "
            f"Valores de estatus existentes: {estatuses['estatus'].tolist()}."
        )
    except Exception:
        return f"La tabla '{TABLE_NAME}' existe pero puede estar vacia."


def main():
    st.set_page_config(page_title="Gestion ggi", layout="wide")
    st.title("Gestion de Asistencia Masiva para ggi")

    conn = init_db()
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # --- BARRA LATERAL ---
    with st.sidebar:
        st.header("Carga de Datos")
        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
        if uploaded_file:
            nombre_sugerido = uploaded_file.name.replace(".xlsx", "")
            nombre_id = st.text_input("Nombre para identificar este archivo", value=nombre_sugerido)

            if st.button("Guardar y Cargar a DB"):
                path = UPLOAD_DIR / f"{nombre_id}.xlsx"
                with open(path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    regs = procesar_flexible(path, conn)
                    st.success(f"Exito! {regs} registros procesados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.header("Archivos Guardados")
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        for arc in archivos:
            col_name, col_del = st.columns([4, 1])
            col_name.text(f"{arc.stem}")
            if col_del.button("X", key=str(arc)):
                arc.unlink()
                st.rerun()

        # Mostrar estado de la DB en sidebar
        st.markdown("---")
        st.header("Estado de la DB")
        try:
            total = pd.read_sql(f"SELECT COUNT(*) as total FROM {TABLE_NAME}", conn).iloc[0, 0]
            st.metric("Registros totales", total)
        except Exception:
            st.info("DB vacia")

    # --- VISTA PRINCIPAL ---
    st.subheader("Resumen General de Novedades")
    try:
        query_view = (
            f"SELECT fecha, identificador, "
            f"CASE WHEN LENGTH(identificador) = 6 THEN 'P00' ELSE 'CI' END AS tipo, "
            f"nombre_trabajador, departamento, estatus "
            f"FROM {TABLE_NAME} "
            f"WHERE estatus LIKE '%Vacacion%' "
            f"OR estatus LIKE '%Teletrabajo%' "
            f"OR estatus LIKE '%Ausente%' "
            f"OR estatus LIKE '%Falta%' "
            f"OR estatus LIKE '%Permiso%' "
            f"ORDER BY fecha DESC LIMIT 100"
        )
        data_view = pd.read_sql(query_view, conn)
        if not data_view.empty:
            st.dataframe(data_view, use_container_width=True)
        else:
            st.info("No hay novedades detectadas. Asegurate de haber cargado un Excel.")
    except Exception as e:
        st.error(f"Error en tabla: {e}")

    # --- CHAT CON IA ---
    st.markdown("---")
    st.subheader("Consultar con IA")

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if p := st.chat_input("Ej: quienes faltaron el 2024-04-24, cuantos teletrabajaron esta semana..."):
        st.session_state.messages.append({"role": "user", "content": p})
        with st.chat_message("user"):
            st.markdown(p)

        if not client:
            st.error("No se encontro GROQ_API_KEY en secrets.")
        else:
            try:
                # Obtener contexto real de la DB
                contexto_db = obtener_muestra_db(conn)

                # 1. Generar SQL con contexto real de la DB
                sys_sql = f"""Eres un experto en SQL para SQLite. Tu unica tarea es generar una consulta SQL SELECT valida.
NUNCA respondas con texto explicativo, NUNCA uses markdown, NUNCA uses bloques de codigo.
Responde UNICAMENTE con la sentencia SQL, sin punto y coma al final.

Base de datos disponible:
- Tabla: {TABLE_NAME}
- Columnas: fecha (TEXT, formato YYYY-MM-DD), identificador (TEXT), nombre_trabajador (TEXT), departamento (TEXT), estatus (TEXT)
- {contexto_db}

Si el usuario pregunta por ausentes/faltas/inasistencias, busca en estatus valores como 'Ausente', 'Falta', 'Inasistencia' o similares.
Si el usuario pregunta por una fecha especifica, usa WHERE fecha = 'YYYY-MM-DD'.
Si el usuario menciona un dia sin año, asume el año mas reciente disponible en la tabla.
"""
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": sys_sql},
                        {"role": "user", "content": p}
                    ],
                    temperature=0,
                    max_tokens=300
                )

                raw_sql = res_sql.choices[0].message.content.strip()
                sql_final = limpiar_sql(raw_sql)

                # Validar que sea un SELECT
                if not sql_final.upper().startswith("SELECT"):
                    st.warning(f"La IA no genero un SELECT valido. SQL recibido: `{sql_final}`")
                    st.stop()

                # 2. Ejecutar consulta
                df_res = pd.read_sql(sql_final, conn)

                # Mostrar tabla de resultados
                if not df_res.empty:
                    st.dataframe(df_res, use_container_width=True)
                else:
                    st.info("La consulta no devolvio resultados. Verifica que los datos esten cargados.")

                # 3. Respuesta en lenguaje natural
                sys_ia = (
                    "Eres un analista de RRHH de la empresa ggi. "
                    "Tienes acceso a datos reales de asistencia. "
                    "Identificadores de 6 digitos son codigos P00, los demas son cedulas (CI). "
                    "Responde de forma directa y concisa en espanol basandote SOLO en los datos proporcionados."
                )
                datos_json = df_res.to_json(orient="records", force_ascii=False)
                user_ia = f"Pregunta del usuario: {p}\n\nDatos obtenidos de la base de datos ({len(df_res)} registros):\n{datos_json}"

                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": sys_ia},
                        {"role": "user", "content": user_ia}
                    ],
                    max_tokens=500
                )

                respuesta = res_ia.choices[0].message.content
                with st.chat_message("assistant"):
                    st.markdown(respuesta)
                st.session_state.messages.append({"role": "assistant", "content": respuesta})

            except Exception as e:
                st.error(f"Error: {e}")
                st.info(f"SQL intentado: `{sql_final if 'sql_final' in locals() else 'no generado'}`")


if __name__ == "__main__":
    main()    posibles_nombres = [c for c in df.columns if any(x in c for x in ["nom", "trabajador", "empleado"])]
    posibles_fechas = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()

    if posibles_fechas:
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors="coerce").dt.strftime("%Y-%m-%d")
    if posibles_ids:
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False).str.strip()
    if posibles_nombres:
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str).str.strip()
    if posibles_deptos:
        df_final["departamento"] = df[posibles_deptos[0]].astype(str).str.strip()
    if posibles_status:
        df_final["estatus"] = df[posibles_status[0]].astype(str).str.strip()

    if not df_final.empty:
        df_final.to_sql(TABLE_NAME, conn, if_exists="append", index=False)
        return len(df_final)
    return 0

def limpiar_sql(raw: str) -> str:
    """Elimina bloques markdown del SQL generado por la IA."""
    cleaned = re.sub(r"```(?:sql)?", "", raw, flags=re.IGNORECASE)
    cleaned = cleaned.replace(";", "").strip()
    return cleaned

def main():
    st.set_page_config(page_title="Gestion ggi", layout="wide")
    st.title("Gestion de Asistencia Masiva para ggi")

    conn = init_db()
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # --- BARRA LATERAL: GESTION DE ARCHIVOS ---
    with st.sidebar:
        st.header("Carga de Datos")
        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
        if uploaded_file:
            nombre_sugerido = uploaded_file.name.replace(".xlsx", "")
            nombre_id = st.text_input("Nombre para identificar este archivo", value=nombre_sugerido)

            if st.button("Guardar y Cargar a DB"):
                path = UPLOAD_DIR / f"{nombre_id}.xlsx"
                with open(path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    regs = procesar_flexible(path, conn)
                    st.success(f"Exito! {regs} registros procesados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.header("Archivos Guardados")
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        for arc in archivos:
            col_name, col_del = st.columns([4, 1])
            col_name.text(f"[doc] {arc.name}")
            if col_del.button("X", key=str(arc)):
                arc.unlink()
                st.rerun()

    # --- VISTA PRINCIPAL ---
    st.subheader("Resumen General de Novedades")
    try:
        query_view = (
            f"SELECT identificador, "
            f"CASE WHEN LENGTH(identificador) = 6 THEN 'P00' ELSE 'CI' END AS tipo, "
            f"nombre_trabajador, estatus, fecha "
            f"FROM {TABLE_NAME} "
            f"WHERE estatus LIKE '%Vacaciones%' "
            f"OR estatus LIKE '%Teletrabajo%' "
            f"OR estatus LIKE '%Ausente%' "
            f"ORDER BY fecha DESC LIMIT 50"
        )
        data_view = pd.read_sql(query_view, conn)
        if not data_view.empty:
            st.dataframe(data_view, use_container_width=True)
        else:
            st.info("No hay novedades detectadas.")
    except Exception as e:
        st.error(f"Error en tabla: {e}")

    # --- CHAT CON IA ---
    st.markdown("---")
    st.subheader("Consultar con IA")

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if p := st.chat_input("Consulta algo..."):
        st.session_state.messages.append({"role": "user", "content": p})
        with st.chat_message("user"):
            st.markdown(p)

        if client:
            try:
                # 1. Generar SQL
                sys_sql = (
                    "Responde SOLO con SQL SELECT. "
                    "Tabla: asistencia_diaria. "
                    "Columnas: fecha, identificador, nombre_trabajador, departamento, estatus."
                )
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": sys_sql},
                        {"role": "user", "content": p}
                    ],
                    temperature=0
                )

                # 2. Limpiar SQL usando la funcion segura (sin strings con saltos de linea literales)
                raw_sql = res_sql.choices[0].message.content.strip()
                sql_final = limpiar_sql(raw_sql)

                # 3. Ejecutar consulta
                df_res = pd.read_sql(sql_final, conn)

                # 4. Generar respuesta en lenguaje natural
                sys_ia = "Analista de ggi. 6 digitos = P00, mas de 6 = CI. Se directo."
                user_ia = f"Pregunta: {p}\nDatos: {df_res.to_json(orient='records')}"
                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": sys_ia},
                        {"role": "user", "content": user_ia}
                    ]
                )

                respuesta = res_ia.choices[0].message.content
                with st.chat_message("assistant"):
                    st.markdown(respuesta)
                st.session_state.messages.append({"role": "assistant", "content": respuesta})

            except Exception as e:
                st.error(f"Consulta fallida: {e}")

if __name__ == "__main__":
    main()
