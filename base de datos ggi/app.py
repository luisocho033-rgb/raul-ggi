import json
import re
import sqlite3
import os
from pathlib import Path
from datetime import date

import groq
import pandas as pd
import streamlit as st

# Configuración de rutas
DB_PATH = Path("asistencia.db")
UPLOAD_DIR = Path("archivos_excel")
UPLOAD_DIR.mkdir(exist_ok=True) # Crea la carpeta si no existe

TABLE_NAME = "asistencia_diaria"

# Se unifica bajo "identificador" para agrupar P00 y CI
COLUMN_ALIASES = {
    "fecha": ["fecha", "date", "dia", "día"],
    "identificador": ["cedula_trabajador", "cedula", "id", "identificacion", "identificación", "ci", "dni", "c.i", "p00"],
    "nombre_trabajador": ["nombre_trabajador", "nombre", "nombre_completo", "trabajador", "personal", "empleado"],
    "departamento": ["departamento", "area", "área", "division", "división", "sector"],
    "estatus": ["estatus", "estado", "asistencia", "situacion", "situación", "motivo", "ausencia", "observacion", "observación"],
}
REQUIRED_FIELDS = list(COLUMN_ALIASES.keys())
AUSENCIAS = ("Vacaciones", "Teletrabajo", "Ausente")

def init_db(path: Path = DB_PATH) -> sqlite3.Connection:
    """Inicializa la base de datos SQLite y crea la tabla de asistencia."""
    conn = sqlite3.connect(path, check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            fecha TEXT NOT NULL,
            identificador TEXT NOT NULL,
            nombre_trabajador TEXT NOT NULL,
            departamento TEXT NOT NULL,
            estatus TEXT NOT NULL,
            PRIMARY KEY(fecha, identificador)
        )
        """
    )
    conn.commit()
    return conn

def guess_column_by_content(df):
    """Detecta las columnas mapeando alias o analizando el contenido."""
    result = {}
    lowered = {col.lower().strip(): col for col in df.columns}
    
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lowered:
                result[field] = lowered[alias]
                break

    for field in REQUIRED_FIELDS:
        if field not in result:
            candidates = []
            for col in df.columns:
                s = df[col]
                if field == "fecha":
                    num_dates = pd.to_datetime(s, errors="coerce").notna().sum()
                    if num_dates > len(s) // 4:
                        candidates.append((col, num_dates))
                elif field == "identificador":
                    mask = s.astype(str).str.match(r"^\d{6,9}$", na=False)
                    if mask.sum() > len(s) // 2:
                        candidates.append((col, mask.sum()))
                elif field == "nombre_trabajador":
                    ntext = s.astype(str).str.count(r" ").sum()
                    if ntext > len(s) // 5:
                        candidates.append((col, ntext))
                elif field == "departamento":
                    nunique = s.nunique()
                    if nunique <= len(s) // 5:
                        candidates.append((col, nunique))
                elif field == "estatus":
                    nunique = s.nunique()
                    if nunique <= 12:
                        candidates.append((col, nunique))
            if candidates:
                result[field] = sorted(candidates, key=lambda x: x[1], reverse=True)[0][0]
                
    missing = [f for f in REQUIRED_FIELDS if f not in result]
    if missing:
        st.warning("Faltan las siguientes columnas esenciales: " + ", ".join(missing))
    return result

def procesar_excel(file_path, conn):
    """Procesa el archivo Excel físico y lo inserta en SQLite."""
    df = pd.read_excel(file_path, engine="openpyxl")
    colmaps = guess_column_by_content(df)
    dfout = pd.DataFrame()
    
    for std, orig in colmaps.items():
        dfout[std] = df[orig]

    dfout["fecha"] = pd.to_datetime(dfout["fecha"], errors="coerce").dt.date
    dfout["identificador"] = dfout["identificador"].astype(str).str.strip()
    dfout["nombre_trabajador"] = dfout["nombre_trabajador"].astype(str).str.strip()
    dfout["departamento"] = dfout["departamento"].astype(str).str.strip()
    dfout["estatus"] = dfout["estatus"].astype(str).str.strip().str.title()

    dfout = dfout.dropna(subset=REQUIRED_FIELDS)
    dfout["fecha"] = dfout["fecha"].astype(str)

    values = [tuple(row) for row in dfout.to_numpy(dtype=object)]
    cursor = conn.cursor()
    prior_changes = conn.total_changes
    cursor.executemany(
        f"""
        INSERT OR IGNORE INTO {TABLE_NAME}
        (fecha, identificador, nombre_trabajador, departamento, estatus)
        VALUES (?, ?, ?, ?, ?)
        """,
        values,
    )
    conn.commit()
    inserted = conn.total_changes - prior_changes
    return inserted, len(values)

def init_groq_client() -> groq.Client:
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("La clave GROQ_API_KEY no está configurada en st.secrets.")
    return groq.Client(api_key=api_key)

def generar_sql(prompt: str, client: groq.Client) -> str:
    system_prompt = (
        "Eres un experto en SQL SQLite. "
        "La tabla 'asistencia_diaria' tiene: fecha, identificador, nombre_trabajador, departamento, estatus. "
        "REGLA VITAL: La columna 'identificador' contiene tanto el P00 (si tiene exactamente 6 caracteres numéricos) "
        "como la Cédula/CI (si tiene más de 6 caracteres numéricos). "
        "Genera solo código SQL válido SELECT o WITH. No incluyas explicaciones."
    )
    user_prompt = f"Pregunta: {prompt}\nGenera solo la consulta SQL."
    full_prompt = f"{system_prompt}\n\n{user_prompt}"

    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[{"role": "user", "content": full_prompt}],
        max_tokens=256,
        temperature=0.0,
    )
    
    return response.choices[0].message.content.strip().rstrip(";")

def validar_sql(sql: str) -> None:
    if not sql:
        raise ValueError("Consulta vacía.")
    if ";" in sql:
        raise ValueError("No usar punto y coma.")
    if not re.match(r"^(SELECT|WITH)\b", sql.strip(), re.IGNORECASE):
        raise ValueError("Sólo SELECT o WITH permitidos.")
    prohibited = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "ATTACH", "DETACH", "PRAGMA"]
    for token in prohibited:
        if re.search(rf"\b{token}\b", sql, re.IGNORECASE):
            raise ValueError(f"Instrucción SQL prohibida: {token}.")

def ejecutar_sql(conn: sqlite3.Connection, sql: str) -> pd.DataFrame:
    validar_sql(sql)
    cursor = conn.cursor()
    cursor.execute(sql)
    columns = [desc[0] for desc in cursor.description] if cursor.description else []
    return pd.DataFrame(cursor.fetchall(), columns=columns)

def resumir_resultado(question: str, sql: str, df: pd.DataFrame, client: groq.Client) -> str:
    if df.empty:
        return "No se encontraron registros para esta consulta."
    
    records = df.to_dict(orient="records")
    result_text = json.dumps(records[:20], ensure_ascii=False)

    prompt = (
        f"Pregunta del usuario: {question}\n"
        f"Datos extraídos: {result_text}\n"
        "Responde en español de forma analítica y directa. "
        "Recuerda que identificadores de 6 dígitos son P00 y los de más de 6 dígitos son CI. "
        "No muestres el SQL."
    )
    
    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=400,
        temperature=0.3,
    )
    
    return response.choices[0].message.content.strip()

def resumen_ausencias(conn: sqlite3.Connection):
    """Muestra la tabla de ausencias procesadas de la DB."""
    query = f"""
    SELECT 
        identificador, 
        CASE 
            WHEN LENGTH(identificador) = 6 THEN 'P00' 
            ELSE 'CI' 
        END AS tipo_id,
        nombre_trabajador, 
        estatus AS motivo, 
        fecha
    FROM {TABLE_NAME}
    WHERE estatus IN {AUSENCIAS}
    ORDER BY identificador, fecha
    """
    try:
        df = pd.read_sql_query(query, conn)
        if df.empty:
            return

        df['fecha'] = pd.to_datetime(df['fecha'])
        df['grupo'] = (df['fecha'].diff().dt.days.ne(1) | 
                       (df['identificador'] != df['identificador'].shift(1)) | 
                       (df['motivo'] != df['motivo'].shift(1))).cumsum()
        resumen = (
            df.groupby(['identificador', 'tipo_id', 'nombre_trabajador', 'motivo', 'grupo'])
            .agg(fecha_inicio=('fecha', 'min'), fecha_fin=('fecha', 'max'))
            .reset_index()
            .drop('grupo', axis=1)
            .sort_values(['fecha_inicio'])
        )
        resumen['fecha_inicio'] = resumen['fecha_inicio'].dt.strftime('%Y-%m-%d')
        resumen['fecha_fin'] = resumen['fecha_fin'].dt.strftime('%Y-%m-%d')

        st.dataframe(resumen, use_container_width=True)
    except Exception as e:
        st.error(f"Error interno en tabla visual: {e}")

def main() -> None:
    st.set_page_config(page_title="Gestión de Asistencia ggi", layout="wide")
    st.title("Gestión de Asistencia Masiva para ggi")

    conn = init_db()
    client = init_groq_client()

    # --- ZONA DE GESTIÓN DE ARCHIVOS ---
    with st.expander("Subir y Guardar Archivo Excel", expanded=True):
        uploaded_file = st.file_uploader("Sube el archivo (.xlsx)", type=["xlsx"])
        
        if uploaded_file is not None:
            custom_name = st.text_input("Asigna un nombre para identificar este archivo (sin .xlsx)", value=uploaded_file.name.replace(".xlsx", ""))
            
            if st.button("Guardar y Procesar en Base de Datos"):
                if custom_name.strip() == "":
                    st.error("El nombre no puede estar vacío.")
                else:
                    file_path = UPLOAD_DIR / f"{custom_name}.xlsx"
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    try:
                        inserted, total = procesar_excel(file_path, conn)
                        if inserted > 0:
                            st.success(f"Guardado. Se importaron {inserted} registros nuevos de {total} procesados.")
                        else:
                            st.info("Archivo guardado. No hay registros nuevos (ya existían).")
                    except Exception as exc:
                        st.error(f"Error al procesar: {exc}")
                        file_path.unlink(missing_ok=True) # Revierte si hay error

    with st.expander("Archivos Guardados en el Servidor", expanded=False):
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        if not archivos:
            st.write("No hay archivos guardados.")
        else:
            for archivo in archivos:
                col1, col2 = st.columns([4, 1])
                col1.write(f"📄 {archivo.name}")
                if col2.button("Borrar", key=archivo.name):
                    archivo.unlink()
                    st.rerun()

    # --- TABLA DE DATOS ---
    resumen_ausencias(conn)

    # --- CHAT IA ---
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    st.markdown("---")
    st.subheader("Consultas de Asistencia (IA)")

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.write(message["content"])

    user_prompt = st.chat_input("Consulta algo. Ejemplo: ¿Cuántas ausencias tiene el P00 123456?")
    if user_prompt:
        st.session_state.chat_history.append({"role": "user", "content": user_prompt})
        with st.chat_message("assistant"):
            st.write("Analizando...")

        try:
            sql_query = generar_sql(user_prompt, client)
            df_result = ejecutar_sql(conn, sql_query)
            answer = resumir_resultado(user_prompt, sql_query, df_result, client)
            st.session_state.chat_history.append({"role": "assistant", "content": answer})
            st.rerun()
        except Exception as exc:
            st.session_state.chat_history.append({"role": "assistant", "content": f"Fallo en consulta: {exc}"})
            st.rerun()

    st.sidebar.header("Estado")
    st.sidebar.write(f"DB: {DB_PATH.name}")
    st.sidebar.write(f"Directorio: {UPLOAD_DIR.name}/")
    st.sidebar.write(f"Fecha: {date.today().isoformat()}")

if __name__ == "__main__":
    main()
