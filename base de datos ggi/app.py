import json
import re
import sqlite3
from pathlib import Path
from datetime import date

import groq
import pandas as pd
import streamlit as st

# Configuración de la base de datos SQLite local
DB_PATH = Path("asistencia.db")
TABLE_NAME = "asistencia_diaria"
REQUIRED_COLUMNS = ["fecha", "cedula_trabajador", "nombre_trabajador", "departamento", "estatus"]
COLUMN_ALIASES = {
    "fecha": ["fecha", "date", "dia"],
    "cedula_trabajador": ["cedula_trabajador", "cedula", "id", "identificacion", "identificacion_trabajador"],
    "nombre_trabajador": ["nombre_trabajador", "nombre", "nombre_completo", "trabajador"],
    "departamento": ["departamento", "area", "division"],
    "estatus": ["estatus", "estado", "asistencia", "situacion"],
}


def init_db(path: Path = DB_PATH) -> sqlite3.Connection:
    """Inicializa la base de datos SQLite y crea la tabla de asistencia si no existe."""
    conn = sqlite3.connect(path, check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            fecha TEXT NOT NULL,
            cedula_trabajador TEXT NOT NULL,
            nombre_trabajador TEXT NOT NULL,
            departamento TEXT NOT NULL,
            estatus TEXT NOT NULL,
            PRIMARY KEY(fecha, cedula_trabajador)
        )
        """
    )
    conn.commit()
    return conn


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nombres de columnas del Excel usando alias simples."""
    lowercase = {col.lower().strip(): col for col in df.columns}
    rename_map = {}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lowercase:
                rename_map[lowercase[alias]] = target
                break
    return df.rename(columns=rename_map)


def procesar_excel(uploaded_file, conn: sqlite3.Connection) -> tuple[int, int]:
    """Procesa el archivo Excel y guarda registros únicos en SQLite."""
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df = normalize_columns(df)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"Columnas faltantes en el archivo Excel: {', '.join(missing)}. "
            "Asegúrate de incluir fecha, cedula_trabajador, nombre_trabajador, departamento y estatus."
        )

    df = df[REQUIRED_COLUMNS].copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
    df["cedula_trabajador"] = df["cedula_trabajador"].astype(str).str.strip()
    df["nombre_trabajador"] = df["nombre_trabajador"].astype(str).str.strip()
    df["departamento"] = df["departamento"].astype(str).str.strip()
    df["estatus"] = df["estatus"].astype(str).str.strip().str.title()

    df = df.dropna(subset=["fecha", "cedula_trabajador", "nombre_trabajador", "departamento", "estatus"])
    df["fecha"] = df["fecha"].astype(str)

    values = [tuple(row) for row in df.to_numpy(dtype=object)]
    cursor = conn.cursor()
    prior_changes = conn.total_changes
    cursor.executemany(
        f"""
        INSERT OR IGNORE INTO {TABLE_NAME} (fecha, cedula_trabajador, nombre_trabajador, departamento, estatus)
        VALUES (?, ?, ?, ?, ?)
        """,
        values,
    )
    conn.commit()
    inserted = conn.total_changes - prior_changes
    return inserted, len(values)


def init_groq_client() -> groq.Client:
    """Inicializa el cliente Groq usando la clave en st.secrets."""
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("La clave GROQ_API_KEY no está configurada en st.secrets.")
    return groq.Client(api_key=api_key)


def extract_groq_text(response) -> str:
    """Extrae texto de la respuesta de Groq de forma robusta."""
    if hasattr(response, "output_text") and response.output_text:
        return response.output_text.strip()
    if isinstance(response, dict):
        output = response.get("output")
        if isinstance(output, list) and output:
            first = output[0]
            if isinstance(first, dict):
                content = first.get("content")
                if isinstance(content, list):
                    return "".join(item.get("text", "") for item in content if isinstance(item, dict)).strip()
                if isinstance(content, str):
                    return content.strip()
    return str(response).strip()


def generar_sql(prompt: str, client: groq.Client) -> str:
    """Genera una consulta SQL SELECT para SQLite a partir de la pregunta del usuario."""
    system_prompt = (
        "Eres un asistente experto en SQL SQLite. "
        "La base de datos contiene la tabla asistencia_diaria con columnas: fecha, cedula_trabajador, nombre_trabajador, departamento, estatus. "
        "Devuelve únicamente una consulta SQL válida de tipo SELECT o WITH. "
        "No incluyas explicaciones, no incluyas formato de texto adicional y no uses punto y coma final."
    )
    user_prompt = (
        f"Pregunta: {prompt}\n"
        "Genera solo la consulta SQL."
    )
    full_prompt = f"{system_prompt}\n\n{user_prompt}"

    response = client.responses.create(
        model="llama3-70b-8192",
        input=full_prompt,
        max_tokens=512,
        temperature=0.0,
    )
    sql = extract_groq_text(response)
    sql = sql.strip().rstrip(";")
    return sql


def validar_sql(sql: str) -> None:
    """Valida que la consulta SQL sea segura y sólo permita SELECT/ WITH."""
    if not sql:
        raise ValueError("La consulta SQL generada está vacía.")
    if ";" in sql:
        raise ValueError("La consulta SQL no puede contener punto y coma.")
    if not re.match(r"^(SELECT|WITH)\b", sql.strip(), re.IGNORECASE):
        raise ValueError("Sólo se permiten consultas SELECT o WITH en esta interfaz.")
    prohibited = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "ATTACH", "DETACH", "PRAGMA"]
    for token in prohibited:
        if re.search(rf"\b{token}\b", sql, re.IGNORECASE):
            raise ValueError(f"Consulta SQL no permitida: {token}.")


def ejecutar_sql(conn: sqlite3.Connection, sql: str) -> pd.DataFrame:
    """Ejecuta la consulta SQL y devuelve el resultado en un DataFrame."""
    validar_sql(sql)
    cursor = conn.cursor()
    cursor.execute(sql)
    columns = [description[0] for description in cursor.description] if cursor.description else []
    rows = cursor.fetchall()
    return pd.DataFrame(rows, columns=columns)


def resumir_resultado(question: str, sql: str, df: pd.DataFrame, client: groq.Client) -> str:
    """Envía el resultado de la consulta a Groq para que lo redacte en español con detalle."""
    if df.empty:
        result_text = "[]"
    else:
        records = df.to_dict(orient="records")
        sample = records[:50]
        result_text = json.dumps(sample, ensure_ascii=False, indent=2)
        if len(records) > 50:
            result_text += f"\n... (se muestran los primeros 50 de {len(records)} registros)"

    prompt = (
        "Eres un redactor experto en la información mostrada. "
        "Recibe la pregunta del usuario, la consulta SQL ejecutada y el resultado en JSON. "
        "Responde en español con lujo de detalle, de forma clara y amigable. "
        "No incluyas la consulta SQL en la respuesta final, solo el análisis y la conclusión.\n\n"
        f"Pregunta: {question}\n"
        f"SQL ejecutado: {sql}\n"
        f"Resultados: {result_text}"
    )
    response = client.responses.create(
        model="llama3-70b-8192",
        input=prompt,
        max_tokens=512,
        temperature=0.3,
    )
    return extract_groq_text(response)


def main() -> None:
    """Aplicación Streamlit principal."""
    st.set_page_config(page_title="Gestión de Asistencia RRHH", layout="wide")
    st.title("Gestión de Asistencia Masiva para RRHH")
    st.write(
        "Sube un archivo .xlsx diario para registrar la asistencia, y usa la consulta de chat para analizar la información.")

    conn = init_db()
    client = init_groq_client()

    with st.expander("Carga diaria de asistencia", expanded=True):
        uploaded_file = st.file_uploader(
            "Sube el archivo de asistencia (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=False,
        )
        if uploaded_file is not None:
            try:
                inserted, total = procesar_excel(uploaded_file, conn)
                if inserted > 0:
                    st.success(f"Se importaron {inserted} registros nuevos de {total} filas procesadas.")
                else:
                    st.info(
                        "No se insertaron registros nuevos. El archivo parece corresponder a fechas ya cargadas o los registros ya existían.")
            except Exception as exc:
                st.error(f"Error al procesar el archivo: {exc}")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    st.markdown("---")
    st.subheader("Chat de consultas SQL para asistencia")

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.write(message["content"])

    user_prompt = st.chat_input("Haz una pregunta sobre asistencia, por ejemplo: ¿Quién faltó más veces en el año en infraestructura?")
    if user_prompt:
        st.session_state.chat_history.append({"role": "user", "content": user_prompt})
        with st.chat_message("assistant"):
            st.write("Procesando tu solicitud... esto puede tardar unos segundos.")

        try:
            sql_query = generar_sql(user_prompt, client)
            df_result = ejecutar_sql(conn, sql_query)
            answer = resumir_resultado(user_prompt, sql_query, df_result, client)
            st.session_state.chat_history.append({"role": "assistant", "content": answer})
        except Exception as exc:
            error_message = (
                "No se pudo obtener la respuesta automática. "
                f"Detalle técnico: {exc}"
            )
            st.session_state.chat_history.append({"role": "assistant", "content": error_message})

    st.sidebar.header("Estado del sistema")
    st.sidebar.write(f"Base de datos: {DB_PATH.name}")
    st.sidebar.write(f"Tabla: {TABLE_NAME}")
    st.sidebar.write(f"Fecha del servidor: {date.today().isoformat()}")


if __name__ == "__main__":
    main()
