import json
import re
import sqlite3
from pathlib import Path
from datetime import date

import groq
import pandas as pd
import streamlit as st

DB_PATH = Path("asistencia.db")
TABLE_NAME = "asistencia_diaria"
COLUMN_ALIASES = {
    "fecha": ["fecha", "date", "dia", "día"],
    "cedula_trabajador": ["cedula_trabajador", "cedula", "id", "identificacion", "identificación", "ci", "dni"],
    "nombre_trabajador": ["nombre_trabajador", "nombre", "nombre_completo", "trabajador", "personal", "empleado"],
    "departamento": ["departamento", "area", "área", "division", "división", "sector"],
    "estatus": ["estatus", "estado", "asistencia", "situacion", "situación", "motivo", "ausencia", "observacion", "observación"],
}
REQUIRED_FIELDS = list(COLUMN_ALIASES.keys())
AUSENCIAS = ("Vacaciones", "Teletrabajo", "Ausente")

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

def guess_column_by_content(df):
    """Intenta detectar automáticamente las columnas esenciales analizando contenido y encabezados."""
    result = {}
    lowered = {col.lower().strip(): col for col in df.columns}
    # Busca primero por alias en encabezado
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lowered:
                result[field] = lowered[alias]
                break

    # Si no se encuentra por alias, intenta por contenido
    for field in REQUIRED_FIELDS:
        if field not in result:
            candidates = []
            for col in df.columns:
                s = df[col]
                # Detecta fecha
                if field == "fecha":
                    num_dates = pd.to_datetime(s, errors="coerce").notna().sum()
                    if num_dates > len(s) // 4: # Al menos 25% parecen fechas
                        candidates.append((col, num_dates))
                # Detecta cédula (5-9 dígitos)
                elif field == "cedula_trabajador":
                    mask = s.astype(str).str.match(r"^\d{5,9}$", na=False)
                    if mask.sum() > len(s) // 2:
                        candidates.append((col, mask.sum()))
                # Detecta nombre (múltiples palabras)
                elif field == "nombre_trabajador":
                    ntext = s.astype(str).str.count(r" ").sum()
                    if ntext > len(s) // 5:
                        candidates.append((col, ntext))
                # Detecta departamento (pocos únicos)
                elif field == "departamento":
                    nunique = s.nunique()
                    if nunique <= len(s) // 5:
                        candidates.append((col, nunique))
                # Detecta estatus (pocos únicos)
                elif field == "estatus":
                    nunique = s.nunique()
                    if nunique <= 12:
                        candidates.append((col, nunique))
            if candidates:
                # Elige el que más coincide
                result[field] = sorted(candidates, key=lambda x: x[1], reverse=True)[0][0]
    # Aviso si falta alguna
    missing = [f for f in REQUIRED_FIELDS if f not in result]
    if missing:
        st.warning("No se detectaron las siguientes columnas, es posible que la información esté incompleta: " + ", ".join(missing))
    return result

def procesar_excel(uploaded_file, conn):
    """Procesa el archivo Excel, detecta columnas y graba en SQLite."""
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    colmaps = guess_column_by_content(df)
    dfout = pd.DataFrame()
    for std, orig in colmaps.items():
        dfout[std] = df[orig]

    # Normalize
    dfout["fecha"] = pd.to_datetime(dfout["fecha"], errors="coerce").dt.date
    dfout["cedula_trabajador"] = dfout["cedula_trabajador"].astype(str).str.strip()
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
        (fecha, cedula_trabajador, nombre_trabajador, departamento, estatus)
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

def resumen_ausencias(conn: sqlite3.Connection):
    """Cálculo y despliegue del resumen de ausencias por persona y periodo."""
    st.markdown("## Resumen de ausencias (Vacaciones, Teletrabajo, Ausente, solo CI 6 dígitos)")
    query = f"""
    SELECT 
        cedula_trabajador, 
        nombre_trabajador, 
        estatus AS motivo, 
        fecha
    FROM {TABLE_NAME}
    WHERE 
        estatus IN {AUSENCIAS}
        AND LENGTH(cedula_trabajador) = 6
        AND cedula_trabajador GLOB '[0-9][0-9][0-9][0-9][0-9][0-9]'
    ORDER BY 
        cedula_trabajador, motivo, fecha
    """
    try:
        df = pd.read_sql_query(query, conn)
        if df.empty:
            st.info("No se encontraron ausencias para CIs de 6 dígitos.")
            return

        df['fecha'] = pd.to_datetime(df['fecha'])
        df['grupo'] = (df['fecha'].diff().dt.days.ne(1) | 
                       (df['cedula_trabajador'] != df['cedula_trabajador'].shift(1)) | 
                       (df['motivo'] != df['motivo'].shift(1))).cumsum()
        resumen = (
            df.groupby(['cedula_trabajador', 'nombre_trabajador', 'motivo', 'grupo'])
            .agg(fecha_inicio=('fecha', 'min'), fecha_fin=('fecha', 'max'))
            .reset_index()
            .drop('grupo', axis=1)
            .sort_values(['fecha_inicio'])
        )
        resumen['fecha_inicio'] = resumen['fecha_inicio'].dt.strftime('%Y-%m-%d')
        resumen['fecha_fin'] = resumen['fecha_fin'].dt.strftime('%Y-%m-%d')

        st.dataframe(resumen, use_container_width=True)
        st.markdown(
            "_Muestra sólo períodos con días consecutivos por cada motivo y persona (Verifica que tu Excel tenga registros diarios para precisión)._"
        )
    except Exception as e:
        st.error(f"Error al calcular el resumen de ausencias: {e}")

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

    resumen_ausencias(conn)

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
