import json
import sqlite3
import os
from pathlib import Path
import groq
import pandas as pd
import streamlit as st

# Configuración
DB_PATH = Path("asistencia.db")
UPLOAD_DIR = Path("archivos_excel")
UPLOAD_DIR.mkdir(exist_ok=True)
TABLE_NAME = "asistencia_diaria"

def init_db():
    """Crea la tabla con la columna 'identificador' unificada."""
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
    return conn

def procesar_flexible(file_path, conn):
    df = pd.read_excel(file_path, engine="openpyxl")
    # Limpieza de nombres de columnas del Excel
    df.columns = [c.lower().strip() for c in df.columns]
    
    # Buscamos la columna de ID (P00 o CI) con varios nombres posibles
    posibles_ids = [c for c in df.columns if any(x in c for x in ["ced", "ci", "p00", "id", "identific"])]
    posibles_nombres = [c for c in df.columns if any(x in c for x in ["nom", "trabajador", "empleado"])]
    posibles_fechas = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()
    
    # Asignación segura: si existe la columna en el Excel, la mapeamos a nuestro estándar
    if posibles_fechas: 
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors='coerce').dt.strftime('%Y-%m-%d')
    if posibles_ids: 
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False)
    if posibles_nombres: 
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str)
    if posibles_deptos: 
        df_final["departamento"] = df[posibles_deptos[0]].astype(str)
    if posibles_status: 
        df_final["estatus"] = df[posibles_status[0]].astype(str)

    # Insertar en la DB (si hay datos)
    if not df_final.empty:
        df_final.to_sql(TABLE_NAME, conn, if_exists='append', index=False)
        return len(df_final)
    return 0

def main():
    st.set_page_config(page_title="Gestión ggi", layout="wide")
    st.title("Gestión de Asistencia Masiva para ggi")

    conn = init_db()
    
    # API Groq - Modelo Llama 3.1 8B (Gratis y rápido)
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # PANEL LATERAL: GESTIÓN DE ARCHIVOS
    with st.sidebar:
        st.header("Carga de Datos")
        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
        if uploaded_file:
            nombre_personalizado = st.text_input("Nombre para este archivo", value=uploaded_file.name)
            if st.button("Guardar y Procesar"):
                path = UPLOAD_DIR / f"{nombre_personalizado}.xlsx"
                with open(path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                try:
                    regs = procesar_flexible(path, conn)
                    st.success(f"Se cargaron {regs} registros correctamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar: {e}")

        st.markdown("---")
        st.header("Archivos guardados")
        for arc in UPLOAD_DIR.glob("*.xlsx"):
            c1, c2 = st.columns([3, 1])
            c1.text(f"📄 {arc.name}")
            if c2.button("🗑️", key=str(arc)):
                arc.unlink()
                st.rerun()

    # TABLA VISUAL DE AUSENCIAS (La que daba el error)
    st.subheader("Resumen de Novedades (P00 y CI)")
    try:
        # Consulta corregida con los nombres exactos de la tabla
        query = f"""
        SELECT 
            identificador, 
            CASE WHEN LENGTH(identificador) = 6 THEN 'P00' ELSE 'CI' END AS tipo,
            nombre_trabajador, 
            estatus, 
            fecha 
        FROM {TABLE_NAME} 
        WHERE estatus LIKE '%Vacaciones%' OR estatus LIKE '%Teletrabajo%' OR estatus LIKE '%Ausente%'
        ORDER BY fecha DESC
        """
        data_view = pd.read_sql(query, conn)
        if not data_view.empty:
            st.dataframe(data_view, use_container_width=True)
        else:
            st.info("No hay registros de ausencias detectados aún.")
    except Exception as e:
        st.error(f"Error al leer la tabla: {e}")

    # CHAT IA
    st.markdown("---")
    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.write(m["content"])

    if p := st.chat_input("Ej: ¿Cuántas veces faltó el P00 123456?"):
        st.session_state.messages.append({"role": "user", "content": p})
        with st.chat_message("user"): st.write(p)

        if client:
            try:
                # 1. Generar SQL usando el esquema real
                sys_sql = "Responde SOLO con el código SQL SELECT para SQLite. Tabla: asistencia_diaria. Columnas: fecha, identificador, nombre_trabajador, departamento, estatus."
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[{"role": "system", "content": sys_sql}, {"role": "user", "content": p}],
                    temperature=0
                )
                sql_limpio = res_sql.choices[0].message.content.strip().replace("```sql", "").replace("
```", "")
                
                # 2. Consultar y Resumir
                res_df = pd.read_sql(sql_limpio, conn)
                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": "Eres un analista de RRHH para ggi. Explica los datos brevemente. Identificadores de 6 números son P00, otros son CI."},
                        {"role": "user", "content": f"Datos: {res_df.to_json()} \nPregunta: {p}"}
                    ]
                )
                msg_ia = res_ia.choices[0].message.content
                with st.chat_message("assistant"): st.write(msg_ia)
                st.session_state.messages.append({"role": "assistant", "content": msg_ia})
            except:
                st.error("No pude procesar esa consulta. Verifica que los datos estén cargados.")

if __name__ == "__main__":
    main()
