import json
import sqlite3
import os
from pathlib import Path
import groq
import pandas as pd
import streamlit as st

# --- CONFIGURACIÓN DE RUTAS Y TABLAS ---
DB_PATH = Path("asistencia.db")
UPLOAD_DIR = Path("archivos_excel")
UPLOAD_DIR.mkdir(exist_ok=True)
TABLE_NAME = "asistencia_diaria"

def init_db():
    """Inicializa la base de datos con la columna unificada 'identificador'."""
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
    """Lee el Excel y mapea columnas automáticamente al estándar de la DB."""
    df = pd.read_excel(file_path, engine="openpyxl")
    df.columns = [c.lower().strip() for c in df.columns]
    
    posibles_ids = [c for c in df.columns if any(x in c for x in ["ced", "ci", "p00", "id", "identific"])]
    posibles_nombres = [c for c in df.columns if any(x in c for x in ["nom", "trabajador", "empleado"])]
    posibles_fechas = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()
    
    if posibles_fechas: 
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors='coerce').dt.strftime('%Y-%m-%d')
    if posibles_ids: 
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False).str.strip()
    if posibles_nombres: 
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str).str.strip()
    if posibles_deptos: 
        df_final["departamento"] = df[posibles_deptos[0]].astype(str).str.strip()
    if posibles_status: 
        df_final["estatus"] = df[posibles_status[0]].astype(str).str.strip()

    if not df_final.empty:
        df_final.to_sql(TABLE_NAME, conn, if_exists='append', index=False)
        return len(df_final)
    return 0

def main():
    st.set_page_config(page_title="Gestión ggi", layout="wide")
    st.title("Gestión de Asistencia Masiva para ggi")

    conn = init_db()
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # --- BARRA LATERAL: GESTIÓN DE ARCHIVOS ---
    with st.sidebar:
        st.header("📂 Carga de Datos")
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
                    st.success(f"¡Éxito! {regs} registros procesados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.header("📋 Archivos Guardados")
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        for arc in archivos:
            col_name, col_del = st.columns([4, 1])
            col_name.text(f"📄 {arc.name}")
            if col_del.button("🗑️", key=str(arc)):
                arc.unlink()
                st.rerun()

    # --- VISTA PRINCIPAL ---
    st.subheader("Resumen General de Novedades")
    try:
        query_view = f"""
            SELECT identificador, 
                   CASE WHEN LENGTH(identificador) = 6 THEN 'P00' ELSE 'CI' END AS tipo,
                   nombre_trabajador, estatus, fecha 
            FROM {TABLE_NAME} 
            WHERE estatus LIKE '%Vacaciones%' 
               OR estatus LIKE '%Teletrabajo%' 
               OR estatus LIKE '%Ausente%'
            ORDER BY fecha DESC LIMIT 50
        """
        data_view = pd.read_sql(query_view, conn)
        if not data_view.empty:
            st.dataframe(data_view, use_container_width=True)
        else:
            st.info("No hay novedades detectadas.")
    except Exception as e:
        st.error(f"Error en tabla: {e}")

    # --- CHAT CON IA ---
    st.markdown("---")
    st.subheader("💬 Consultar con IA")
    
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
                sys_sql = "Responde SOLO con SQL SELECT. Tabla: asistencia_diaria. Columnas: fecha, identificador, nombre_trabajador, departamento, estatus."
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[{"role": "system", "content": sys_sql}, {"role": "user", "content": p}],
                    temperature=0
                )
                
                # LIMPIEZA DE SQL SEGURA
                raw_sql = res_sql.choices[0].message.content.strip()
                sql_final = raw_sql.replace("```sql", "").replace("
```", "")
                sql_final = sql_final.replace(";", "").strip()
                
                # 2. Consultar y responder
                df_res = pd.read_sql(sql_final, conn)
                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": "Analista de ggi. 6 dígitos = P00, +6 = CI. Sé directo."},
                        {"role": "user", "content": f"Pregunta: {p} \n Datos: {df_res.to_json(orient='records')}"}
                    ]
                )
                
                respuesta = res_ia.choices[0].message.content
                with st.chat_message("assistant"):
                    st.markdown(respuesta)
                st.session_state.messages.append({"role": "assistant", "content": respuesta})
                
            except Exception as e:
                st.error(f"Consulta fallida: {e}")

if __name__ == "__main__":
    main()    posibles_fechas = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()
    
    if posibles_fechas: 
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors='coerce').dt.strftime('%Y-%m-%d')
    if posibles_ids: 
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False).str.strip()
    if posibles_nombres: 
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str).str.strip()
    if posibles_deptos: 
        df_final["departamento"] = df[posibles_deptos[0]].astype(str).str.strip()
    if posibles_status: 
        df_final["estatus"] = df[posibles_status[0]].astype(str).str.strip()

    if not df_final.empty:
        df_final.to_sql(TABLE_NAME, conn, if_exists='append', index=False)
        return len(df_final)
    return 0

def main():
    st.set_page_config(page_title="Gestión ggi", layout="wide")
    st.title("Gestión de Asistencia Masiva para ggi")

    conn = init_db()
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # --- BARRA LATERAL: GESTIÓN DE ARCHIVOS ---
    with st.sidebar:
        st.header("📂 Carga de Datos")
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
                    st.success(f"¡Éxito! {regs} registros procesados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.header("📋 Archivos Guardados")
        archivos = list(UPLOAD_DIR.glob("*.xlsx"))
        for arc in archivos:
            col_name, col_del = st.columns([4, 1])
            col_name.text(f"📄 {arc.name}")
            if col_del.button("🗑️", key=str(arc)):
                arc.unlink()
                st.rerun()

    # --- VISTA PRINCIPAL ---
    st.subheader("Resumen General de Novedades")
    try:
        query_view = f"""
            SELECT identificador, 
                   CASE WHEN LENGTH(identificador) = 6 THEN 'P00' ELSE 'CI' END AS tipo,
                   nombre_trabajador, estatus, fecha 
            FROM {TABLE_NAME} 
            WHERE estatus LIKE '%Vacaciones%' 
               OR estatus LIKE '%Teletrabajo%' 
               OR estatus LIKE '%Ausente%'
            ORDER BY fecha DESC LIMIT 50
        """
        data_view = pd.read_sql(query_view, conn)
        if not data_view.empty:
            st.dataframe(data_view, use_container_width=True)
        else:
            st.info("No hay novedades detectadas.")
    except Exception as e:
        st.error(f"Error en tabla: {e}")

    # --- CHAT CON IA ---
    st.markdown("---")
    st.subheader("💬 Consultar con IA")
    
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
                sys_sql = "Responde SOLO con SQL SELECT. Tabla: asistencia_diaria. Columnas: fecha, identificador, nombre_trabajador, departamento, estatus."
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[{"role": "system", "content": sys_sql}, {"role": "user", "content": p}],
                    temperature=0
                )
                
                # LIMPIEZA DE SQL SEGURA
                raw_sql = res_sql.choices[0].message.content.strip()
                sql_final = raw_sql.replace("```sql", "").replace("
```", "")
                sql_final = sql_final.replace(";", "").strip()
                
                # 2. Consultar y responder
                df_res = pd.read_sql(sql_final, conn)
                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": "Analista de ggi. 6 dígitos = P00, +6 = CI. Sé directo."},
                        {"role": "user", "content": f"Pregunta: {p} \n Datos: {df_res.to_json(orient='records')}"}
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
