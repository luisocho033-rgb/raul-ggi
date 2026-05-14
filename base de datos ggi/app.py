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
    # Normalizar nombres de columnas del Excel para buscarlas fácilmente
    df.columns = [c.lower().strip() for c in df.columns]
    
    # Identificar columnas por palabras clave (sin requerimientos estrictos)
    posibles_ids = [c for c in df.columns if any(x in c for x in ["ced", "ci", "p00", "id", "identific"])]
    posibles_nombres = [c for c in df.columns if any(x in c for x in ["nom", "trabajador", "empleado"])]
    posibles_fechas = [c for c in df.columns if any(x in c for x in ["fech", "dia", "date"])]
    posibles_deptos = [c for c in df.columns if any(x in c for x in ["dep", "area", "div"])]
    posibles_status = [c for c in df.columns if any(x in c for x in ["est", "situ", "motivo", "asist"])]

    df_final = pd.DataFrame()
    
    if posibles_fechas: 
        df_final["fecha"] = pd.to_datetime(df[posibles_fechas[0]], errors='coerce').dt.strftime('%Y-%m-%d')
    if posibles_ids: 
        # Limpiar el ID: quitar decimales si vienen del Excel (.0)
        df_final["identificador"] = df[posibles_ids[0]].astype(str).str.replace(".0", "", regex=False).str.strip()
    if posibles_nombres: 
        df_final["nombre_trabajador"] = df[posibles_nombres[0]].astype(str).str.strip()
    if posibles_deptos: 
        df_final["departamento"] = df[posibles_deptos[0]].astype(str).str.strip()
    if posibles_status: 
        df_final["estatus"] = df[posibles_status[0]].astype(str).str.strip()

    if not df_final.empty:
        # Usamos 'REPLACE' para actualizar registros si ya existen la misma fecha e ID
        df_final.to_sql(TABLE_NAME, conn, if_exists='append', index=False)
        return len(df_final)
    return 0

def main():
    st.set_page_config(page_title="Gestión ggi", layout="wide")
    st.title("Gestión de Asistencia Masiva para ggi")

    conn = init_db()
    
    # Configuración de Groq (Modelo rápido y gratuito)
    api_key = st.secrets.get("GROQ_API_KEY")
    client = groq.Client(api_key=api_key) if api_key else None

    # --- BARRA LATERAL: GESTIÓN DE ARCHIVOS ---
    with st.sidebar:
        st.header("📂 Carga de Datos")
        uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
        if uploaded_file:
            # Nombre para identificar el archivo en la lista
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
        if not archivos:
            st.info("No hay archivos en el servidor.")
        for arc in archivos:
            col_name, col_del = st.columns([4, 1])
            col_name.text(f"📄 {arc.name}")
            if col_del.button("🗑️", key=str(arc)):
                arc.unlink()
                st.rerun()

    # --- VISTA PRINCIPAL: TABLA DE NOVEDADES ---
    st.subheader("Resumen General de Novedades")
    try:
        # Buscamos ausencias típicas para mostrar en la tabla principal
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
            st.info("Sube archivos para visualizar las novedades aquí.")
    except Exception as e:
        st.error(f"Error al cargar tabla visual: {e}")

    # --- CHAT CON INTELIGENCIA ARTIFICIAL ---
    st.markdown("---")
    st.subheader("💬 Consultar con IA")
    
    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if p := st.chat_input("Ej: ¿Cuántas vacaciones tiene el P00 123456 este mes?"):
        st.session_state.messages.append({"role": "user", "content": p})
        with st.chat_message("user"):
            st.markdown(p)

        if not client:
            st.warning("Falta la clave de API de Groq en los secretos.")
        else:
            try:
                # 1. Generación de SQL
                sys_sql = (
                    "Eres un experto en SQLite. Tabla: 'asistencia_diaria'. "
                    "Columnas: fecha, identificador, nombre_trabajador, departamento, estatus. "
                    "Responde SOLO con el código SQL SELECT."
                )
                res_sql = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[{"role": "system", "content": sys_sql}, {"role": "user", "content": p}],
                    temperature=0
                )
                
                # Limpieza del SQL generado
                sql_final = res_sql.choices[0].message.content.strip()
                sql_final = sql_final.replace("```sql", "").replace("
```", "").replace(";", "")
                
                # 2. Ejecución y Resumen
                df_res = pd.read_sql(sql_final, conn)
                
                res_ia = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": "Eres un analista de RRHH para ggi. Explica los resultados de forma directa. Recuerda: 6 dígitos es P00, más es CI."},
                        {"role": "user", "content": f"Pregunta: {p} \n Datos obtenidos: {df_res.to_json(orient='records')}"}
                    ]
                )
                
                respuesta = res_ia.choices[0].message.content
                with st.chat_message("assistant"):
                    st.markdown(respuesta)
                st.session_state.messages.append({"role": "assistant", "content": respuesta})
                
            except Exception as e:
                error_msg = "No encontré datos suficientes para esa consulta o el formato es inválido."
                st.error(f"{error_msg} (Detalle: {e})")

if __name__ == "__main__":
    main()
