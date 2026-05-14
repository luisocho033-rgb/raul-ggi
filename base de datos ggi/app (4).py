"""
Gestión de Asistencia · GGI
============================
FIXES aplicados:
  1. groq.Client → groq.Groq  (bug que hacía fallar toda la IA)
  2. Datos en st.session_state → persisten mientras el navegador esté abierto
  3. Lectura de Excel mejorada: detecta layouts horizontales y verticales
  4. Contexto completo a la IA (hasta 300 filas + resúmenes)
  5. Google Gemini como alternativa gratuita a Groq
  6. Sin dependencia de SQLite en disco (Streamlit Cloud borra archivos)
"""

import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

# ─── CLIENTES IA ─────────────────────────────────────────────────────────────
def get_groq_client(api_key: str):
    try:
        import groq  # pip install groq
        return groq.Groq(api_key=api_key)   # ← CORRECTO (antes era groq.Client → BUG)
    except Exception as e:
        st.error(f"Error cargando Groq: {e}")
        return None

def get_gemini_client(api_key: str):
    try:
        import google.generativeai as genai  # pip install google-generativeai
        genai.configure(api_key=api_key)
        return genai.GenerativeModel("gemini-1.5-flash")   # gratis y rápido
    except Exception as e:
        st.error(f"Error cargando Gemini: {e}")
        return None

# ─── LECTURA DE EXCEL ─────────────────────────────────────────────────────────
POSIBLES_FECHA   = ["fech", "dia", "date", "periodo"]
POSIBLES_ID      = ["ced", "ci", "p00", "identific", "id", "num", "ficha", "cod"]
POSIBLES_NOMBRE  = ["nom", "trabajador", "empleado", "personal", "apellido"]
POSIBLES_DEPTO   = ["dep", "area", "div", "unidad", "secc", "gerencia", "dpto"]
POSIBLES_STATUS  = ["estat", "situ", "motivo", "asist", "nov", "tipo", "condic",
                    "falt", "ausent", "vacac", "telet"]

def _buscar_col(cols, keywords):
    for c in cols:
        if any(k in c for k in keywords):
            return c
    return None

def _detectar_orientacion(df: pd.DataFrame):
    """
    Devuelve 'vertical' (filas=registros) o 'horizontal' (fechas como columnas).
    Heurística: si más del 40 % de los encabezados parecen fechas → horizontal.
    """
    date_like = sum(
        1 for c in df.columns
        if re.search(r"\d{1,2}[\/\-]\d{1,2}[\/\-]?\d{0,4}", str(c))
           or re.search(r"^\d{5}$", str(c))   # número serial de fecha Excel
    )
    return "horizontal" if date_like / max(len(df.columns), 1) > 0.4 else "vertical"

def _pivot_horizontal(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte tabla horizontal (empleado × fecha) a vertical (registro por fila).
    Detecta la columna de nombre/id y todas las columnas de fecha como valores.
    """
    cols = [str(c).lower().strip() for c in df.columns]
    df.columns = cols

    col_id     = _buscar_col(cols, POSIBLES_ID)
    col_nombre = _buscar_col(cols, POSIBLES_NOMBRE)
    col_depto  = _buscar_col(cols, POSIBLES_DEPTO)

    # Las columnas de fecha son las que no son meta-datos
    meta = [c for c in [col_id, col_nombre, col_depto] if c]
    fecha_cols = [c for c in cols if c not in meta]

    if not fecha_cols:
        return df  # no se puede transformar, devolver tal cual

    id_vars = meta
    melted = df.melt(id_vars=id_vars, value_vars=fecha_cols,
                     var_name="fecha", value_name="estatus")

    rename = {}
    if col_id:     rename[col_id]     = "identificador"
    if col_nombre: rename[col_nombre] = "nombre_trabajador"
    if col_depto:  rename[col_depto]  = "departamento"
    melted = melted.rename(columns=rename)

    # Convertir fecha serial o string
    def parse_fecha(v):
        try:
            # Puede ser número serial de Excel
            if str(v).isdigit():
                from datetime import datetime, timedelta
                return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
            return pd.to_datetime(v, dayfirst=True, errors="coerce").strftime("%Y-%m-%d")
        except Exception:
            return str(v)

    melted["fecha"] = melted["fecha"].apply(parse_fecha)
    melted = melted[melted["estatus"].notna() & (melted["estatus"].astype(str).str.strip() != "")]
    return melted.reset_index(drop=True)

def procesar_excel(file_bytes: bytes, nombre_archivo: str) -> pd.DataFrame:
    """
    Lee un Excel (cualquier orientación) y devuelve un DataFrame normalizado con columnas:
    fecha, identificador, nombre_trabajador, departamento, estatus
    """
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        raise ValueError(f"No se pudo leer el Excel: {e}")

    df_raw.columns = [str(c).lower().strip() for c in df_raw.columns]
    df_raw = df_raw.dropna(how="all")

    orientacion = _detectar_orientacion(df_raw)

    if orientacion == "horizontal":
        df = _pivot_horizontal(df_raw)
    else:
        # Layout vertical normal
        cols = list(df_raw.columns)
        col_id     = _buscar_col(cols, POSIBLES_ID)
        col_nombre = _buscar_col(cols, POSIBLES_NOMBRE)
        col_fecha  = _buscar_col(cols, POSIBLES_FECHA)
        col_depto  = _buscar_col(cols, POSIBLES_DEPTO)
        col_status = _buscar_col(cols, POSIBLES_STATUS)

        df = pd.DataFrame()
        if col_fecha:
            df["fecha"] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
        if col_id:
            df["identificador"] = df_raw[col_id].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        if col_nombre:
            df["nombre_trabajador"] = df_raw[col_nombre].astype(str).str.strip()
        if col_depto:
            df["departamento"] = df_raw[col_depto].astype(str).str.strip()
        if col_status:
            df["estatus"] = df_raw[col_status].astype(str).str.strip()

        # Fallback: sin columna de status, intentar usar todas las columnas como contexto
        if "estatus" not in df.columns:
            df["estatus"] = df_raw.apply(
                lambda r: " | ".join(str(v) for v in r if pd.notna(v) and str(v).strip()),
                axis=1
            )

    # Garantizar columnas mínimas
    for col, default in [
        ("fecha", "sin-fecha"),
        ("identificador", None),
        ("nombre_trabajador", ""),
        ("departamento", ""),
        ("estatus", ""),
    ]:
        if col not in df.columns:
            if col == "identificador":
                df["identificador"] = [f"fila_{i}" for i in range(len(df))]
            else:
                df[col] = default

    df["_archivo"] = nombre_archivo
    df = df.fillna("").astype(str)
    df = df[df["estatus"].str.strip() != ""]
    return df.reset_index(drop=True)

# ─── SESSION STATE ────────────────────────────────────────────────────────────
def init_state():
    if "archivos" not in st.session_state:
        st.session_state["archivos"] = {}   # {nombre: bytes}
    if "dataframes" not in st.session_state:
        st.session_state["dataframes"] = {}  # {nombre: DataFrame}
    if "messages" not in st.session_state:
        st.session_state["messages"] = []

def get_df_global() -> pd.DataFrame:
    dfs = list(st.session_state["dataframes"].values())
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)

# ─── CONTEXTO PARA LA IA ──────────────────────────────────────────────────────
def build_context(df: pd.DataFrame, max_filas: int = 300) -> str:
    if df.empty:
        return "No hay datos cargados en este momento."

    total = len(df)
    fechas = sorted(df["fecha"].dropna().unique().tolist())

    # Conteo por estatus
    by_status = df["estatus"].value_counts().head(20).to_dict()
    status_txt = "\n".join(f"  {k}: {v}" for k, v in by_status.items())

    # Patrones de ausencia/novedad
    pat_aus  = r"ausent|falt|no asist|inasist"
    pat_tele = r"telet|telew|remot|home"
    pat_vac  = r"vacac|vacation"
    pat_perm = r"permis|licen|reposo|médic|medic"

    def filtrar(pat):
        return df[df["estatus"].str.contains(pat, case=False, na=False)]

    aus   = filtrar(pat_aus)
    tele  = filtrar(pat_tele)
    vacas = filtrar(pat_vac)
    perms = filtrar(pat_perm)

    def tabla_mini(sub, cols=("nombre_trabajador","identificador","departamento","fecha","estatus")):
        cols_disp = [c for c in cols if c in sub.columns]
        return sub[cols_disp].head(100).to_string(index=False) if not sub.empty else "Ninguno"

    # Muestra amplia de datos (hasta max_filas)
    muestra = df.head(max_filas).to_string(index=False)

    archivos = df["_archivo"].unique().tolist() if "_archivo" in df.columns else []

    ctx = f"""
=== DATOS DE ASISTENCIA · GGI ===
Archivos cargados: {', '.join(archivos) if archivos else 'desconocido'}
Total registros  : {total}
Fechas presentes : {fechas}

── CONTEO POR ESTATUS ──
{status_txt}

── AUSENTES / FALTAS ({len(aus)}) ──
{tabla_mini(aus)}

── TELETRABAJO ({len(tele)}) ──
{tabla_mini(tele)}

── VACACIONES ({len(vacas)}) ──
{tabla_mini(vacas)}

── PERMISOS / LICENCIAS ({len(perms)}) ──
{tabla_mini(perms)}

── TODOS LOS DATOS (hasta {max_filas} filas) ──
{muestra}

NOTA IMPORTANTE:
- Identificador de 6 dígitos = ficha P00 (ej: 123456 → P00123456)
- Más de 6 dígitos = cédula (CI)
- Si el usuario pregunta con palabras aproximadas (ej: "faltaron", "no vinieron",
  "ausentes", "no asistieron") interprétalas todas como ausencias.
- Si pregunta por un nombre, busca coincidencia parcial insensible a mayúsculas.
""".strip()
    return ctx

# ─── LLAMADA A LA IA ──────────────────────────────────────────────────────────
SYSTEM_BASE = """Eres el analista de RRHH de GGI. Tienes acceso a los datos REALES de asistencia.

REGLAS:
1. SIEMPRE responde usando los datos que te doy. NUNCA digas que no tienes datos si hay registros.
2. Si preguntan por ausentes, lista nombres, ID, departamento y fecha.
3. Da números exactos basados en los datos reales.
4. Si el estatus usa sinónimos ("No asistió", "Falta", "Inasistencia") tratalos como ausencias.
5. Responde en español, claro y directo. Usa listas cuando hay varios nombres.
6. Si los datos están incompletos, dilo claramente pero responde con lo que hay.
7. Puedes hacer cálculos, comparaciones y resumenes sobre los datos.
"""

def ask_groq(client, messages_hist, system_prompt):
    history = [{"role": m["role"], "content": m["content"]} for m in messages_hist]
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "system", "content": system_prompt}, *history],
        temperature=0.1,
        max_tokens=2000,
    )
    return resp.choices[0].message.content

def ask_gemini(client, messages_hist, system_prompt):
    # Gemini usa una API diferente
    conv_text = "\n".join(
        f"{'Usuario' if m['role']=='user' else 'Asistente'}: {m['content']}"
        for m in messages_hist
    )
    prompt = f"{system_prompt}\n\nConversación:\n{conv_text}"
    resp = client.generate_content(prompt)
    return resp.text

# ─── UI PRINCIPAL ─────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Gestión GGI", layout="wide", page_icon="📋")
    st.title("📋 Gestión de Asistencia · GGI")

    init_state()

    # ── Detectar qué IA usar ─────────────────────────────────────────────────
    groq_key   = st.secrets.get("GROQ_API_KEY", "")
    gemini_key = st.secrets.get("GEMINI_API_KEY", "")

    ai_client  = None
    ai_tipo    = None

    if groq_key:
        ai_client = get_groq_client(groq_key)
        if ai_client:
            ai_tipo = "groq"
    if ai_client is None and gemini_key:
        ai_client = get_gemini_client(gemini_key)
        if ai_client:
            ai_tipo = "gemini"

    # ── SIDEBAR ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("📂 Carga de Excel")

        uploaded = st.file_uploader("Subir archivo Excel", type=["xlsx", "xls"])
        if uploaded:
            nombre_id = st.text_input(
                "Nombre para identificar este archivo",
                value=re.sub(r"\.xlsx?$", "", uploaded.name, flags=re.I)
            )
            if st.button("✅ Procesar y Cargar"):
                try:
                    file_bytes = uploaded.read()
                    df_new = procesar_excel(file_bytes, nombre_id)
                    st.session_state["archivos"][nombre_id]   = file_bytes
                    st.session_state["dataframes"][nombre_id] = df_new
                    st.success(f"¡Listo! {len(df_new)} registros cargados de **{nombre_id}**")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

        st.markdown("---")
        st.subheader("📁 Archivos en Sesión")

        if not st.session_state["dataframes"]:
            st.caption("Ningún archivo cargado aún.")
        else:
            for nombre, df_arc in list(st.session_state["dataframes"].items()):
                c1, c2 = st.columns([5, 1])
                c1.markdown(f"📄 **{nombre}** _{len(df_arc)} reg._")
                if c2.button("✕", key=f"del_{nombre}"):
                    del st.session_state["dataframes"][nombre]
                    st.session_state["archivos"].pop(nombre, None)
                    st.rerun()

        if st.session_state["dataframes"]:
            st.markdown("---")
            if st.button("🗑️ Borrar todos los datos", type="secondary"):
                st.session_state["dataframes"] = {}
                st.session_state["archivos"]   = {}
                st.rerun()

        st.markdown("---")
        # Indicador de IA activa
        if ai_tipo == "groq":
            st.success("🤖 IA: Groq (Llama 3.3 70B)")
        elif ai_tipo == "gemini":
            st.success("🤖 IA: Google Gemini Flash")
        else:
            st.warning("⚠️ Sin IA configurada")
            with st.expander("¿Cómo agregar IA?"):
                st.markdown("""
**Opción A – Groq (gratis)**
1. Ve a https://console.groq.com → crea cuenta gratis
2. Genera una API Key
3. En Streamlit Cloud → Settings → Secrets agrega:
```
GROQ_API_KEY = "gsk_xxxxx"
```

**Opción B – Google Gemini (gratis)**
1. Ve a https://aistudio.google.com
2. Genera API Key gratis
3. En Secrets agrega:
```
GEMINI_API_KEY = "AIzaSy_xxxxx"
```
""")

    # ── DATOS GLOBALES ────────────────────────────────────────────────────────
    df_full = get_df_global()

    # ── MÉTRICAS ─────────────────────────────────────────────────────────────
    if not df_full.empty:
        n_aus  = len(df_full[df_full["estatus"].str.contains(r"ausent|falt|no asist|inasist", case=False, na=False)])
        n_tele = len(df_full[df_full["estatus"].str.contains(r"telet|telew|remot", case=False, na=False)])
        n_vac  = len(df_full[df_full["estatus"].str.contains(r"vacac", case=False, na=False)])
        n_perm = len(df_full[df_full["estatus"].str.contains(r"permis|licen|reposo", case=False, na=False)])

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total registros", len(df_full))
        c2.metric("🔴 Ausentes",    n_aus)
        c3.metric("💻 Teletrabajo", n_tele)
        c4.metric("🌴 Vacaciones",  n_vac)
        c5.metric("📝 Permisos",    n_perm)
        st.markdown("---")

    # ── TABLA DE NOVEDADES ────────────────────────────────────────────────────
    st.subheader("📊 Resumen de Novedades")

    if df_full.empty:
        st.info("No hay datos. Carga un archivo Excel en el panel izquierdo.")
    else:
        novedades = df_full[
            df_full["estatus"].str.contains(
                r"vacac|telet|telew|ausent|falt|permis|licen|reposo|no asist|inasist|remot",
                case=False, na=False
            )
        ].copy()

        if "_archivo" in novedades.columns:
            novedades = novedades.rename(columns={"_archivo": "archivo"})

        cols_mostrar = [c for c in
            ["fecha", "identificador", "nombre_trabajador", "departamento", "estatus", "archivo"]
            if c in novedades.columns]

        if novedades.empty:
            st.info("Sin novedades detectadas. Puede que el estatus use términos no reconocidos; "
                    "revisa la tabla completa abajo.")
        else:
            st.dataframe(novedades[cols_mostrar], use_container_width=True, hide_index=True)

        with st.expander("📋 Ver TODOS los datos cargados"):
            st.dataframe(df_full, use_container_width=True, hide_index=True)

        with st.expander("🔍 Filtrar por fecha o departamento"):
            fechas_disp = sorted(df_full["fecha"].unique().tolist())
            fecha_sel   = st.selectbox("Fecha", ["Todas"] + fechas_disp)
            deptos_disp = sorted(df_full["departamento"].unique().tolist())
            depto_sel   = st.selectbox("Departamento", ["Todos"] + deptos_disp)

            df_filtrado = df_full.copy()
            if fecha_sel != "Todas":
                df_filtrado = df_filtrado[df_filtrado["fecha"] == fecha_sel]
            if depto_sel != "Todos":
                df_filtrado = df_filtrado[df_filtrado["departamento"] == depto_sel]

            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
            st.caption(f"{len(df_filtrado)} registros filtrados.")

    # ── CHAT CON IA ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🤖 Preguntar a la IA sobre los datos")

    if ai_client is None:
        st.warning(
            "Configura **GROQ_API_KEY** o **GEMINI_API_KEY** en los secrets de Streamlit "
            "para activar el chat con IA. Ver instrucciones en el panel izquierdo."
        )
    else:
        # Mostrar historial
        for m in st.session_state["messages"]:
            with st.chat_message(m["role"]):
                st.markdown(m["content"])

        if prompt := st.chat_input("Pregunta lo que quieras sobre la asistencia..."):
            st.session_state["messages"].append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            # Construir contexto fresco
            ctx = build_context(get_df_global(), max_filas=300)
            system_prompt = SYSTEM_BASE + f"\n\nDATOS ACTUALES:\n{ctx}"

            with st.chat_message("assistant"):
                with st.spinner("Analizando datos..."):
                    try:
                        if ai_tipo == "groq":
                            respuesta = ask_groq(ai_client, st.session_state["messages"], system_prompt)
                        else:
                            respuesta = ask_gemini(ai_client, st.session_state["messages"], system_prompt)

                        st.markdown(respuesta)
                        st.session_state["messages"].append(
                            {"role": "assistant", "content": respuesta}
                        )
                    except Exception as e:
                        msg = str(e)
                        st.error(f"Error al consultar la IA: {msg}")

                        # Sugerencias según el error
                        if "api_key" in msg.lower() or "auth" in msg.lower() or "401" in msg:
                            st.info("💡 Verifica que tu API Key esté bien configurada en Secrets.")
                        elif "rate" in msg.lower() or "429" in msg:
                            st.info("💡 Límite de peticiones alcanzado. Espera un momento e intenta de nuevo.")
                        elif "model" in msg.lower():
                            st.info("💡 El modelo especificado no está disponible. Revisa el nombre del modelo.")

        if st.session_state["messages"]:
            if st.button("🗑️ Limpiar chat"):
                st.session_state["messages"] = []
                st.rerun()

    # ── FOOTER ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.caption(
        "💾 Los datos se guardan en la sesión actual del navegador. "
        "Si recargas la página debes subir los archivos nuevamente. "
        "Para persistencia permanente, consulta con tu administrador sobre conectar una DB externa."
    )

if __name__ == "__main__":
    main()
