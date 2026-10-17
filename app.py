"""
App de actualización de la base de datos del Despacho SEIN (Supabase).

Pestañas:
  1. Despacho    -> descarga IEOD del COES y guarda en "Despacho"
  2. CMG         -> descarga CMg del COES y guarda SANTA ROSA 220 en "CMG"
  3. Demanda     -> subes DemandaCOES.xlsx y guarda en "Demanda_E"
  4. Temperatura -> subes table.csv y guarda el promedio mensual en "Temperatura"
  5. Recursos    -> tabla editable, guardas en "Recursos" (energía mensual GWh)
  6. Potencia    -> tabla editable, guardas en "Potencia_Instalada" (MW mensual)
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import bd
import fuentes

st.set_page_config(page_title="Actualización de Datos SEIN", layout="wide")

HOY = datetime.now(ZoneInfo("America/Lima")).date()

# ---------------------------------------------------------------------------
# Estilo (mismo diseño que la app "Actualización de Datos COES")
# ---------------------------------------------------------------------------
st.markdown("""
<style>
/* Fondo general — azul oscuro Osinergmin */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(160deg, #001a5e, #0039AA, #001233);
    min-height: 100vh;
}
[data-testid="stHeader"] { background: transparent; }
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none; }

/* Esferas animadas */
[data-testid="stAppViewContainer"]::before,
[data-testid="stAppViewContainer"]::after {
    content: ""; position: fixed; border-radius: 50%;
    filter: blur(65px); opacity: 0.45; pointer-events: none; z-index: 0;
}
[data-testid="stAppViewContainer"]::before {
    width: 650px; height: 650px;
    background: radial-gradient(circle, #FCE122, #E08A00);
    top: -200px; left: -150px;
    animation: floatUp 9s ease-in-out infinite alternate;
}
[data-testid="stAppViewContainer"]::after {
    width: 550px; height: 550px;
    background: radial-gradient(circle, #0039AA, #12996B);
    bottom: -150px; right: -120px;
    animation: floatDown 11s ease-in-out infinite alternate;
}
@keyframes floatUp   { 0% {transform: translateY(0) scale(1);} 50% {transform: translateY(80px) scale(1.06);} 100% {transform: translateY(160px) scale(.96);} }
@keyframes floatDown { 0% {transform: translateY(0) scale(1);} 50% {transform: translateY(-80px) scale(1.05);} 100% {transform: translateY(-160px) scale(.97);} }

/* Puntitos decorativos + ancho del contenido */
[data-testid="stMain"] {
    background-image: radial-gradient(rgba(255,255,255,0.05) 1px, transparent 1px);
    background-size: 28px 28px; position: relative; z-index: 1;
}
.block-container { max-width: 1100px; padding-top: 2.5rem; }

/* Títulos */
.titulo-app { text-align: center; font-size: 2.4rem; font-weight: 800; color: #fff;
              margin-bottom: .2rem; letter-spacing: -.5px; }
.subtitulo-app { text-align: center; color: #a0a8c0; font-size: .95rem; margin-bottom: 1.6rem; }
.seccion-titulo { color: #c9d1f5; font-size: 1.05rem; font-weight: 700; margin: 1.2rem 0 .8rem 0;
                  text-transform: uppercase; letter-spacing: 1px; }
h3, h4 { color: #e0e6ff !important; }
p, li, [data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"] { color: #c9d1f5; }

/* Tarjetas de estado */
.grid-estado { display: grid; grid-template-columns: repeat(auto-fit, minmax(330px, 1fr)); gap: .6rem; }
.card-estado {
    background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.10);
    border-radius: 14px; padding: 1rem 1.3rem; display: flex;
    justify-content: space-between; align-items: center; backdrop-filter: blur(8px);
}
.card-tabla { font-weight: 600; color: #e0e6ff; font-size: .95rem; }
.card-fecha { color: #8892b0; font-size: .82rem; white-space: nowrap; }
.badge-ok, .badge-warn, .badge-err { border-radius: 20px; padding: 3px 14px; font-size: .8rem;
                                     font-weight: 700; white-space: nowrap; }
.badge-ok   { background: #0a3322; color: #12996B; border: 1px solid #12996B44; }
.badge-warn { background: #3a2800; color: #FCE122; border: 1px solid #FCE12244; }
.badge-err  { background: #3a0a0a; color: #D14343; border: 1px solid #D1434344; }

/* Etiquetas de los campos */
label, [data-testid="stWidgetLabel"] p { color: #a0a8c0 !important; font-size: .85rem !important; }

/* Pestañas tipo "píldora" */
[data-testid="stTabs"] [role="tablist"] { gap: .4rem; background: rgba(255,255,255,0.06); padding: .35rem;
        border-radius: 14px; border: 1px solid rgba(255,255,255,0.10); }
[data-testid="stTab"] { border-radius: 10px; padding: .5rem 1.1rem; color: #c9d1f5;
        background: transparent; border: none; }
[data-testid="stTab"][aria-selected="true"] { background: rgba(255,255,255,0.16); color: #fff;
        box-shadow: inset 0 -2px 0 #EF5350; }
[data-testid="stTab"]:hover { background: rgba(255,255,255,0.10); }
[data-testid="stTab"] p { font-weight: 600; font-size: .92rem; }
[data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display: none !important; }

/* Contenido de cada pestaña como tarjeta */
[data-testid="stTabs"] [role="tabpanel"] { background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08); border-radius: 16px;
        padding: 1.2rem 1.4rem 1.4rem; margin-top: .6rem; }

/* Botones */
.stButton button, .stDownloadButton button, [data-testid="stFormSubmitButton"] button {
        border-radius: 10px; font-weight: 600; }
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-primaryFormSubmit"] { width: 100%; }

/* Tablas, uploader, expander */
[data-testid="stDataFrame"], [data-testid="stFileUploaderDropzone"], [data-testid="stExpander"] {
    border-radius: 12px; }

/* Tarjeta de acceso */
.login-box { max-width: 420px; margin: 0 auto; }
</style>
""", unsafe_allow_html=True)


def encabezado(subtitulo: str):
    st.markdown('<div class="titulo-app">Actualización BD Despacho SEIN</div>'
                f'<div class="subtitulo-app">{subtitulo}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Acceso con contraseña (la app es pública en Streamlit Cloud)
# ---------------------------------------------------------------------------
def acceso_ok() -> bool:
    clave = st.secrets.get("APP_PASSWORD")
    if not clave or st.session_state.get("auth"):
        return True
    encabezado("Ingresa la contraseña para acceder.")
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        with st.form("login"):
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar", type="primary", width="stretch"):
                if pwd == clave:
                    st.session_state["auth"] = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
    return False


if not acceso_ok():
    st.stop()


# ---------------------------------------------------------------------------
# Conexión (se reutiliza; se reconecta si Supabase la cerró)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _con_cache():
    return bd.conectar(st.secrets["SUPABASE_DB_URL"])


def con():
    c = _con_cache()
    try:
        with c.cursor() as cur:
            cur.execute("select 1")
        c.rollback()
    except Exception:
        _con_cache.clear()
        c = _con_cache()
    return c


def a_excel(df: pd.DataFrame, hoja: str) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=hoja[:31])
    return buf.getvalue()


def descargas(df: pd.DataFrame, nombre: str):
    c1, c2, _ = st.columns([1, 1, 4])
    c1.download_button("⬇️ CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       f"{nombre}.csv", "text/csv", key=f"csv_{nombre}")
    c2.download_button("⬇️ Excel", a_excel(df, nombre), f"{nombre}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"xlsx_{nombre}")


# ---------------------------------------------------------------------------
# Encabezado + estado de la base (tarjetas)
# ---------------------------------------------------------------------------
def _meses_entre(a: date, b: date) -> int:
    return (b.year - a.year) * 12 + (b.month - a.month)


def tarjeta(nombre: str, detalle: str, pendiente: int, unidad: str) -> str:
    if pendiente is None:
        badge = '<span class="badge-err">Sin datos</span>'
    elif pendiente <= 0:
        badge = '<span class="badge-ok">Al día</span>'
    else:
        badge = f'<span class="badge-warn">Pendiente: {pendiente} {unidad}</span>'
    return (f'<div class="card-estado"><div><div class="card-tabla">{nombre}</div>'
            f'<div class="card-fecha">{detalle}</div></div>{badge}</div>')


encabezado("Consulta el estado de la base en Supabase y actualiza la información pendiente.")

try:
    est = bd.estado_tablas(con()).set_index("Tabla")
except Exception as e:
    st.error(f"No se pudo conectar a Supabase: {e}")
    st.stop()

ayer = HOY - timedelta(days=1)
mes_ant = (pd.Timestamp(HOY.replace(day=1)) - pd.offsets.MonthBegin(1)).date()
tarjetas = []
for tabla in ("Despacho", "CMG", "Demanda_E"):
    ult = est.loc[tabla, "Último dato"]
    if ult is None or pd.isna(ult):
        tarjetas.append(tarjeta(tabla.upper(), "Sin registros", None, ""))
        continue
    ult = pd.Timestamp(ult)
    dia_completo = (ult - pd.Timedelta(minutes=30)).date()  # 00:00 cierra el día anterior
    tarjetas.append(tarjeta(tabla.upper(), f"Último registro: {ult:%d/%m/%Y %H:%M}",
                            (ayer - dia_completo).days, "día(s)"))
ult_t = est.loc["Temperatura", "Último dato"]
tarjetas.append(tarjeta("TEMPERATURA", f"Último mes: {ult_t:%Y-%m}" if ult_t else "Sin registros",
                        _meses_entre(ult_t, HOY) if ult_t else None, "mes(es)"))
for tabla, nombre in (("Recursos", "RECURSOS"), ("Potencia_Instalada", "POTENCIA INSTALADA")):
    ult_m = est.loc[tabla, "Último dato"]
    tarjetas.append(tarjeta(nombre, f"Último mes: {ult_m:%Y-%m}" if ult_m else "Sin registros",
                            _meses_entre(ult_m, mes_ant) if ult_m else None, "mes(es)"))

c_tit, c_btn = st.columns([5, 1])
c_tit.markdown('<div class="seccion-titulo">Estado de los datos</div>', unsafe_allow_html=True)
with c_btn:
    st.write("")
    if st.button("🔄 Refrescar", width="stretch"):
        st.rerun()
st.markdown('<div class="grid-estado">' + "".join(tarjetas) + "</div>", unsafe_allow_html=True)

st.markdown('<div class="seccion-titulo">Actualizar</div>', unsafe_allow_html=True)
tab_desp, tab_cmg, tab_dem, tab_temp, tab_rec, tab_pot = st.tabs(
    ["⚡ Despacho Ejecutado", "💲 CMG", "📈 Demanda", "🌡️ Temperatura", "🏭 Recursos", "🔌 Potencia"])


def rango_fechas(clave: str, ultimo) -> tuple[date, date]:
    sugerido = (ultimo.date() if hasattr(ultimo, "date") else HOY - timedelta(days=1))
    c1, c2 = st.columns(2)
    ini = c1.date_input("Desde", value=min(sugerido, HOY), format="DD/MM/YYYY", key=f"ini_{clave}")
    fin = c2.date_input("Hasta", value=HOY - timedelta(days=1), format="DD/MM/YYYY", key=f"fin_{clave}")
    return ini, fin


def ultimo_de(tabla: str):
    u = est.loc[tabla, "Último dato"]
    return None if u is None or pd.isna(u) else pd.Timestamp(u)


# ===========================================================================
# 1) DESPACHO
# ===========================================================================
with tab_desp:
    st.markdown("#### " + "Despacho Ejecutado (IEOD - COES)")
    ult = ultimo_de("Despacho")
    st.caption(f"Último dato en la base: **{ult:%d/%m/%Y %H:%M}**" if ult is not None else "Base vacía")
    ini, fin = rango_fechas("desp", ult)

    if st.button("1️⃣ Descargar del COES", type="primary", key="btn_desp", width="stretch"):
        dias, frames = list(fuentes._fechas(ini, fin)), []
        barra = st.progress(0.0, text="Descargando…")
        sin_datos = []
        for i, dia in enumerate(dias, 1):
            barra.progress(i / len(dias), text=f"Descargando {dia:%d/%m/%Y} ({i}/{len(dias)})")
            df_dia = fuentes.leer_ieod_dia(dia)
            if df_dia is None or df_dia.empty:
                sin_datos.append(dia)
            else:
                frames.append(df_dia)
        barra.empty()
        if sin_datos:
            st.warning("Sin datos en el COES para: " + ", ".join(f"{d:%d/%m}" for d in sin_datos))
        if frames:
            try:
                crudo = pd.concat(frames, ignore_index=True)
                df, nuevas = fuentes.despacho_a_bd(crudo, bd.mapeo_centrales(con()))
                st.session_state["desp_crudo"] = crudo
                st.session_state["desp_df"], st.session_state["desp_nuevas"] = df, nuevas
            except Exception as e:
                st.error(f"Error procesando el IEOD: {e}")

    if "desp_df" in st.session_state:
        df = st.session_state["desp_df"]
        nuevas = st.session_state["desp_nuevas"]
        st.success(f"{len(df)} registros de {df['FECHA'].min():%d/%m/%Y %H:%M} a {df['FECHA'].max():%d/%m/%Y %H:%M}")
        if nuevas:
            st.warning("⚠️ **Centrales nuevas en el COES** (no están en tu base):\n\n" + "\n".join(f"- `{n}`" for n in nuevas))
            agregar = st.checkbox("Agregarlas como columnas nuevas en la base", value=False)
        else:
            agregar = False
        st.dataframe(df, width="stretch", height=300)
        descargas(df, f"Despacho_{ini}_{fin}")

        if st.button("2️⃣ Guardar en Supabase", type="primary", key="save_desp", width="stretch"):
            with st.spinner("Guardando…"):
                if agregar and nuevas:
                    bd.agregar_centrales(con(), nuevas)
                    # re-procesar con el mapeo actualizado para incluir las nuevas columnas
                    df, _ = fuentes.despacho_a_bd(st.session_state["desp_crudo"], bd.mapeo_centrales(con()))
                    st.info("Centrales nuevas agregadas: " + ", ".join(nuevas))
                n = bd.upsert(con(), "Despacho", df, ["FECHA"])
            st.success(f"✅ {n} registros guardados en Despacho")
            for k in ("desp_df", "desp_crudo", "desp_nuevas"):
                st.session_state.pop(k, None)


# ===========================================================================
# 2) CMG
# ===========================================================================
with tab_cmg:
    st.markdown("#### " + "Costo Marginal - barra SANTA ROSA 220 (COES)")
    ult = ultimo_de("CMG")
    st.caption(f"Último dato en la base: **{ult:%d/%m/%Y %H:%M}**" if ult is not None else "Base vacía")
    ini, fin = rango_fechas("cmg", ult)

    if st.button("1️⃣ Descargar del COES", type="primary", key="btn_cmg", width="stretch"):
        dias, frames, sin_datos = list(fuentes._fechas(ini, fin)), [], []
        barra = st.progress(0.0, text="Descargando…")
        for i, dia in enumerate(dias, 1):
            barra.progress(i / len(dias), text=f"Descargando {dia:%d/%m/%Y} ({i}/{len(dias)})")
            df_dia = fuentes.leer_cmg_dia(dia)
            if df_dia is None or df_dia.empty:
                sin_datos.append(dia)
                continue
            try:
                frames.append(fuentes.cmg_a_bd(df_dia))
            except Exception as e:
                st.error(f"{dia:%d/%m/%Y}: {e}")
        barra.empty()
        if sin_datos:
            st.warning("Sin datos en el COES para: " + ", ".join(f"{d:%d/%m}" for d in sin_datos))
        if frames:
            st.session_state["cmg_df"] = (pd.concat(frames).drop_duplicates("FECHA", keep="last")
                                          .sort_values("FECHA").reset_index(drop=True))

    if "cmg_df" in st.session_state:
        df = st.session_state["cmg_df"]
        st.success(f"{len(df)} registros de {df['FECHA'].min():%d/%m/%Y %H:%M} a {df['FECHA'].max():%d/%m/%Y %H:%M}")
        st.dataframe(df, width="stretch", height=300)
        descargas(df, f"CMG_{ini}_{fin}")
        if st.button("2️⃣ Guardar en Supabase", type="primary", key="save_cmg", width="stretch"):
            n = bd.upsert(con(), "CMG", df, ["FECHA"])
            st.success(f"✅ {n} registros guardados en CMG")
            del st.session_state["cmg_df"]


# ===========================================================================
# 3) DEMANDA
# ===========================================================================
with tab_dem:
    st.markdown("#### " + "Demanda ejecutada (DemandaCOES.xlsx)")
    arch = st.file_uploader("Sube DemandaCOES.xlsx", type=["xlsx", "xls"], key="up_dem")
    if arch is not None:
        try:
            df = fuentes.demanda_desde_excel(arch)
            st.success(f"{len(df)} registros de {df['FECHA'].min():%d/%m/%Y %H:%M} a {df['FECHA'].max():%d/%m/%Y %H:%M}")
            st.dataframe(df, width="stretch", height=300)
            if st.button("Guardar en Supabase", type="primary", key="save_dem", width="stretch"):
                n = bd.upsert(con(), "Demanda_E", df, ["FECHA"])
                st.success(f"✅ {n} registros guardados en Demanda_E")
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")


# ===========================================================================
# 4) TEMPERATURA
# ===========================================================================
with tab_temp:
    st.markdown("#### " + "Temperatura media mensual (table.csv - SENAMHI)")
    st.caption("Se calcula el promedio de la columna TEMPERATURA por mes. Si el mes ya existe, se reemplaza.")
    arch = st.file_uploader("Sube table.csv", type=["csv"], key="up_temp")
    if arch is not None:
        try:
            res = fuentes.temperatura_desde_csv(arch)
            st.dataframe(res, hide_index=True, width="stretch")
            if st.button("Guardar en Supabase", type="primary", key="save_temp", width="stretch"):
                n = bd.upsert(con(), "Temperatura", res[["AÑO", "MES", "MES_NOMBRE", "TEMP_MEDIA"]], ["AÑO", "MES"])
                st.success(f"✅ {n} mes(es) guardados en Temperatura")
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
    st.markdown("**En la base:**")
    st.dataframe(bd.leer(con(), 'select * from public."Temperatura" order by "AÑO" desc, "MES" desc limit 12'),
                 hide_index=True, width="stretch")


# ===========================================================================
# 5) RECURSOS
# ===========================================================================
def editor_mensual(tabla: str, clave: str, copiar_mes_anterior: bool, etiquetas: dict | None = None):
    """Tabla mensual editable (una fila por mes, clave "Fecha"): agregar, eliminar y editar meses.

    copiar_mes_anterior: al agregar un mes nuevo, lo llena con los valores del último mes
    (útil para la potencia instalada, que cambia poco); si es False lo llena con ceros.
    """
    etiquetas = etiquetas or {}
    nom_mes = lambda d: f"{fuentes.MESES[d.month - 1]} {d.year}"

    df = bd.leer(con(), f'select * from public.{bd.q(tabla)} order by "Fecha" desc')
    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.date
    num = [c for c in df.columns if c != "Fecha"]

    with st.expander("➕ Agregar un mes nuevo", expanded=False):
        ultimo_mes = df["Fecha"].max() if not df.empty else date(HOY.year, HOY.month, 1)
        sig = (pd.Timestamp(ultimo_mes) + pd.offsets.MonthBegin(1)).date()
        c1, c2 = st.columns(2)
        anio = c1.number_input("Año", 2010, 2100, sig.year, key=f"anio_{clave}")
        mes = c2.selectbox("Mes", range(1, 13), index=sig.month - 1,
                           format_func=lambda m: fuentes.MESES[m - 1], key=f"mes_{clave}")
        if copiar_mes_anterior and not df.empty:
            st.caption("Se copiarán los valores del mes anterior; luego cambia solo lo que varió.")
        if st.button("Agregar fila", key=f"add_{clave}"):
            fecha = date(int(anio), int(mes), 1)
            if fecha in set(df["Fecha"]):
                st.warning(f"{nom_mes(fecha)} ya existe; edítalo directamente en la tabla.")
            else:
                if copiar_mes_anterior and not df.empty:
                    previos = df[df["Fecha"] < fecha]  # df está ordenado de más reciente a más antiguo
                    base = (previos.iloc[0] if not previos.empty else df.iloc[-1])[num].to_dict()
                else:
                    base = {c: 0.0 for c in num}
                bd.upsert(con(), tabla, pd.DataFrame([{"Fecha": fecha, **base}]), ["Fecha"])
                st.success(f"Mes {nom_mes(fecha)} agregado. Complétalo en la tabla.")
                st.rerun()

    with st.expander("🗑️ Eliminar un mes", expanded=False):
        if df.empty:
            st.info("No hay meses para eliminar.")
        else:
            mes_borrar = st.selectbox("Mes a eliminar", df["Fecha"].tolist(), format_func=nom_mes,
                                      key=f"del_sel_{clave}")
            confirmo = st.checkbox(f"Sí, quiero eliminar {nom_mes(mes_borrar)}", key=f"del_ok_{clave}")
            if st.button("Eliminar mes", disabled=not confirmo, key=f"del_{clave}"):
                bd.eliminar_mes(con(), tabla, mes_borrar)
                st.success(f"Mes {nom_mes(mes_borrar)} eliminado.")
                st.rerun()

    editado = st.data_editor(
        df, hide_index=True, width="stretch", height=450, num_rows="fixed",
        column_config={"Fecha": st.column_config.DateColumn("Fecha", format="MMM-YYYY", disabled=True)}
        | {c: st.column_config.NumberColumn(etiquetas.get(c, c), format="%.3f") for c in num},
        key=f"ed_{clave}",
    )
    if st.button("💾 Guardar cambios", type="primary", key=f"save_{clave}", width="stretch"):
        cambios = editado.merge(df, on="Fecha", how="left", suffixes=("", "_orig"), indicator=False)
        mod = pd.Series(False, index=cambios.index)
        for c in num:
            mod |= ~((cambios[c] == cambios[f"{c}_orig"]) | (cambios[c].isna() & cambios[f"{c}_orig"].isna()))
        df_mod = editado.loc[mod.values].copy()
        df_mod[num] = df_mod[num].fillna(0)
        if df_mod.empty:
            st.info("No hay cambios que guardar.")
        else:
            n = bd.upsert(con(), tabla, df_mod, ["Fecha"])
            st.success(f"✅ {n} mes(es) actualizados: " + ", ".join(f"{d:%b-%Y}" for d in df_mod["Fecha"]))


with tab_rec:
    st.markdown("#### " + "Recursos - energía mensual (GWh)")
    st.caption("Edita las celdas o agrega un mes nuevo y luego presiona **Guardar cambios**.")
    editor_mensual("Recursos", "rec", copiar_mes_anterior=False)


# ===========================================================================
# 6) POTENCIA INSTALADA
# ===========================================================================
with tab_pot:
    st.markdown("#### " + "Potencia instalada mensual (MW)")
    st.caption("Potencia instalada del SEIN por tecnología (antes columnas S, T, U, V de Energía_Mensual del Excel). "
               "Power BI usa Hidro y Eólica para calcular el **Factor de Planta**: un mes sin potencia no se grafica. "
               "Edita las celdas o agrega un mes nuevo y luego presiona **Guardar cambios**.")
    editor_mensual("Potencia_Instalada", "pot", copiar_mes_anterior=True,
                   etiquetas={"Hidro_MW": "Hidro (MW)", "Termo_MW": "Termoeléctrica (MW)",
                              "Eolica_MW": "Eólica (MW)", "Solar_MW": "Solar (MW)"})

st.markdown(f'<div class="subtitulo-app" style="margin-top:2rem">© BD Despacho SEIN · USGE · '
            f'{datetime.now(ZoneInfo("America/Lima")):%d/%m/%Y %H:%M}</div>', unsafe_allow_html=True)
