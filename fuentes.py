"""
Lectura de las fuentes y conversión al formato de cada tabla de Supabase.

- Despacho  : IEOD del COES (hoja DESPACHO_EJECUTADO)          -> tabla "Despacho"
- CMG       : CMg del COES (hoja Cmg_Barra) solo SANTA ROSA 220 -> tabla "CMG"
- Demanda   : DemandaCOES.xlsx (FECHA, EJECUTADO)              -> tabla "Demanda_E"
- Temperat. : table.csv del SENAMHI (promedio mensual)         -> tabla "Temperatura"
"""
from __future__ import annotations

import io
import zipfile
from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Setiembre", "Octubre", "Noviembre", "Diciembre"]
MES_CORTO = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

BASE_IEOD = ("https://www.coes.org.pe/portal/browser/download?"
             "url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
             "{y}%2F{m}_{M}%2F{d}%2FAnexoA_{ddmm}.xlsx")
BASE_CMG = ("https://www.coes.org.pe/portal/browser/download?url=Post%20Operaci%C3%B3n%2FReportes%2FIEOD%2F"
            "{y}%2F{m}_{M}%2F{d}%2FCMg{y}{m}{d}.zip")

# Nombres que el COES usa distinto en algunos días
ALIAS_CENTRALES = {"MCH TUPURI": "TUPURI", "8AGOSTO": "8AGOSTO"}


def _norm(s) -> str:
    return " ".join(str(s).replace("\t", " ").split()).upper()


COLS_EXCLUIR = {
    _norm("POTENCIA ACTIVA EJECUTADA DE LAS UNIDADES DE GENERACIÓN DEL SEIN (MW)"),
    _norm("MW"), _norm("Unnamed: 0"), _norm("FECHA"), _norm("COL_1"), _norm("HORA"),
    _norm("GT11"), _norm("GT12"),
}


def _columnas_unicas(cols):
    conteo, nuevas = {}, []
    for c in cols:
        if c not in conteo:
            conteo[c] = 0
            nuevas.append(c)
        else:
            conteo[c] += 1
            nuevas.append(f"{c}_{conteo[c]}")
    return nuevas


def _fechas(ini: date, fin: date):
    d = ini
    while d <= fin:
        yield d
        d += timedelta(days=1)


def _hora_a_minutos(h) -> int | None:
    """'00:30', '00:30:00', time(0,30), '24:00' -> minutos del día."""
    try:
        if hasattr(h, "hour"):
            return h.hour * 60 + h.minute
        p = str(h).strip().split(":")
        return int(p[0]) * 60 + int(p[1])
    except Exception:
        return None


# ===========================================================================
# DESPACHO EJECUTADO (IEOD)
# ===========================================================================
def leer_ieod_dia(dia: date) -> pd.DataFrame | None:
    y, m, d = dia.year, f"{dia.month:02d}", f"{dia.day:02d}"
    url = BASE_IEOD.format(y=y, m=m, d=d, M=MESES[dia.month - 1], ddmm=f"{d}{m}")
    try:
        r = requests.get(url, verify=False, timeout=60)
        r.raise_for_status()
        xls = pd.ExcelFile(BytesIO(r.content))
        if "DESPACHO_EJECUTADO" not in xls.sheet_names:
            return None
        raw = pd.read_excel(xls, sheet_name="DESPACHO_EJECUTADO", header=None)
    except Exception:
        return None

    headers = []
    for col in range(raw.shape[1]):
        partes = [str(raw.iloc[i, col]) for i in (4, 8, 9) if not pd.isna(raw.iloc[i, col])]
        h = " | ".join(p for p in partes if p.strip())
        headers.append(h if h else f"COL_{col}")
    validas = []
    for i, h in enumerate(headers):
        validas.append(i)
        if h.strip() == "MW":
            break
    df = raw.iloc[10:, validas].copy()
    df.columns = _columnas_unicas([headers[i] for i in validas])
    df = df.dropna(how="all").reset_index(drop=True)
    df["__DIA__"] = dia
    return df


def despacho_a_bd(df_raw: pd.DataFrame, mapeo: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    Convierte el IEOD crudo al formato de la tabla "Despacho".
    mapeo: DataFrame con nombre_coes / columna_bd (tabla central_mapeo).
    Devuelve (df listo para Supabase, lista de centrales nuevas no mapeadas).
    """
    # nombre normalizado de la central -> columna real del IEOD
    central_map: dict[str, str] = {}
    for col in df_raw.columns:
        central = col.split("|", 1)[1].strip() if "|" in col else col
        central_map[_norm(central)] = col
    for col in df_raw.columns:
        if col.endswith("| FENIX "):
            central_map[_norm("FENIX GT11")] = col
        if col.endswith("| FENIX .1"):
            central_map[_norm("FENIX GT12")] = col

    # columna de hora
    col_hora = central_map.get(_norm("COL_1")) or ("COL_1" if "COL_1" in df_raw.columns else None)
    if col_hora is None:
        raise ValueError("No se encontró la columna de HORA en el IEOD.")

    columnas = {"__DIA__": df_raw["__DIA__"], "__HORA__": df_raw[col_hora]}
    usadas = {col_hora, "__DIA__"}
    for _, fila in mapeo.iterrows():
        nombre, destino = fila["nombre_coes"], fila["columna_bd"]
        col = central_map.get(_norm(nombre)) or central_map.get(_norm(ALIAS_CENTRALES.get(nombre, nombre)))
        if col and col not in usadas:
            columnas[destino] = pd.to_numeric(df_raw[col], errors="coerce")
            usadas.add(col)
        else:
            columnas[destino] = pd.Series(float("nan"), index=df_raw.index)
    out = pd.DataFrame(columnas)

    # centrales que el COES trae y no están en el mapeo
    nuevas = []
    for col in df_raw.columns:
        if col in usadas:
            continue
        central = col.split("|", 1)[1].strip() if "|" in col else col
        if _norm(central) not in COLS_EXCLUIR and not _norm(central).startswith("COL_"):
            nuevas.append(" ".join(central.replace("\t", " ").split()))

    # FECHA = día + hora ; "00:00" (o "24:00") es el cierre del día -> día siguiente
    minutos = out["__HORA__"].map(_hora_a_minutos)
    out = out[minutos.notna()].copy()
    minutos = minutos[minutos.notna()].astype(int)
    minutos = minutos.where(minutos != 0, 24 * 60)
    out.insert(0, "FECHA", pd.to_datetime(out["__DIA__"]) + pd.to_timedelta(minutos, unit="m"))
    out = out.drop(columns=["__DIA__", "__HORA__"])
    datos = [c for c in out.columns if c != "FECHA"]
    out[datos] = out[datos].astype(float).fillna(0)
    out["FECHA"] = out["FECHA"].dt.round("30min")
    out = out.drop_duplicates("FECHA", keep="last").sort_values("FECHA").reset_index(drop=True)
    return out, sorted(set(nuevas))


# ===========================================================================
# CMG (solo SANTA ROSA 220)
# ===========================================================================
def leer_cmg_dia(dia: date) -> pd.DataFrame | None:
    y, m, d = dia.year, f"{dia.month:02d}", f"{dia.day:02d}"
    url = BASE_CMG.format(y=y, m=m, d=d, M=MESES[dia.month - 1])
    try:
        r = requests.get(url, verify=False, timeout=60)
        r.raise_for_status()
        z = zipfile.ZipFile(BytesIO(r.content))
        nombre = next((f for f in z.namelist() if f.endswith(".xlsx") and "CMgCP" in f), None)
        if nombre is None:
            return None
        with z.open(nombre) as f:
            raw = pd.read_excel(f, sheet_name="Cmg_Barra", header=None)
    except Exception:
        return None

    header = raw.iloc[2]
    validas = [0]
    for i in range(1, len(header)):
        if pd.isna(header[i]):
            break
        validas.append(i)
    cols = _columnas_unicas([str(header[i]) if not pd.isna(header[i]) else f"COL_{i}" for i in validas])
    df = raw.iloc[3:51, validas].copy()
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)
    df["__DIA__"] = dia
    return df


def cmg_a_bd(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Deja solo FECHA (cada 30 min) y SANTA ROSA 220."""
    col_sr = next((c for c in df_raw.columns if _norm(c).startswith("SANTA ROSA 220")), None)
    if col_sr is None:
        raise ValueError("No se encontró la barra SANTA ROSA 220 en el archivo de CMg.")

    # columna con fecha/hora (la que más valores convertibles tiene)
    mejor, mejor_n = None, 0
    for c in df_raw.columns:
        if c in (col_sr, "__DIA__"):
            continue
        n = pd.to_datetime(df_raw[c], errors="coerce").notna().sum()
        if n > mejor_n:
            mejor, mejor_n = c, n

    out = pd.DataFrame({"__DIA__": df_raw["__DIA__"]})
    if mejor is not None and mejor_n >= len(df_raw) * 0.8:
        ts = pd.to_datetime(df_raw[mejor], errors="coerce")
        tiene_fecha = ts.dt.year.fillna(1900).gt(2000).mean() > 0.8
        if tiene_fecha:
            fecha = ts
        else:  # solo trae la hora -> se une con el día
            fecha = pd.to_datetime(out["__DIA__"]) + pd.to_timedelta(ts.dt.hour * 60 + ts.dt.minute, unit="m")
            fecha = fecha.where(~((ts.dt.hour == 0) & (ts.dt.minute == 0)), fecha + pd.Timedelta(days=1))
    else:  # sin columna de hora: 48 filas = 00:30 … 24:00
        orden = out.groupby("__DIA__").cumcount() + 1
        fecha = pd.to_datetime(out["__DIA__"]) + pd.to_timedelta(orden * 30, unit="m")

    # 23:59 -> 00:00 del día siguiente
    fecha = pd.to_datetime(fecha)
    es_2359 = (fecha.dt.hour == 23) & (fecha.dt.minute == 59)
    fecha = fecha.where(~es_2359, fecha + pd.Timedelta(minutes=1))

    res = pd.DataFrame({"FECHA": fecha.dt.round("30min"),
                        "SANTA ROSA 220": pd.to_numeric(df_raw[col_sr], errors="coerce")})
    res = res.dropna(subset=["FECHA", "SANTA ROSA 220"])
    return res.drop_duplicates("FECHA", keep="last").sort_values("FECHA").reset_index(drop=True)


# ===========================================================================
# DEMANDA (DemandaCOES.xlsx)
# ===========================================================================
def demanda_desde_excel(archivo) -> pd.DataFrame:
    raw = pd.read_excel(archivo, header=None)
    # buscar la fila de encabezados (la que contiene FECHA y EJECUTADO)
    fila = next(i for i in range(min(30, len(raw)))
                if {"FECHA", "EJECUTADO"} <= {_norm(v) for v in raw.iloc[i].tolist()})
    df = raw.iloc[fila + 1:].copy()
    df.columns = [_norm(v) for v in raw.iloc[fila].tolist()]
    df = df[["FECHA", "EJECUTADO"]]

    def parse(v):
        if isinstance(v, pd.Timestamp) or hasattr(v, "year"):
            return pd.Timestamp(v)
        return pd.to_datetime(str(v).strip(), format="%d/%m/%Y %H:%M", errors="coerce")

    df["FECHA"] = df["FECHA"].map(parse)
    df["EJECUTADO"] = pd.to_numeric(df["EJECUTADO"], errors="coerce")
    df = df.dropna(subset=["FECHA", "EJECUTADO"])
    df["FECHA"] = pd.to_datetime(df["FECHA"]).dt.round("30min")
    return df.drop_duplicates("FECHA", keep="last").sort_values("FECHA").reset_index(drop=True)


# ===========================================================================
# TEMPERATURA (table.csv SENAMHI) -> promedio mensual
# ===========================================================================
def temperatura_desde_csv(archivo) -> pd.DataFrame:
    contenido = archivo.read() if hasattr(archivo, "read") else open(archivo, "rb").read()
    for enc in ("utf-8-sig", "latin-1"):
        try:
            texto = contenido.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lineas = texto.splitlines()
    fila = next(i for i, l in enumerate(lineas) if "TEMPERATURA" in l.upper())
    df = pd.read_csv(io.StringIO("\n".join(lineas[fila:])))
    col_fecha = df.columns[0]
    col_temp = next(c for c in df.columns if "TEMPERATURA" in c.upper())

    df["_f"] = pd.to_datetime(df[col_fecha].astype(str).str.strip(), format="%Y/%m/%d", errors="coerce")
    df["_t"] = pd.to_numeric(df[col_temp], errors="coerce")  # "S/D" -> vacío
    df = df.dropna(subset=["_f", "_t"])

    res = (df.groupby([df["_f"].dt.year.rename("AÑO"), df["_f"].dt.month.rename("MES")])["_t"]
             .agg(["mean", "count"]).reset_index())
    res["MES_NOMBRE"] = res["MES"].map(lambda m: MES_CORTO[m - 1])
    res["TEMP_MEDIA"] = res["mean"].round(1)
    res = res.rename(columns={"count": "Horas"})
    return res[["AÑO", "MES", "MES_NOMBRE", "TEMP_MEDIA", "Horas"]]
