"""
Conexión a Supabase (PostgreSQL) y escritura de datos.
Todas las escrituras son UPSERT: si la fecha ya existe se actualiza, nunca se duplica.
"""
from __future__ import annotations

import re

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

LOTE = 1000


# ---------------------------------------------------------------------------
# Conexión
# ---------------------------------------------------------------------------
def conectar(url: str):
    """Conecta aunque la contraseña tenga caracteres especiales (@ # / : ? ...)."""
    url = url.strip().strip('"').strip("'")
    m = re.match(r"^postgres(?:ql)?://([^:]+):(.*)@([^@/:]+):(\d+)/([^?]+)", url)
    if not m:
        # cadena en otro formato (p. ej. socket local para pruebas)
        return psycopg2.connect(url)
    user, pwd, host, port, db = m.groups()
    return psycopg2.connect(host=host, port=int(port), user=user, password=pwd,
                            dbname=db, sslmode="require", connect_timeout=20)


def q(ident: str) -> str:
    return '"' + ident.replace('"', '""') + '"'


def _filas(df: pd.DataFrame) -> list[tuple]:
    df = df.astype(object).where(pd.notna(df), None)
    return [tuple(r) for r in df.itertuples(index=False, name=None)]


def upsert(con, tabla: str, df: pd.DataFrame, pk: list[str]) -> int:
    """Inserta o actualiza df en public."tabla" usando pk como clave."""
    if df.empty:
        return 0
    cols = list(df.columns)
    set_cols = [c for c in cols if c not in pk]
    sql = (
        f"insert into public.{q(tabla)} ({', '.join(q(c) for c in cols)}) values %s "
        f"on conflict ({', '.join(q(c) for c in pk)}) do update set "
        + ", ".join(f"{q(c)} = excluded.{q(c)}" for c in set_cols)
    )
    filas = _filas(df)
    with con.cursor() as cur:
        for i in range(0, len(filas), LOTE):
            execute_values(cur, sql, filas[i:i + LOTE], page_size=LOTE)
    con.commit()
    return len(filas)


# ---------------------------------------------------------------------------
# Lecturas
# ---------------------------------------------------------------------------
def leer(con, sql: str, params=None) -> pd.DataFrame:
    with con.cursor() as cur:
        cur.execute(sql, params)
        cols = [d[0] for d in cur.description]
        return pd.DataFrame(cur.fetchall(), columns=cols)


def estado_tablas(con) -> pd.DataFrame:
    """Filas y última fecha de cada tabla (para el panel lateral)."""
    consultas = [
        ("Despacho", 'select count(*), max("FECHA") from public."Despacho"'),
        ("CMG", 'select count(*), max("FECHA") from public."CMG"'),
        ("Demanda_E", 'select count(*), max("FECHA") from public."Demanda_E"'),
        ("Temperatura", 'select count(*), max(make_date("AÑO","MES",1)) from public."Temperatura"'),
        ("Recursos", 'select count(*), max("Fecha") from public."Recursos"'),
        ("Potencia_Instalada", 'select count(*), max("Fecha") from public."Potencia_Instalada"'),
    ]
    out = []
    with con.cursor() as cur:
        for nombre, sql in consultas:
            cur.execute(sql)
            n, ult = cur.fetchone()
            out.append({"Tabla": nombre, "Filas": n, "Último dato": ult})
    return pd.DataFrame(out)


def mapeo_centrales(con) -> pd.DataFrame:
    return leer(con, "select nombre_coes, columna_bd, orden from public.central_mapeo order by orden")


def columnas_despacho(con) -> list[str]:
    df = leer(con, "select column_name from information_schema.columns "
                   "where table_schema='public' and table_name='Despacho' order by ordinal_position")
    return df["column_name"].tolist()


def agregar_centrales(con, nombres: list[str]) -> None:
    """Agrega centrales nuevas del COES como columnas de "Despacho" y al mapeo."""
    with con.cursor() as cur:
        cur.execute("select coalesce(max(orden), 0) from public.central_mapeo")
        orden = cur.fetchone()[0]
        for nombre in nombres:
            orden += 1
            cur.execute(f'alter table public."Despacho" add column if not exists {q(nombre)} double precision')
            cur.execute(
                "insert into public.central_mapeo (nombre_coes, columna_bd, orden) values (%s, %s, %s) "
                "on conflict (nombre_coes) do nothing",
                (nombre, nombre, orden),
            )
    con.commit()


def eliminar_mes(con, tabla: str, fecha) -> int:
    """Borra un mes de una tabla mensual ("Recursos" o "Potencia_Instalada")."""
    with con.cursor() as cur:
        cur.execute(f'delete from public.{q(tabla)} where "Fecha" = %s', (fecha,))
        n = cur.rowcount
    con.commit()
    return n


def eliminar_mes_recursos(con, fecha) -> int:
    """Borra un mes de la tabla "Recursos". Devuelve cuántas filas borró."""
    return eliminar_mes(con, "Recursos", fecha)
