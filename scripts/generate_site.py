"""
Genera la página estática de contabilidad del condominio a partir de un
Google Sheet, usando una Service Account de Google Cloud.

Se ejecuta diariamente vía GitHub Actions y sobreescribe docs/index.html.
"""

import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from jinja2 import Environment, FileSystemLoader

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SHEET_ID = "1t2hw6tfXEDH-rKabDTxlMWI7eoCi-FTG"
SHEET_TAB = "Ingresos"
BUILDING_NAME = "Zempoala 437"

# Rangos exactos dentro de la hoja "Ingresos"
RANGE_DUENOS = "C8:S24"        # Estado de cuenta por dueño
RANGE_GASTOS_MES = "C35:P43"   # Gastos mensuales del condominio
RANGE_NOTAS = "C45:D56"        # Notas por mes (incluye el título "Notas:" en C45)
RANGE_EXTRAORDINARIOS = "C62:O200"  # con margen de sobra para que la tabla crezca
RANGE_RESUMEN_ANUAL = "C80:E88"    # Resumen de gastos anualizados

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS_DIR = os.path.join(BASE_DIR, "docs")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")


# ---------------------------------------------------------------------------
# Conexión a Google Sheets
# ---------------------------------------------------------------------------

def get_worksheet():
    creds_raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
    if not creds_raw:
        raise RuntimeError(
            "No se encontró la variable de entorno GCP_SERVICE_ACCOUNT_JSON"
        )
    creds_info = json.loads(creds_raw)
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    return sh.worksheet(SHEET_TAB)


# ---------------------------------------------------------------------------
# Helpers de parseo
# ---------------------------------------------------------------------------

def clean_money(value: str) -> str:
    """Deja el valor de moneda tal cual viene (ya trae formato $ del sheet)."""
    return (value or "").strip()


def is_zero_or_empty(value: str) -> bool:
    v = (value or "").strip().replace("$", "").replace(",", "")
    return v in ("", "0", "0.0", "0.00", "-")


def parse_money(value: str) -> float:
    v = (value or "").strip().replace("$", "").replace(",", "")
    if v in ("", "-", "—"):
        return 0.0
    try:
        return float(v)
    except ValueError:
        return 0.0


def format_money(value: float) -> str:
    sign = "-" if value < 0 else ""
    return f"{sign}${abs(value):,.2f}"


def rows_to_table(values, header_rows=1):
    """Convierte una lista de listas (get() de gspread) en dict con
    'headers' y 'rows', normalizando el largo de cada fila al del header."""
    if not values:
        return {"headers": [], "rows": []}
    headers = values[0]
    width = len(headers)
    data_rows = []
    for row in values[header_rows:]:
        padded = row + [""] * (width - len(row))
        data_rows.append(padded[:width])
    return {"headers": headers, "rows": data_rows}


def build_dueños_table(values):
    """C8:S24 -> header, fila de totales, dueños, fila 'Total Ingresos'.

    No se expone el nombre del dueño(a) en la página pública: se quita esa
    columna y cada quien se identifica solo por su número de depto.
    """
    table = rows_to_table(values, header_rows=1)
    rows = table["rows"]
    headers = table["headers"]

    # La primera fila de datos (fila 9 del sheet) es un renglón de referencia
    # sin nombre de dueño; la última ("Total Ingresos") es el total general.
    owner_rows = rows[1:-1] if len(rows) > 2 else []
    total_ingresos_row = rows[-1] if rows else None

    # Anotar cada fila de dueño con si debe algo (columna "Debe", índice 4
    # mientras la columna de nombre sigue presente).
    debe_idx = 4
    for r in owner_rows:
        r_debe = r[debe_idx] if len(r) > debe_idx else ""
        r.append("al-corriente" if is_zero_or_empty(r_debe) else "con-adeudo")

    headers_public = headers[1:]  # quita "Dueño(a)"

    owner_rows_public = []
    for r in owner_rows:
        flag = r[-1]
        data_no_flag = r[:-1]
        owner_rows_public.append(data_no_flag[1:] + [flag])

    total_row_public = None
    if total_ingresos_row:
        total_row_public = total_ingresos_row[1:]
        # La celda de "Depto" en esta fila viene vacía en el sheet; ahí ponemos
        # la etiqueta ya que la columna de nombre (que la traía) se quitó.
        if total_row_public and not total_row_public[0].strip():
            total_row_public[0] = "Total Ingresos"

    return {
        "headers": headers_public,
        "owner_rows": owner_rows_public,
        "total_row": total_row_public,
    }


def build_gastos_mes_table(values):
    table = rows_to_table(values, header_rows=1)
    rows = table["rows"]
    total_row = None
    diferencia_row = None
    service_rows = []
    for r in rows:
        first_cell = (r[0] or "").strip().lower()
        if first_cell.startswith("total gastos"):
            total_row = r
        elif first_cell.startswith("diferencia"):
            diferencia_row = r
        elif first_cell:
            service_rows.append(r)
    return {
        "headers": table["headers"],
        "service_rows": service_rows,
        "total_row": total_row,
        "diferencia_row": diferencia_row,
    }


def build_notas(values):
    """C45:D56. Primera fila es el título 'Notas:' (se ignora), luego
    12 filas mes/comentario."""
    notas = []
    for row in values[1:]:
        mes = (row[0] if len(row) > 0 else "").strip()
        comentario = (row[1] if len(row) > 1 else "").strip()
        if not mes:
            continue
        tiene_nota = comentario not in ("", "-----", "-", "—")
        notas.append({"mes": mes, "comentario": comentario, "tiene_nota": tiene_nota})
    return notas


def build_extraordinarios_table(values):
    table = rows_to_table(values, header_rows=1)
    rows = table["rows"]
    total_row = None
    total_depto_row = None
    concept_rows = []
    for r in rows:
        first_cell = (r[0] or "").strip().lower()
        if first_cell == "total":
            total_row = r
        elif first_cell.startswith("total x depto"):
            total_depto_row = r
            break  # "Total x depto" siempre marca el final de esta tabla;
                   # cortamos aquí para no leer la siguiente tabla del sheet
                   # (el resumen anual), que vive más abajo en las mismas columnas.
        elif first_cell:
            concept_rows.append(r)
    return {
        "headers": table["headers"],
        "concept_rows": concept_rows,
        "total_row": total_row,
        "total_depto_row": total_depto_row,
    }


def build_resumen_anual(values):
    table = rows_to_table(values, header_rows=1)
    rows = table["rows"]
    total_row = None
    total_depto_row = None
    concept_rows = []
    for r in rows:
        first_cell = (r[0] or "").strip().lower()
        if first_cell == "total":
            total_row = r
        elif first_cell.startswith("total x depto"):
            total_depto_row = r
        elif first_cell:
            concept_rows.append(r)
    return {
        "headers": table["headers"],
        "concept_rows": concept_rows,
        "total_row": total_row,
        "total_depto_row": total_depto_row,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build_resumen_hero(dueños_table, gastos_mes_table):
    """Calcula los 3 números principales que se muestran arriba de la página."""
    total_row = dueños_table["total_row"] or []
    # headers ya sin nombre: Depto, Total, Pagado, Debe, Enero..Diciembre
    total_cuotas = parse_money(total_row[1] if len(total_row) > 1 else "0")
    total_pagado = parse_money(total_row[2] if len(total_row) > 2 else "0")
    total_debe = parse_money(total_row[3] if len(total_row) > 3 else "0")

    gastos_total_row = gastos_mes_table["total_row"] or []
    # headers: Servicio, Total anual, Enero..Diciembre
    total_gastos = parse_money(gastos_total_row[1] if len(gastos_total_row) > 1 else "0")

    remanente = total_pagado - total_gastos

    return {
        "total_cuotas": format_money(total_cuotas),
        "total_pagado": format_money(total_pagado),
        "total_debe": format_money(total_debe),
        "total_gastos": format_money(total_gastos),
        "remanente": format_money(remanente),
        "remanente_positivo": remanente >= 0,
    }


def main():
    ws = get_worksheet()

    dueños = build_dueños_table(ws.get(RANGE_DUENOS))
    gastos_mes = build_gastos_mes_table(ws.get(RANGE_GASTOS_MES))
    notas = build_notas(ws.get(RANGE_NOTAS))
    extraordinarios = build_extraordinarios_table(ws.get(RANGE_EXTRAORDINARIOS))
    resumen_anual = build_resumen_anual(ws.get(RANGE_RESUMEN_ANUAL))
    resumen_hero = build_resumen_hero(dueños, gastos_mes)

    now_cdmx = datetime.now(ZoneInfo("America/Mexico_City"))
    fecha_actualizacion = now_cdmx.strftime("%d de %B de %Y, %H:%M hrs (CDMX)")

    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
    template = env.get_template("index.html.j2")

    html = template.render(
        building_name=BUILDING_NAME,
        fecha_actualizacion=fecha_actualizacion,
        dueños=dueños,
        gastos_mes=gastos_mes,
        notas=notas,
        extraordinarios=extraordinarios,
        resumen_anual=resumen_anual,
        resumen_hero=resumen_hero,
    )

    os.makedirs(DOCS_DIR, exist_ok=True)
    with open(os.path.join(DOCS_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Página generada correctamente: {fecha_actualizacion}")


if __name__ == "__main__":
    main()
