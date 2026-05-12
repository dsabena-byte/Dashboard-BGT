#!/usr/bin/env python3
"""
Descarga el Excel publicado en SharePoint via link anonimo y lo convierte
a data.json para el dashboard.

Requiere las variables de entorno:
  SHAREPOINT_URL  -> link "anyone with the link" generado desde SharePoint
                     (ej: https://contoso.sharepoint.com/:x:/s/site/abc?e=xxx)
  SHAREPOINT_SHEET (opcional) -> nombre de la hoja a leer. Si no se setea,
                     el script escanea todas las hojas del workbook y usa
                     la primera donde encuentre los headers esperados.

Headers esperados (case-insensitive, acepta sinonimos):
  presupuesto, nombre cta, año, mes, concepto, $, usd
"""

from __future__ import annotations

import io
import json
import os
import sys
from pathlib import Path
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode

import requests
from openpyxl import load_workbook


REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT = REPO_ROOT / "data.json"

# Mapeo header_normalizado -> campo interno.
# Notar: "cuenta" sola NO mapea aca porque en algunas planillas "CUENTA"
# es un codigo numerico (ej 6202010072), mientras que la cuenta que el
# dashboard necesita es el NOMBRE de la cuenta (ej PUBLICIDAD EXHIBICIONES).
COLUMN_ALIASES = {
    "presupuesto": "presupuesto",
    "tipo": "presupuesto",
    "tipo de presupuesto": "presupuesto",
    # OJO: "cuenta" sola se mapea a un campo distinto a proposito, porque en
    # planillas tipo "Seguimiento BGT" la columna "CUENTA" es un codigo numerico
    # (ej 6202010072) y el dashboard usa el nombre legible ("NOMBRE CTA").
    "cuenta": "_cuenta_codigo",
    "nombre cta": "cuenta",
    "nombre cuenta": "cuenta",
    "nombre de cuenta": "cuenta",
    "categoria": "cuenta",
    "categoria contable": "cuenta",
    "anio": "anio",
    "ano": "anio",
    "año": "anio",
    "year": "anio",
    "mes": "mes",
    "month": "mes",
    "concepto": "concepto",
    "subcuenta": "concepto",
    "$": "ars",
    "ars": "ars",
    "importe ars": "ars",
    "monto ars": "ars",
    "pesos": "ars",
    "usd": "usd",
    "u$s": "usd",
    "importe usd": "usd",
    "monto usd": "usd",
    "dolares": "usd",
    "dólares": "usd",
}

REQUIRED = ["presupuesto", "cuenta", "anio", "mes", "concepto", "ars", "usd"]

MESES_VALIDOS = {
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
}

MIN_MATCH = 5  # cuantas columnas REQUIRED tiene que matchear una fila para tratarla como headers
HEADER_SCAN_ROWS = 25


def build_download_url(share_url: str) -> str:
    parsed = urlparse(share_url)
    query = dict(parse_qsl(parsed.query))
    query["download"] = "1"
    return urlunparse(parsed._replace(query=urlencode(query)))


def download_excel(share_url: str) -> bytes:
    url = build_download_url(share_url)
    resp = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0 (SharePointSync Dashboard-BGT)"},
        allow_redirects=True,
        timeout=60,
    )
    resp.raise_for_status()
    if resp.content[:4] != b"PK\x03\x04":
        ctype = resp.headers.get("Content-Type", "")
        raise RuntimeError(
            "La URL no devolvio un .xlsx. Verifica que el link sea publico "
            f"('anyone with the link'). Content-Type: {ctype}"
        )
    return resp.content


def normalize_header(h) -> str:
    key = (str(h) if h is not None else "").strip().lower()
    return COLUMN_ALIASES.get(key, key)


def find_header_row(rows: list[tuple]) -> tuple[int, list[str]] | None:
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        normalized = [normalize_header(c) for c in row]
        matches = sum(1 for c in REQUIRED if c in normalized)
        if matches >= MIN_MATCH:
            return i, normalized
    return None


def parse_sheet(ws, sheet_name: str) -> tuple[list[list] | None, str]:
    """Devuelve (rows, log_msg). rows=None si la hoja no se pudo parsear."""
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return None, f"hoja '{sheet_name}': vacia"

    found = find_header_row(all_rows)
    if not found:
        preview = [
            [("" if c is None else str(c))[:25] for c in r[:8]]
            for r in all_rows[:3]
        ]
        return None, (
            f"hoja '{sheet_name}' ({len(all_rows)} filas): "
            f"no se encontraron headers. Primeras filas: {preview}"
        )

    header_idx, headers = found
    missing = [c for c in REQUIRED if c not in headers]
    if missing:
        return None, (
            f"hoja '{sheet_name}': headers en fila {header_idx+1} pero "
            f"faltan columnas {missing}. Encontradas: {headers}"
        )

    idx = {col: headers.index(col) for col in REQUIRED}
    out: list[list] = []
    skipped = 0
    for row in all_rows[header_idx + 1:]:
        if row is None or all(v is None or v == "" for v in row):
            continue
        try:
            presupuesto = str(row[idx["presupuesto"]] or "").strip()
            cuenta = str(row[idx["cuenta"]] or "").strip()
            anio_raw = row[idx["anio"]]
            anio = str(anio_raw if anio_raw is not None else "").strip().split(".")[0]
            mes = str(row[idx["mes"]] or "").strip().upper()
            concepto = str(row[idx["concepto"]] or "").strip()
            ars_raw = row[idx["ars"]]
            usd_raw = row[idx["usd"]]
        except IndexError:
            skipped += 1
            continue

        if not presupuesto or presupuesto.lower() == "none":
            skipped += 1
            continue
        if mes not in MESES_VALIDOS:
            skipped += 1
            continue
        try:
            ars_num = float(ars_raw) if ars_raw not in (None, "", "-") else 0.0
            usd_num = float(usd_raw) if usd_raw not in (None, "", "-") else 0.0
        except (TypeError, ValueError):
            skipped += 1
            continue

        out.append([
            presupuesto, cuenta, anio, mes, concepto,
            round(ars_num), round(usd_num),
        ])

    msg = (
        f"hoja '{sheet_name}': headers en fila {header_idx+1}, "
        f"{len(out)} filas validas, {skipped} descartadas"
    )
    return out, msg


def parse_workbook(blob: bytes, sheet_name: str | None) -> list[list]:
    wb = load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    print(f"  Hojas en el workbook: {wb.sheetnames}")

    sheets_to_try = [sheet_name] if sheet_name else wb.sheetnames
    errors = []

    for sn in sheets_to_try:
        if sn not in wb.sheetnames:
            errors.append(f"hoja '{sn}' no existe en el workbook")
            continue
        rows, log = parse_sheet(wb[sn], sn)
        print(f"  -> {log}")
        if rows is not None and len(rows) > 0:
            return rows
        errors.append(log)

    raise RuntimeError(
        "No se pudo extraer datos de ninguna hoja. Detalle:\n  - "
        + "\n  - ".join(errors)
    )


def main() -> int:
    share_url = os.environ.get("SHAREPOINT_URL")
    if not share_url:
        print("ERROR: falta la variable de entorno SHAREPOINT_URL", file=sys.stderr)
        return 1
    sheet = os.environ.get("SHAREPOINT_SHEET") or None

    print("Descargando Excel desde SharePoint...")
    blob = download_excel(share_url)
    print(f"Descarga OK ({len(blob)} bytes). Parseando hoja={sheet or 'auto'}...")
    rows = parse_workbook(blob, sheet)
    print(f"Total: {len(rows)} filas validas.")

    payload = {
        "columns": ["presupuesto", "cuenta", "anio", "mes", "concepto", "ars", "usd"],
        "rows": rows,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"Escrito {OUTPUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
