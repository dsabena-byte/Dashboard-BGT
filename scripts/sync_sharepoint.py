#!/usr/bin/env python3
"""
Descarga el Excel publicado en SharePoint via link anonimo y lo convierte
a data.json para el dashboard.

Requiere las variables de entorno:
  SHAREPOINT_URL  -> link "anyone with the link" generado desde SharePoint
                     (ej: https://contoso.sharepoint.com/:x:/s/site/abc?e=xxx)
  SHAREPOINT_SHEET (opcional) -> nombre de la hoja a leer. Default: primera hoja.

La planilla debe tener una fila de encabezados con estas columnas (en cualquier orden):
  presupuesto, cuenta, anio, mes, concepto, ars, usd

Acepta tambien variantes en castellano:
  ano / año -> anio
  Tipo / Tipo de presupuesto -> presupuesto
  Importe ARS / ARS -> ars
  Importe USD / USD -> usd
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

COLUMN_ALIASES = {
    "presupuesto": "presupuesto",
    "tipo": "presupuesto",
    "tipo de presupuesto": "presupuesto",
    "cuenta": "cuenta",
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
    "detalle": "concepto",
    "ars": "ars",
    "importe ars": "ars",
    "monto ars": "ars",
    "pesos": "ars",
    "usd": "usd",
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


def build_download_url(share_url: str) -> str:
    """Convierte una URL de share de SharePoint en una URL de descarga directa."""
    parsed = urlparse(share_url)
    query = dict(parse_qsl(parsed.query))
    query["download"] = "1"
    return urlunparse(parsed._replace(query=urlencode(query)))


def download_excel(share_url: str) -> bytes:
    url = build_download_url(share_url)
    headers = {
        "User-Agent": "Mozilla/5.0 (SharePointSync Dashboard-BGT)",
    }
    resp = requests.get(url, headers=headers, allow_redirects=True, timeout=60)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "")
    if "html" in ctype.lower() and not resp.content[:4] == b"PK\x03\x04":
        raise RuntimeError(
            "La URL devolvio HTML en vez del archivo. "
            "Verifica que el link sea publico ('anyone with the link') y que "
            "el tenant permita compartir externamente. Content-Type: " + ctype
        )
    return resp.content


def normalize_header(h: str) -> str:
    key = (h or "").strip().lower()
    return COLUMN_ALIASES.get(key, key)


def parse_workbook(blob: bytes, sheet_name: str | None) -> list[list]:
    wb = load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    rows_iter = ws.iter_rows(values_only=True)
    headers_raw = next(rows_iter)
    headers = [normalize_header(str(h) if h is not None else "") for h in headers_raw]

    missing = [c for c in REQUIRED if c not in headers]
    if missing:
        raise RuntimeError(
            f"Faltan columnas en la planilla: {missing}. "
            f"Encabezados encontrados: {headers}"
        )

    idx = {col: headers.index(col) for col in REQUIRED}
    out: list[list] = []
    for row in rows_iter:
        if row is None or all(v is None or v == "" for v in row):
            continue
        try:
            presupuesto = str(row[idx["presupuesto"]]).strip()
            cuenta = str(row[idx["cuenta"]]).strip()
            anio = str(row[idx["anio"]]).strip().split(".")[0]
            mes = str(row[idx["mes"]]).strip().upper()
            concepto = str(row[idx["concepto"]]).strip()
            ars = row[idx["ars"]]
            usd = row[idx["usd"]]
        except (IndexError, AttributeError):
            continue

        if not presupuesto or presupuesto.lower() == "none":
            continue
        if mes not in MESES_VALIDOS:
            continue
        ars_num = float(ars) if ars not in (None, "") else 0.0
        usd_num = float(usd) if usd not in (None, "") else 0.0
        out.append([
            presupuesto, cuenta, anio, mes, concepto,
            round(ars_num), round(usd_num),
        ])
    return out


def main() -> int:
    share_url = os.environ.get("SHAREPOINT_URL")
    if not share_url:
        print("ERROR: falta la variable de entorno SHAREPOINT_URL", file=sys.stderr)
        return 1
    sheet = os.environ.get("SHAREPOINT_SHEET") or None

    print(f"Descargando Excel desde SharePoint...")
    blob = download_excel(share_url)
    print(f"Descarga OK ({len(blob)} bytes). Parseando hoja={sheet or 'primera'}...")
    rows = parse_workbook(blob, sheet)
    print(f"Parseadas {len(rows)} filas validas.")

    payload = {
        "columns": ["presupuesto", "cuenta", "anio", "mes", "concepto", "ars", "usd"],
        "rows": rows,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"Escrito {OUTPUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
