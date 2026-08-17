#!/usr/bin/env python3
"""
Valida si una API de vuelos devuelve REALMENTE numero de puerta y terminal
para los vuelos del grupo MABE Cancun (agosto 2026).

El punto no es "la API responde 200". El punto es si los campos `gate` y
`terminal` vienen con dato o vienen en null, que es lo que define si la app
de escala puede dibujar el recorrido sola o si el pasajero tiene que leer la
puerta de la app de la aerolinea.

Uso:
    pip install -r scripts/requirements.txt

    # Aviationstack (plan gratis: 100 consultas/mes, solo HTTP)
    export AVIATIONSTACK_KEY=xxxxx
    python3 scripts/test_gate_api.py

    # AeroDataBox via RapidAPI (plan gratis: 600 unidades/mes)
    export RAPIDAPI_KEY=xxxxx
    python3 scripts/test_gate_api.py

    # un vuelo puntual
    python3 scripts/test_gate_api.py CM320

Que mirar en la salida:
    - Si `gate` sale OK para vuelos del dia de hoy -> la API sirve.
    - Si `gate` sale NULL siempre -> la puerta hay que sacarla de la app de
      Copa y la nuestra solo la recibe cargada a mano. El modelo 1xx/2xx
      sigue funcionando igual, pero no se automatiza.
    - Anota tambien cuantas HORAS ANTES aparece la puerta. Ese numero decide
      si tiene sentido pollear la API o no.
"""

import json
import os
import sys
from datetime import date

import requests

# Vuelos del grupo que hacen escala en PTY, MDE o BOG.
VUELOS = [
    ("CM320", "COR->PTY", "Sabena, ida tramo 1"),
    ("CM324", "PTY->CUN", "Sabena, ida tramo 2 - escala 64 min"),
    ("CM326", "CUN->PTY", "regreso, 8 pax"),
    ("CM788", "PTY->COR", "Sabena, regreso tramo 2 - escala 68 min"),
    ("CM329", "PTY->EZE", "regreso grupo CATT6D - escala 45 min"),
    ("CM252", "COR->PTY", "ida Cordoba - escala 66 min"),
    ("CM122", "PTY->CUN", "ida, 8 pax convergen aca"),
    ("CM439", "EZE->PTY", "ida Ezeiza - escala 90 min"),
    ("AV112", "EZE->MDE", "ida Medellin - escala 60 min"),
    ("AV268", "MDE->CUN", "ida Medellin tramo 2"),
]

TIMEOUT = 20


def _cell(value):
    """Marca si el campo trae dato util o vino vacio."""
    if value in (None, "", "null"):
        return "NULL"
    return str(value)


def probe_aviationstack(key, flight_iata):
    """Aviationstack documenta gate/terminal/baggage. Plan gratis: solo HTTP."""
    url = "http://api.aviationstack.com/v1/flights"
    params = {"access_key": key, "flight_iata": flight_iata, "limit": 1}
    r = requests.get(url, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    payload = r.json()

    if "error" in payload:
        return {"error": payload["error"].get("message", json.dumps(payload["error"]))}

    data = payload.get("data") or []
    if not data:
        return {"error": "sin resultados para hoy"}

    f = data[0]
    dep, arr = f.get("departure") or {}, f.get("arrival") or {}
    return {
        "estado": f.get("flight_status"),
        "fecha": f.get("flight_date"),
        "dep_iata": dep.get("iata"),
        "dep_terminal": dep.get("terminal"),
        "dep_gate": dep.get("gate"),
        "arr_iata": arr.get("iata"),
        "arr_terminal": arr.get("terminal"),
        "arr_gate": arr.get("gate"),
    }


def probe_aerodatabox(key, flight_iata):
    """AeroDataBox: los campos de puerta figuran como 'a veces' disponibles."""
    hoy = date.today().isoformat()
    url = "https://aerodatabox.p.rapidapi.com/flights/number/{}/{}".format(flight_iata, hoy)
    headers = {"X-RapidAPI-Key": key, "X-RapidAPI-Host": "aerodatabox.p.rapidapi.com"}
    r = requests.get(url, headers=headers, params={"withLocation": "false"}, timeout=TIMEOUT)

    if r.status_code == 404:
        return {"error": "sin vuelo para hoy"}
    r.raise_for_status()

    data = r.json()
    if not isinstance(data, list) or not data:
        return {"error": "respuesta vacia"}

    f = data[0]
    dep, arr = f.get("departure") or {}, f.get("arrival") or {}
    return {
        "estado": f.get("status"),
        "fecha": hoy,
        "dep_iata": (dep.get("airport") or {}).get("iata"),
        "dep_terminal": dep.get("terminal"),
        "dep_gate": dep.get("gate"),
        "arr_iata": (arr.get("airport") or {}).get("iata"),
        "arr_terminal": arr.get("terminal"),
        "arr_gate": arr.get("gate"),
    }


PROVEEDORES = [
    ("aviationstack", "AVIATIONSTACK_KEY", probe_aviationstack),
    ("aerodatabox", "RAPIDAPI_KEY", probe_aerodatabox),
]


def correr(nombre, probe, key, vuelos):
    print("\n" + "=" * 78)
    print("PROVEEDOR: {}".format(nombre))
    print("=" * 78)
    print("{:<8} {:<10} {:<12} {:<9} {:<7} {:<9} {:<7}".format(
        "VUELO", "ESTADO", "RUTA", "DEP TERM", "DEP GT", "ARR TERM", "ARR GT"))
    print("-" * 78)

    con_gate = 0
    consultados = 0

    for iata, ruta, nota in vuelos:
        try:
            res = probe(key, iata)
        except requests.RequestException as exc:
            print("{:<8} ERROR RED: {}".format(iata, exc))
            continue
        except ValueError:
            print("{:<8} ERROR: respuesta no es JSON".format(iata))
            continue

        consultados += 1

        if "error" in res:
            print("{:<8} {}".format(iata, res["error"]))
            continue

        if res.get("dep_gate") or res.get("arr_gate"):
            con_gate += 1

        print("{:<8} {:<10} {:<12} {:<9} {:<7} {:<9} {:<7}".format(
            iata,
            _cell(res.get("estado"))[:10],
            "{}->{}".format(res.get("dep_iata") or "?", res.get("arr_iata") or "?"),
            _cell(res.get("dep_terminal"))[:9],
            _cell(res.get("dep_gate"))[:7],
            _cell(res.get("arr_terminal"))[:9],
            _cell(res.get("arr_gate"))[:7],
        ))

    print("-" * 78)
    if not consultados:
        print("VEREDICTO: no se pudo consultar ningun vuelo. Revisa la key.")
    elif con_gate == 0:
        print("VEREDICTO: 0 de {} vuelos trajeron puerta.".format(consultados))
        print("           La puerta hay que leerla de la app de Copa.")
        print("           La app de escala funciona igual, con carga manual.")
    else:
        print("VEREDICTO: {} de {} vuelos trajeron puerta.".format(con_gate, consultados))
        print("           Repetir el test a T-3h y a T-24h del vuelo para saber")
        print("           con cuanta anticipacion aparece el dato.")


def main():
    filtro = sys.argv[1].upper() if len(sys.argv) > 1 else None
    vuelos = [v for v in VUELOS if not filtro or v[0] == filtro]

    if not vuelos:
        print("No conozco el vuelo {}. Vuelos cargados: {}".format(
            filtro, ", ".join(v[0] for v in VUELOS)))
        return 1

    activos = [(n, p, os.environ.get(e)) for n, e, p in PROVEEDORES if os.environ.get(e)]

    if not activos:
        print(__doc__)
        print("Falta la key. Defini AVIATIONSTACK_KEY o RAPIDAPI_KEY y volve a correr.")
        return 1

    print("Probando {} vuelo(s) contra {} proveedor(es).".format(len(vuelos), len(activos)))
    print("Ojo: solo devuelve datos de vuelos que operan HOY.")

    for nombre, probe, key in activos:
        correr(nombre, probe, key, vuelos)

    print("\nRecorda anotar la hora de cada corrida. Lo que estamos midiendo es")
    print("cuantas horas antes de la salida aparece la puerta, no si la API anda.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
