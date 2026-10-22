"""
asientos.py
===========
Construcción de los asientos contables de la planilla de sueldos y salarios.
No genera archivos ni usa Streamlit: devuelve estructuras de datos que luego
dibuja ``comprobantes.py``.

Dos formatos:
* **Empresa privada**: asiento con cuentas Producción en Proceso / Gastos de
  Administración y Comercialización contra Sueldos por pagar, Gestora, CNS,
  RC-IVA, Aguinaldos y Previsión para indemnizaciones.
* **Empresa pública**: comprobante patrimonial con partida presupuestaria,
  código y cuenta contable.

Si no hay empleados de Producción (o de Administración), solo se incluyen las
cuentas del departamento que sí tiene personal. Las líneas con monto cero no
se muestran.

Los códigos de partida y de cuenta del sector público están en
``PARTIDAS_PUBLICAS`` y ``CUENTAS_PUBLICAS_PAGAR``: se pueden cambiar sin tocar
la lógica.
"""

from __future__ import annotations

import unicodedata
from decimal import Decimal
from typing import Any, Sequence

PRODUCCION = "PRODUCCION"
ADMIN = "ADMIN"
NOMBRE_DEPTO = {PRODUCCION: "Producción", ADMIN: "Administración y Comercialización"}
OPCIONES_DEPARTAMENTO = [NOMBRE_DEPTO[PRODUCCION], NOMBRE_DEPTO[ADMIN]]

CERO = Decimal("0.00")


def normalizar_departamento(valor: Any) -> str:
    """Devuelve ``PRODUCCION`` o ``ADMIN`` a partir del texto del Excel.

    Todo lo que contenga "produc" es Producción; cualquier otro valor (incluido
    "Comercialización" o vacío) se trata como Administración y Comercialización.
    """
    if valor is None or valor != valor:
        return ADMIN
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode().lower()
    return PRODUCCION if "produc" in texto else ADMIN


# --------------------------------------------------------------------------- #
# Sumas por departamento
# --------------------------------------------------------------------------- #
_CAMPOS = ("haber_basico", "bono_antiguedad", "pago_horas_extra", "pago_nocturno",
           "pago_domingos", "pago_feriados", "total_ganado", "patronal_cns",
           "patronal_solidario", "patronal_riesgo_profesional", "patronal_provivienda",
           "aportes_patronales", "aguinaldo", "indemnizacion")


def _suma(filas: Sequence[dict], clave: str) -> Decimal:
    return sum((Decimal(str(f[clave])) for f in filas), CERO)


def _por_departamento(filas: Sequence[dict]) -> dict[str, dict[str, Decimal]]:
    """{depto: {campo: suma}} solo para los departamentos con empleados."""
    out: dict[str, dict[str, Decimal]] = {}
    for depto in (PRODUCCION, ADMIN):
        grupo = [f for f in filas if f.get("departamento", ADMIN) == depto]
        if grupo:
            out[depto] = {k: _suma(grupo, k) for k in _CAMPOS}
    return out


def _globales(filas: Sequence[dict]) -> dict[str, Decimal]:
    claves = ("total_ganado", "aportes_laborales", "rciva_retener", "liquido_pagable",
              "patronal_cns", "aportes_patronales", "aguinaldo", "indemnizacion")
    return {k: _suma(filas, k) for k in claves}


def _cerrar(lineas: list[dict]) -> dict[str, Any]:
    debe = sum((l["debe"] for l in lineas if l.get("debe") is not None), CERO)
    haber = sum((l["haber"] for l in lineas if l.get("haber") is not None), CERO)
    return {"lineas": lineas, "total_debe": debe, "total_haber": haber,
            "cuadra": debe == haber}


# --------------------------------------------------------------------------- #
# Empresa privada
# --------------------------------------------------------------------------- #
def asiento_privado(filas: Sequence[dict]) -> dict[str, Any]:
    """Asiento de la planilla para empresa privada.

    Cada línea: ``cuenta``, ``nivel`` (0 = cuenta principal, 1 = detalle),
    ``parcial``, ``debe``, ``haber``, ``subrayar``.
    """
    lineas: list[dict] = []
    for depto, s in _por_departamento(filas).items():
        titulo = ("Producción en Proceso" if depto == PRODUCCION
                  else "Gastos de Administración y Comercialización")
        detalle = [("Sueldos y Salarios", s["total_ganado"]),
                   ("Aportes Patronales", s["aportes_patronales"]),
                   ("Aguinaldos", s["aguinaldo"]),
                   ("Indemnizaciones", s["indemnizacion"])]
        lineas.append({"cuenta": titulo, "nivel": 0, "parcial": None,
                       "debe": sum((m for _, m in detalle), CERO), "haber": None,
                       "subrayar": True})
        for nombre, monto in detalle:
            lineas.append({"cuenta": nombre, "nivel": 1, "parcial": monto, "debe": None,
                           "haber": None, "subrayar": False})

    g = _globales(filas)
    gestora_patronal = g["aportes_patronales"] - g["patronal_cns"]
    lineas.append({"cuenta": "Sueldos y Salarios por pagar", "nivel": 2, "parcial": None,
                   "debe": None, "haber": g["liquido_pagable"], "subrayar": True})
    lineas.append({"cuenta": "Aportes a la Gestora por Pagar", "nivel": 2, "parcial": None,
                   "debe": None, "haber": g["aportes_laborales"] + gestora_patronal,
                   "subrayar": True})
    lineas.append({"cuenta": "Aportes Laborales", "nivel": 3, "parcial": g["aportes_laborales"],
                   "debe": None, "haber": None, "subrayar": False})
    lineas.append({"cuenta": "Aportes Patronales", "nivel": 3, "parcial": gestora_patronal,
                   "debe": None, "haber": None, "subrayar": False})
    lineas.append({"cuenta": "C.N.S. por pagar", "nivel": 2, "parcial": None, "debe": None,
                   "haber": g["patronal_cns"], "subrayar": True})
    if g["rciva_retener"] > 0:
        lineas.append({"cuenta": "RC-IVA dependientes por pagar", "nivel": 2, "parcial": None,
                       "debe": None, "haber": g["rciva_retener"], "subrayar": True})
    lineas.append({"cuenta": "Aguinaldos por pagar", "nivel": 2, "parcial": None,
                   "debe": None, "haber": g["aguinaldo"], "subrayar": True})
    lineas.append({"cuenta": "Previsión para indemnizaciones", "nivel": 2, "parcial": None,
                   "debe": None, "haber": g["indemnizacion"], "subrayar": True})
    return _cerrar(lineas)


# --------------------------------------------------------------------------- #
# Empresa pública (comprobante patrimonial)
# --------------------------------------------------------------------------- #
# (rubro, partida presupuestaria, clave del monto, cuenta Producción, cuenta Administración)
# El orden agrupa las partidas que van a una misma cuenta, para poder mostrar
# los montos parciales y el total de la cuenta.
PARTIDAS_PUBLICAS = [
    ("11700", "Sueldos", "sueldos", ("1155", "Inventario de Producción en Proceso"),
     ("611", "Sueldos y Salarios")),
    ("11220", "Bono de Antigüedad", "bono_antiguedad", ("1155", "Inventario de Producción en Proceso"),
     ("611", "Sueldos y Salarios")),
    ("11910", "Horas extraordinarias", "extraordinarias",
     ("1155", "Inventario de Producción en Proceso"), ("611", "Sueldos y Salarios")),
    ("11400", "Aguinaldo", "aguinaldo", ("1155", "Inventario de Producción en Proceso"),
     ("6133", "Otros Beneficios sociales")),
    ("13110", "Régimen de Corto Plazo Salud", "patronal_cns",
     ("1155", "Inventario de Producción en Proceso"), ("6121", "Aportes patronales al seguro social")),
    ("13120", "Prima de riesgo profesional Régimen de LP", "patronal_riesgo_profesional",
     ("1155", "Inventario de Producción en Proceso"), ("6123", "Otros aportes patronales")),
    ("13131", "Aporte Patronal Solidario", "patronal_solidario",
     ("1155", "Inventario de Producción en Proceso"), ("6123", "Otros aportes patronales")),
    ("13200", "Aporte patronal para vivienda", "patronal_provivienda",
     ("1155", "Inventario de Producción en Proceso"), ("6122", "Aportes patronales para vivienda")),
    ("94100", "Indemnización", "indemnizacion", ("1155", "Inventario de Producción en Proceso"),
     ("6131", "Indemnizaciones")),
]
# Se ordena para que las cuentas queden juntas dentro de cada departamento.
_ORDEN_ADMIN = ["11700", "11220", "11910", "11400", "13110", "13120", "13131", "13200", "94100"]

CUENTAS_PUBLICAS_PAGAR = {
    "sueldos": ("2113", "Sueldos y Salarios a pagar a corto plazo"),
    "patronales": ("21141", "Aportes patronales a pagar a corto plazo"),
    "retenciones": ("21142", "Retenciones a pagar a corto plazo"),
    "impuestos": ("21161", "Impuestos a pagar a corto plazo"),
    "aguinaldo": ("2141", "Previsión para Beneficios sociales a corto plazo"),
    "indemnizacion": ("2251", "Previsión para beneficios sociales a largo plazo"),
}


def asiento_publico(filas: Sequence[dict]) -> dict[str, Any]:
    """Líneas del comprobante patrimonial (sector público).

    Cada línea: ``rubro``, ``partida``, ``codigo``, ``cuenta``, ``parcial``,
    ``debe``, ``haber``. Cuando varias partidas consecutivas van a la misma
    cuenta, cada una muestra su monto en ``parcial`` y el total de la cuenta
    aparece en ``debe`` de la última.
    """
    lineas: list[dict] = []
    for depto, s in _por_departamento(filas).items():
        montos = {
            "sueldos": s["haber_basico"],
            "bono_antiguedad": s["bono_antiguedad"],
            "extraordinarias": s["pago_horas_extra"] + s["pago_nocturno"]
                               + s["pago_domingos"] + s["pago_feriados"],
            "aguinaldo": s["aguinaldo"], "patronal_cns": s["patronal_cns"],
            "patronal_riesgo_profesional": s["patronal_riesgo_profesional"],
            "patronal_solidario": s["patronal_solidario"],
            "patronal_provivienda": s["patronal_provivienda"],
            "indemnizacion": s["indemnizacion"],
        }
        crudas = []
        for rubro, partida, clave, c_prod, c_adm in PARTIDAS_PUBLICAS:
            monto = montos[clave]
            if monto <= 0:
                continue
            codigo, cuenta = c_prod if depto == PRODUCCION else c_adm
            crudas.append((rubro, partida, codigo, cuenta, monto))
        if depto == ADMIN:
            crudas.sort(key=lambda t: _ORDEN_ADMIN.index(t[0]))
        # agrupa partidas consecutivas con la misma cuenta
        i = 0
        while i < len(crudas):
            j = i
            while j + 1 < len(crudas) and crudas[j + 1][2] == crudas[i][2]:
                j += 1
            grupo = crudas[i:j + 1]
            total = sum((t[4] for t in grupo), CERO)
            for k, (rubro, partida, codigo, cuenta, monto) in enumerate(grupo):
                if len(grupo) == 1:
                    lineas.append({"rubro": rubro, "partida": partida, "codigo": codigo,
                                   "cuenta": cuenta, "parcial": None, "debe": monto,
                                   "haber": None})
                else:
                    lineas.append({"rubro": rubro, "partida": partida, "codigo": codigo,
                                   "cuenta": cuenta, "parcial": monto,
                                   "debe": total if k == len(grupo) - 1 else None,
                                   "haber": None})
            i = j + 1

    g = _globales(filas)
    for clave, monto in (("sueldos", g["liquido_pagable"]),
                         ("patronales", g["aportes_patronales"]),
                         ("retenciones", g["aportes_laborales"]),
                         ("impuestos", g["rciva_retener"]),
                         ("aguinaldo", g["aguinaldo"]),
                         ("indemnizacion", g["indemnizacion"])):
        if monto > 0:
            codigo, cuenta = CUENTAS_PUBLICAS_PAGAR[clave]
            lineas.append({"rubro": "", "partida": "", "codigo": codigo, "cuenta": cuenta,
                           "parcial": None, "debe": None, "haber": monto})
    return _cerrar(lineas)
