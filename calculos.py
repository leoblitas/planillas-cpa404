"""
calculos.py
===========
Motor de cálculo de la planilla de sueldos y salarios y de la planilla
impositiva (RC-IVA) para Bolivia.

Diseño
------
* Sin dependencias externas (solo biblioteca estándar), para que pueda
  importarse desde Streamlit u otro módulo sin instalar nada adicional.
* Aritmética con ``Decimal`` y redondeo ``ROUND_HALF_UP`` a 2 decimales en
  cada monto (igual que una planilla en Excel). Evita errores de coma flotante.
* Ninguna constante legal está "quemada" en las funciones: todo llega por el
  diccionario de parámetros (hoja ``Parametros_LEY``) y por la tabla de
  antigüedad (hoja ``Antiguedad``). ``PARAMETROS_DEFECTO`` solo completa los
  parámetros que el Excel todavía no trae (ANS, aportes patronales, etc.).

Flujo típico (ver ``calcular_planilla_empleado``)
-------------------------------------------------
1. Bono de antigüedad          -> ``calcular_bono_antiguedad``
2. Componentes variables       -> ``calcular_componentes_variables``
3. Aportes laborales y RC-IVA  -> ``calcular_aportes_y_rciva``
4. Aportes patronales y previsiones -> ``calcular_aportes_patronales``,
   ``calcular_previsiones``
"""

from __future__ import annotations

import calendar
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Iterable, Mapping, Optional, Union

Numero = Union[int, float, str, Decimal]
Fecha = Union[date, datetime]

# --------------------------------------------------------------------------- #
# Parámetros por defecto (solo se usan si faltan en la hoja Parametros_LEY)
# --------------------------------------------------------------------------- #
PARAMETROS_DEFECTO: dict[str, Any] = {
    # Crédito fiscal por dependiente: 13% de N salarios mínimos (regla indicada: 1 SMN)
    "SMN_CREDITO_DEP": 1,
    # Aporte Nacional Solidario (ANS): umbrales (Bs) y tasas vigentes
    "ANS_T1": 13000, "ANS_T2": 25000, "ANS_T3": 35000,
    "ANS_P1": 0.0115, "ANS_P2": 0.0574, "ANS_P3": 0.1148,
    # Aportes patronales
    "AP_CNS": 0.10,      # Caja Nacional de Salud
    "AP_SOL": 0.035,     # Aporte solidario patronal
    "AP_RP": 0.0171,     # Prima de riesgo profesional
    "AP_PV": 0.02,       # Aporte pro-vivienda
    # Valores base de la hoja Parametros_LEY (por si no se cargan)
    "H_MES": 240,
    "AGUINALDO": 0.0833,
    "INDEMNIZACION": 0.0833,
}

DOS_DECIMALES = Decimal("0.01")


# --------------------------------------------------------------------------- #
# Utilidades numéricas
# --------------------------------------------------------------------------- #
def D(valor: Numero | None) -> Decimal:
    """Convierte a Decimal de forma segura (pasando por str para no arrastrar
    el error binario de los float). ``None`` o cadena vacía -> 0."""
    if valor is None or valor == "":
        return Decimal("0")
    if isinstance(valor, Decimal):
        return valor
    return Decimal(str(valor))


def redondear(valor: Numero) -> Decimal:
    """Redondea a 2 decimales (mitad hacia arriba)."""
    return D(valor).quantize(DOS_DECIMALES, rounding=ROUND_HALF_UP)


def _no_negativo(valor: Decimal, nombre: str) -> Decimal:
    if valor < 0:
        raise ValueError(f"'{nombre}' no puede ser negativo (recibido: {valor}).")
    return valor


# --------------------------------------------------------------------------- #
# Parámetros y fechas
# --------------------------------------------------------------------------- #
def cargar_parametros(filas: Union[Mapping[str, Any], Iterable[Mapping[str, Any]]]
                      ) -> dict[str, Any]:
    """Construye el diccionario de parámetros listo para usar.

    Acepta:
      * un dict ``{"SMN": 3300, ...}``, o
      * las filas de la hoja ``Parametros_LEY`` como lista de dicts con las
        columnas ``Codigo_Parametro`` y ``Valor`` (p. ej. ``df.to_dict("records")``).

    Los códigos ausentes se completan con ``PARAMETROS_DEFECTO``.
    """
    if isinstance(filas, Mapping):
        datos = dict(filas)
    else:
        datos = {f["Codigo_Parametro"]: f["Valor"] for f in filas}
    params = dict(PARAMETROS_DEFECTO)
    params.update(datos)
    if "SMN" not in params:
        raise KeyError("Falta el parámetro obligatorio 'SMN'.")
    return params


def ultimo_dia_periodo(anio: int, mes: int) -> date:
    """Último día del mes (fecha de corte para la antigüedad)."""
    return date(anio, mes, calendar.monthrange(anio, mes)[1])


def _a_fecha(f: Fecha) -> date:
    return f.date() if isinstance(f, datetime) else f


def anios_cumplidos(fecha_ingreso: Fecha, fecha_actual: Fecha) -> int:
    """Años completos de servicio entre ``fecha_ingreso`` y ``fecha_actual``."""
    ing, act = _a_fecha(fecha_ingreso), _a_fecha(fecha_actual)
    if act < ing:
        raise ValueError("La fecha de corte es anterior a la fecha de ingreso.")
    anios = act.year - ing.year
    if (act.month, act.day) < (ing.month, ing.day):
        anios -= 1
    return anios


# --------------------------------------------------------------------------- #
# 1. Bono de antigüedad
# --------------------------------------------------------------------------- #
def _normalizar_tabla_antiguedad(tabla: Iterable[Mapping[str, Any]]
                                 ) -> list[tuple[int, Decimal]]:
    """Devuelve [(años_mínimo, porcentaje)] ordenado. Acepta las columnas de la
    hoja Excel ('Años Minimo', 'Porcentaje') o las claves 'min' y 'pct'."""
    filas = []
    for fila in tabla:
        minimo = fila.get("Años Minimo", fila.get("min"))
        pct = fila.get("Porcentaje", fila.get("pct"))
        if minimo is None or pct is None:
            raise ValueError(f"Fila inválida en la tabla de antigüedad: {fila}")
        filas.append((int(minimo), D(pct)))
    if not filas:
        raise ValueError("La tabla de antigüedad está vacía.")
    return sorted(filas)


def porcentaje_antiguedad(anios: int, tabla_antiguedad: Iterable[Mapping[str, Any]]
                          ) -> Decimal:
    """Porcentaje del tramo al que pertenecen los ``anios`` cumplidos.

    Se elige el tramo con el mayor 'Años Mínimo' que no supere los años
    cumplidos. Así el solapamiento del Excel (0-2 y 2-4) se resuelve solo:
    1 año -> 0 %, 2 años -> 5 %.
    """
    pct = Decimal("0")
    for minimo, p in _normalizar_tabla_antiguedad(tabla_antiguedad):
        if anios >= minimo:
            pct = p
    return pct


def calcular_bono_antiguedad(fecha_ingreso: Fecha, fecha_actual: Fecha,
                             smn: Numero,
                             tabla_antiguedad: Iterable[Mapping[str, Any]]
                             ) -> dict[str, Any]:
    """Bono de antigüedad = porcentaje del tramo x (3 x SMN).

    Retorna ``{"anios": int, "porcentaje": Decimal, "bono": Decimal}``.
    """
    anios = anios_cumplidos(fecha_ingreso, fecha_actual)
    pct = porcentaje_antiguedad(anios, tabla_antiguedad)
    bono = redondear(3 * D(smn) * pct)
    return {"anios": anios, "porcentaje": pct, "bono": bono}


# --------------------------------------------------------------------------- #
# 2. Componentes variables y total ganado
# --------------------------------------------------------------------------- #
def calcular_componentes_variables(haber_basico: Numero,
                                   horas_extra: Numero,
                                   horas_nocturnas: Numero,
                                   dias_domingo: Numero,
                                   dias_feriado: Numero,
                                   factor_he: Numero,
                                   pct_nocturno: Numero,
                                   factor_domingo: Numero,
                                   factor_feriado: Numero,
                                   bono_antiguedad: Numero = 0,
                                   horas_mes: Numero = 240,
                                   descontar_dia_base: bool = False
                                   ) -> dict[str, Decimal]:
    """Calcula los recargos y el Total Ganado.

    * Valor hora        = haber básico / horas_mes (240)
    * Horas extra       = valor hora x horas extra x factor_he
    * Nocturno          = valor hora x horas nocturnas x pct_nocturno
    * Domingos          = (haber / 30 / 8) x (días x 8) x factor_domingo
    * Feriados          = (haber / 30 / 8) x (días x 8) x factor_feriado

    ``descontar_dia_base``: el haber mensual ya paga el día de descanso. Si es
    True, a domingos/feriados se les aplica (factor - 1), es decir, solo el
    excedente. Por defecto False (se suma el monto completo).
    """
    haber = _no_negativo(D(haber_basico), "haber_basico")
    h_extra = _no_negativo(D(horas_extra), "horas_extra")
    h_noct = _no_negativo(D(horas_nocturnas), "horas_nocturnas")
    d_dom = _no_negativo(D(dias_domingo), "dias_domingo")
    d_fer = _no_negativo(D(dias_feriado), "dias_feriado")
    h_mes = D(horas_mes)
    if h_mes <= 0:
        raise ValueError("'horas_mes' debe ser mayor que cero.")

    valor_hora = haber / h_mes                      # sin redondear (precisión interna)
    valor_hora_dia = (haber / 30) / 8               # equivale a haber / 240

    f_dom, f_fer = D(factor_domingo), D(factor_feriado)
    if descontar_dia_base:
        f_dom, f_fer = f_dom - 1, f_fer - 1

    pago_he = redondear(valor_hora * h_extra * D(factor_he))
    pago_noct = redondear(valor_hora * h_noct * D(pct_nocturno))
    pago_dom = redondear(valor_hora_dia * (d_dom * 8) * f_dom)
    pago_fer = redondear(valor_hora_dia * (d_fer * 8) * f_fer)
    bono = redondear(bono_antiguedad)

    total = redondear(haber + bono + pago_he + pago_noct + pago_dom + pago_fer)
    return {
        "valor_hora": redondear(valor_hora),
        "haber_basico": redondear(haber),
        "bono_antiguedad": bono,
        "pago_horas_extra": pago_he,
        "pago_nocturno": pago_noct,
        "pago_domingos": pago_dom,
        "pago_feriados": pago_fer,
        "total_ganado": total,
    }


# --------------------------------------------------------------------------- #
# 3. Aportes laborales y RC-IVA
# --------------------------------------------------------------------------- #
def calcular_ans(total_ganado: Numero, parametros: Mapping[str, Any]) -> Decimal:
    """Aporte Nacional Solidario (a cargo del trabajador).

    Se aplica cada tasa sobre el excedente del Total Ganado respecto de su
    umbral (acumulativo):
        P1 x (TG - T1) + P2 x (TG - T2) + P3 x (TG - T3), cada término >= 0.
    """
    tg = D(total_ganado)
    ans = Decimal("0")
    for t, p in (("ANS_T1", "ANS_P1"), ("ANS_T2", "ANS_P2"), ("ANS_T3", "ANS_P3")):
        exceso = tg - D(parametros[t])
        if exceso > 0:
            ans += exceso * D(parametros[p])
    return redondear(ans)


def calcular_aportes_y_rciva(total_ganado: Numero,
                             smn: Numero,
                             facturas_f110: Numero,
                             parametros_afp: Mapping[str, Any],
                             saldo_favor_anterior: Numero = 0,
                             ufv_anterior: Numero = 1,
                             ufv_actual: Numero = 1,
                             incluir_ans: bool = True) -> dict[str, Decimal]:
    """Aportes laborales (Gestora Pública + ANS), RC-IVA y líquido pagable.

    ``parametros_afp`` es el diccionario completo de parámetros (usa AL_CP,
    AL_RC, AL_CG, AL_AS, T_RCIVA, SMN_EXENTO, SMN_CREDITO_DEP y ANS_*).

    Pasos:
      1. Aportes AFP = TG x (AL_CP + AL_RC + AL_CG + AL_AS)  [+ ANS]
      2. Base imponible = TG - aportes laborales
      3. Impuesto determinado = T_RCIVA x max(base - SMN_EXENTO x SMN, 0)
      4. Crédito dependiente = T_RCIVA x SMN_CREDITO_DEP x SMN
      5. Crédito F-110 = T_RCIVA x facturas
      6. Saldo a favor anterior actualizado por UFV = saldo x (UFV_actual / UFV_anterior),
         sin reducirse si la UFV bajó.
      7. Neto = impuesto - (créditos + saldo anterior actualizado)
         * Neto > 0  -> RC-IVA a retener
         * Neto < 0  -> saldo a favor del dependiente (pasa al mes siguiente)
      8. Líquido pagable = TG - aportes laborales - RC-IVA a retener
    """
    p = parametros_afp
    tg = _no_negativo(D(total_ganado), "total_ganado")
    smn_d = D(smn)
    tasa = D(p["T_RCIVA"])

    # --- Aportes laborales -------------------------------------------------
    pct_afp = D(p["AL_CP"]) + D(p["AL_RC"]) + D(p["AL_CG"]) + D(p["AL_AS"])
    aporte_cp = redondear(tg * D(p["AL_CP"]))
    aporte_rc = redondear(tg * D(p["AL_RC"]))
    aporte_cg = redondear(tg * D(p["AL_CG"]))
    aporte_as = redondear(tg * D(p["AL_AS"]))
    aportes_afp = aporte_cp + aporte_rc + aporte_cg + aporte_as
    ans = calcular_ans(tg, p) if incluir_ans else Decimal("0.00")
    aportes_laborales = aportes_afp + ans

    # --- RC-IVA --------------------------------------------------------------
    base_imponible = redondear(tg - aportes_laborales)
    minimo_no_imponible = redondear(D(p["SMN_EXENTO"]) * smn_d)
    exceso = max(base_imponible - minimo_no_imponible, Decimal("0"))
    impuesto_determinado = redondear(exceso * tasa)

    credito_dependiente = redondear(tasa * D(p["SMN_CREDITO_DEP"]) * smn_d)
    credito_f110 = redondear(tasa * _no_negativo(D(facturas_f110), "facturas_f110"))

    saldo_ant = _no_negativo(D(saldo_favor_anterior), "saldo_favor_anterior")
    ufv_ini, ufv_fin = D(ufv_anterior), D(ufv_actual)
    if ufv_ini <= 0 or ufv_fin <= 0:
        raise ValueError("Los valores de UFV deben ser mayores que cero.")
    factor_ufv = max(ufv_fin / ufv_ini, Decimal("1"))
    saldo_ant_actualizado = redondear(saldo_ant * factor_ufv)

    total_creditos = credito_dependiente + credito_f110 + saldo_ant_actualizado
    neto = impuesto_determinado - total_creditos
    rciva_retener = redondear(max(neto, Decimal("0")))
    saldo_favor_nuevo = redondear(max(-neto, Decimal("0")))

    liquido = redondear(tg - aportes_laborales - rciva_retener)

    return {
        "aporte_cuenta_personal": aporte_cp,
        "aporte_riesgo_comun": aporte_rc,
        "aporte_comision": aporte_cg,
        "aporte_solidario_laboral": aporte_as,
        "aportes_afp": aportes_afp,
        "aporte_nacional_solidario": ans,
        "aportes_laborales": aportes_laborales,
        "base_imponible": base_imponible,
        "minimo_no_imponible": minimo_no_imponible,
        "exceso_imponible": redondear(exceso),
        "impuesto_determinado": impuesto_determinado,
        "credito_dependiente": credito_dependiente,
        "credito_f110": credito_f110,
        "saldo_favor_anterior_actualizado": saldo_ant_actualizado,
        "rciva_neto": redondear(neto),
        "rciva_retener": rciva_retener,
        "saldo_favor_nuevo": saldo_favor_nuevo,
        "liquido_pagable": liquido,
    }


# --------------------------------------------------------------------------- #
# 4. Aportes patronales y previsiones
# --------------------------------------------------------------------------- #
def calcular_aportes_patronales(total_ganado: Numero,
                                parametros: Mapping[str, Any]) -> dict[str, Decimal]:
    """CNS, solidario patronal, prima de riesgo profesional y pro-vivienda,
    todos sobre el Total Ganado."""
    tg = D(total_ganado)
    cns = redondear(tg * D(parametros["AP_CNS"]))
    sol = redondear(tg * D(parametros["AP_SOL"]))
    rp = redondear(tg * D(parametros["AP_RP"]))
    pv = redondear(tg * D(parametros["AP_PV"]))
    return {
        "patronal_cns": cns,
        "patronal_solidario": sol,
        "patronal_riesgo_profesional": rp,
        "patronal_provivienda": pv,
        "aportes_patronales": cns + sol + rp + pv,
    }


def calcular_previsiones(total_ganado: Numero,
                         parametros: Mapping[str, Any]) -> dict[str, Decimal]:
    """Cuota mensual de previsión para aguinaldo e indemnización."""
    tg = D(total_ganado)
    return {
        "aguinaldo": redondear(tg * D(parametros["AGUINALDO"])),
        "indemnizacion": redondear(tg * D(parametros["INDEMNIZACION"])),
    }


# --------------------------------------------------------------------------- #
# Orquestador: una función para toda la planilla de un empleado
# --------------------------------------------------------------------------- #
@dataclass
class ResultadoPlanilla:
    """Resultado completo de un empleado en un periodo."""
    anios_antiguedad: int
    pct_antiguedad: Decimal
    detalle: dict[str, Decimal] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        d = {"anios_antiguedad": self.anios_antiguedad,
             "pct_antiguedad": self.pct_antiguedad}
        d.update(self.detalle)
        return d

    def a_fila_historico(self, id_planilla: str, periodo: str, id_empleado: str,
                         fecha_registro: Optional[Fecha] = None) -> dict[str, Any]:
        """Fila con las columnas de la hoja ``Historico`` (recargos de
        domingo y feriado van en columnas separadas)."""
        d = self.detalle
        return {
            "ID_Planilla": id_planilla,
            "Periodo": periodo,
            "ID_Empleado": id_empleado,
            "Haber_Basico": d["haber_basico"],
            "Recargo_Horas_Extra": d["pago_horas_extra"],
            "Recargo_Horas_Nocturnas": d["pago_nocturno"],
            "Recargo_Dias_Domingo": d["pago_domingos"],
            "Recargo_Dias_Feriado": d["pago_feriados"],
            "Bono_Antiguedad": d["bono_antiguedad"],
            "Total_Ganado": d["total_ganado"],
            "Aportes_Laborales": d["aportes_laborales"],
            "RCIVA": d["rciva_retener"],
            "Liquido_Pagable": d["liquido_pagable"],
            "Aportes_Patronales": d["aportes_patronales"],
            "Aguinaldo": d["aguinaldo"],
            "Indemnizacion": d["indemnizacion"],
            "Fecha de Registro": fecha_registro or datetime.now(),
        }


def calcular_planilla_empleado(haber_basico: Numero,
                               fecha_ingreso: Fecha,
                               fecha_corte: Fecha,
                               parametros: Mapping[str, Any],
                               tabla_antiguedad: Iterable[Mapping[str, Any]],
                               horas_extra: Numero = 0,
                               horas_nocturnas: Numero = 0,
                               dias_domingo: Numero = 0,
                               dias_feriado: Numero = 0,
                               facturas_f110: Numero = 0,
                               saldo_favor_anterior: Numero = 0,
                               ufv_anterior: Numero = 1,
                               ufv_actual: Numero = 1,
                               descontar_dia_base: bool = False
                               ) -> ResultadoPlanilla:
    """Calcula todo el flujo de un empleado y devuelve ``ResultadoPlanilla``.

    ``parametros`` debe venir de ``cargar_parametros``. ``fecha_corte`` es el
    último día del periodo (ver ``ultimo_dia_periodo``).
    """
    p = parametros
    antig = calcular_bono_antiguedad(fecha_ingreso, fecha_corte, p["SMN"],
                                     tabla_antiguedad)
    ganado = calcular_componentes_variables(
        haber_basico, horas_extra, horas_nocturnas, dias_domingo, dias_feriado,
        factor_he=p["R_HE"], pct_nocturno=p["R_HN"],
        factor_domingo=p["R_DD"], factor_feriado=p["R_DF"],
        bono_antiguedad=antig["bono"], horas_mes=p["H_MES"],
        descontar_dia_base=descontar_dia_base)
    tg = ganado["total_ganado"]
    impuestos = calcular_aportes_y_rciva(
        tg, p["SMN"], facturas_f110, p, saldo_favor_anterior,
        ufv_anterior, ufv_actual)

    detalle: dict[str, Decimal] = {}
    detalle.update(ganado)
    detalle.update(impuestos)
    detalle.update(calcular_aportes_patronales(tg, p))
    detalle.update(calcular_previsiones(tg, p))
    return ResultadoPlanilla(antig["anios"], antig["porcentaje"], detalle)
