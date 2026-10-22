"""
app.py
======
Interfaz Streamlit: ingreso de datos, cálculo de la planilla y descarga de la
Planilla Laboral, la Planilla Impositiva (RC-IVA) y el asiento contable, en
Excel y PDF.

Ejecutar localmente:  streamlit run app.py
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st

import asientos
import calculos as c
import comprobantes as cmp
import documentos as docs
import planilla as pl

RUTA_EXCEL_BASE = Path(__file__).parent / "datos" / "PlanillasCPA404_BD.xlsx"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Planillas de Sueldos y RC-IVA", page_icon="🧾",
                   layout="wide")


# --------------------------------------------------------------------------- #
# Carga de datos base
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def _leer_base(contenido: bytes | None):
    """Lee el Excel subido, o el del repositorio si no se subió ninguno."""
    import io
    origen = io.BytesIO(contenido) if contenido else str(RUTA_EXCEL_BASE)
    return pl.cargar_datos_excel(origen)


def _df_empleados(empleados: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(empleados)
    columnas = ["ID_Empleado", "CI", "Apellidos", "Nombres", "Departamento",
                "Fecha_Ingreso", "Haber_Basico"]
    for col in columnas:
        if col not in df:
            df[col] = None
    df = df[columnas].copy()
    df["Departamento"] = df["Departamento"].map(
        lambda v: asientos.NOMBRE_DEPTO[asientos.normalizar_departamento(v)])
    df["Fecha_Ingreso"] = pd.to_datetime(df["Fecha_Ingreso"], errors="coerce").dt.date
    df["Haber_Basico"] = pd.to_numeric(df["Haber_Basico"], errors="coerce")
    for col in pl.COLUMNAS_VARIABLES:
        df[col] = 0.0
    return df


# --------------------------------------------------------------------------- #
# Barra lateral: datos generales
# --------------------------------------------------------------------------- #
st.sidebar.header("Datos generales")
tipo_empresa = st.sidebar.radio("Tipo de empresa", ["Empresa privada", "Empresa pública"],
                                help="Define el formato del asiento contable.")
es_publica = tipo_empresa == "Empresa pública"
empresa = st.sidebar.text_input("Nombre de la institución" if es_publica
                                else "Nombre de la empresa",
                                value="Mi Institución" if es_publica else "Mi Empresa S.R.L.")
nit = st.sidebar.text_input("NIT", value="")

hoy = date.today()
col_m, col_a = st.sidebar.columns([3, 2])
mes_nombre = col_m.selectbox("Mes", pl.MESES, index=hoy.month - 1)
anio = int(col_a.number_input("Año", min_value=2000, max_value=2100, value=hoy.year, step=1))
mes = pl.MESES.index(mes_nombre) + 1
fecha_corte = c.ultimo_dia_periodo(anio, mes)

st.sidebar.subheader("UFV (actualización del saldo a favor)")
ufv_anterior = st.sidebar.number_input("UFV al cierre del periodo anterior", min_value=0.00001,
                                       value=1.0, step=0.00001, format="%.5f")
ufv_actual = st.sidebar.number_input("UFV al cierre del periodo actual", min_value=0.00001,
                                     value=1.0, step=0.00001, format="%.5f")

archivo = st.sidebar.file_uploader(
    "Cargar otro Excel base (opcional)", type=["xlsx"],
    help="Debe tener las hojas de parámetros, antigüedad y empleados con la misma estructura.")

try:
    filas_param, tabla_antig, empleados_base = _leer_base(archivo.getvalue() if archivo else None)
except Exception as exc:
    st.error(f"No se pudo leer el Excel base: {exc}")
    st.stop()

# Reinicia la tabla de empleados cuando cambia el origen de datos
origen_id = archivo.name + str(archivo.size) if archivo else "base"
if st.session_state.get("origen_id") != origen_id:
    st.session_state["origen_id"] = origen_id
    st.session_state["df_empleados"] = _df_empleados(empleados_base)
    st.session_state["editor_version"] = st.session_state.get("editor_version", 0) + 1

# --------------------------------------------------------------------------- #
# Cuerpo
# --------------------------------------------------------------------------- #
st.title("Planillas de Sueldos y Salarios y RC-IVA")
st.caption(f"{tipo_empresa}  ·  Periodo: {mes_nombre} {anio}  ·  Fecha de corte: "
           f"{fecha_corte.strftime('%d/%m/%Y')}")

tab_datos, tab_result, tab_asiento, tab_desc = st.tabs(
    ["1 · Empleados y variables del mes", "2 · Resultados", "3 · Asiento contable",
     "4 · Descargas"])

# ---- Parámetros y empleados ---------------------------------------------------
with tab_datos:
    with st.expander("Parámetros de ley (editables)"):
        df_param = pd.DataFrame(filas_param)[["Codigo_Parametro", "Concepto", "Valor", "Unidad"]]
        df_param = st.data_editor(
            df_param, hide_index=True, disabled=["Codigo_Parametro", "Concepto", "Unidad"],
            key=f"param_{origen_id}")
    with st.expander("Tabla de bono de antigüedad (editable)"):
        df_antig = pd.DataFrame(tabla_antig)
        df_antig = st.data_editor(df_antig, hide_index=True, key=f"antig_{origen_id}")

    st.subheader("Empleados")
    st.write("Indique el departamento de cada empleado (define las cuentas del asiento) e "
             "ingrese las variables del mes. Se pueden agregar o quitar filas.")
    df_editado = st.data_editor(
        st.session_state["df_empleados"],
        key=f"empleados_{origen_id}_{st.session_state['editor_version']}",
        num_rows="dynamic", hide_index=True,
        column_config={
            "ID_Empleado": st.column_config.TextColumn("ID"),
            "CI": st.column_config.TextColumn("CI"),
            "Departamento": st.column_config.SelectboxColumn(
                "Departamento", options=asientos.OPCIONES_DEPARTAMENTO, required=True,
                default=asientos.NOMBRE_DEPTO[asientos.ADMIN]),
            "Fecha_Ingreso": st.column_config.DateColumn("Fecha de ingreso", format="DD/MM/YYYY"),
            "Haber_Basico": st.column_config.NumberColumn("Haber básico (Bs)", min_value=0.0,
                                                          format="%.2f"),
            "Horas_Extra": st.column_config.NumberColumn("Horas extra", min_value=0.0, step=0.5),
            "Horas_Nocturnas": st.column_config.NumberColumn("Horas nocturnas", min_value=0.0, step=0.5),
            "Dias_Domingo": st.column_config.NumberColumn("Domingos trabajados", min_value=0.0, step=1.0),
            "Dias_Feriado": st.column_config.NumberColumn("Feriados trabajados", min_value=0.0, step=1.0),
            "Facturas_F110": st.column_config.NumberColumn("Facturas F-110 (Bs)", min_value=0.0,
                                                           format="%.2f"),
            "Saldo_Favor_Anterior": st.column_config.NumberColumn("Saldo a favor anterior (Bs)",
                                                                  min_value=0.0, format="%.2f"),
        })

# ---- Cálculo ------------------------------------------------------------------
try:
    parametros = c.cargar_parametros(df_param.to_dict("records"))
    tabla_antiguedad = df_antig.dropna(how="all").to_dict("records")
    registros = df_editado.dropna(how="all").to_dict("records")
    filas, errores = pl.calcular_todos(registros, parametros, tabla_antiguedad, anio, mes,
                                       ufv_anterior, ufv_actual)
except Exception as exc:
    st.error(f"Error en los parámetros: {exc}")
    st.stop()

# Verifica que calculos.py, planilla.py, asientos.py y documentos.py sean de la misma versión
_CLAVES_REQUERIDAS = [
    "nombre_completo", "departamento", "impuesto_neto", "saldo_favor_fisco",
    "saldo_favor_dependiente", "saldo_favor_anterior", "mantenimiento_valor",
    "saldo_utilizado", "exceso_imponible", "aportes_patronales", "total_descuentos"]
if filas:
    _faltan = [k for k in _CLAVES_REQUERIDAS if k not in filas[0]]
    if _faltan:
        st.error("Los archivos del repositorio son de versiones distintas. Falta el resultado: "
                 + ", ".join(_faltan) + ". Suba de nuevo TODOS los archivos .py del proyecto "
                 "(app.py, calculos.py, planilla.py, asientos.py, documentos.py, comprobantes.py) "
                 "y reinicie la app.")
        st.stop()

meta = {
    "empresa": empresa.strip() or "—", "nit": nit.strip(),
    "periodo": f"{mes_nombre} {anio}", "fecha_corte": fecha_corte,
    "parametros": parametros, "ufv_anterior": ufv_anterior, "ufv_actual": ufv_actual,
}

# ---- Resultados -----------------------------------------------------------------
with tab_result:
    for e in errores:
        st.warning(f"Se omitió {e}")
    if not filas:
        st.info("No hay empleados válidos para calcular.")
    else:
        suma = lambda k: sum(float(f[k]) for f in filas)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total ganado", f"Bs {suma('total_ganado'):,.2f}")
        m2.metric("Aportes laborales", f"Bs {suma('aportes_laborales'):,.2f}")
        m3.metric("RC-IVA retenido", f"Bs {suma('rciva_retener'):,.2f}")
        m4.metric("Líquido pagable", f"Bs {suma('liquido_pagable'):,.2f}")

        def tabla(columnas: dict[str, str]) -> pd.DataFrame:
            return pd.DataFrame([{etq: (float(f[k]) if isinstance(f[k], Decimal) else f[k])
                                  for etq, k in columnas.items()} for f in filas])

        st.subheader("Planilla laboral")
        st.dataframe(tabla({
            "Empleado": "nombre", "Años": "anios_antiguedad", "Haber básico": "haber_basico",
            "Bono antig.": "bono_antiguedad", "H. extra": "pago_horas_extra",
            "Nocturno": "pago_nocturno", "Domingos": "pago_domingos", "Feriados": "pago_feriados",
            "Total ganado": "total_ganado", "Aportes AFP": "aportes_afp",
            "ANS": "aporte_nacional_solidario", "RC-IVA": "rciva_retener",
            "Líquido pagable": "liquido_pagable"}), hide_index=True)

        st.subheader("Planilla impositiva (RC-IVA)")
        st.dataframe(tabla({
            "Empleado": "nombre_completo", "Monto ingreso neto": "base_imponible",
            "Importe sujeto a impuesto": "exceso_imponible",
            "Impuesto RC-IVA": "impuesto_determinado",
            "13% de 1 SMN": "credito_dependiente", "Impuesto neto": "impuesto_neto",
            "F-110 casilla 693": "credito_f110", "Saldo a favor del Fisco": "saldo_favor_fisco",
            "Saldo a favor del dependiente": "saldo_favor_dependiente",
            "Saldo anterior actualizado": "saldo_favor_anterior_actualizado",
            "Saldo utilizado": "saldo_utilizado", "RC-IVA retenido": "rciva_retener",
            "Saldo para el mes siguiente": "saldo_favor_nuevo"}), hide_index=True)

# ---- Asiento contable -----------------------------------------------------------
asiento = None
with tab_asiento:
    if not filas:
        st.info("Primero ingrese al menos un empleado válido.")
    else:
        st.subheader("Comprobante patrimonial" if es_publica else "Asiento contable")
        a, b = st.columns(2)
        fecha_asiento = a.date_input("Fecha del asiento", value=fecha_corte)
        numero = b.text_input("N° de comprobante" if es_publica else "N° de asiento", value="")
        glosa = st.text_input(
            "Glosa",
            value=f"Registro de la planilla de sueldos y salarios correspondiente a "
                  f"{mes_nombre} {anio}.")
        meta.update({"fecha_asiento": fecha_asiento, "glosa": glosa, "numero_asiento": numero})

        deptos = sorted({asientos.NOMBRE_DEPTO[f["departamento"]] for f in filas})
        st.caption("Departamentos con personal: " + ", ".join(deptos) +
                   ". Solo se incluyen las cuentas de esos departamentos.")

        if es_publica:
            with st.expander("Datos del comprobante", expanded=True):
                x1, x2 = st.columns(2)
                areas = x1.multiselect("Tipo de comprobante", cmp.AREAS,
                                       default=["Presupuesto", "Patrimonio"])
                imputacion = x2.radio("Imputación presupuestaria", cmp.IMPUTACIONES)
                y1, y2, y3 = st.columns(3)
                operacion = y1.selectbox("Operación especial", ["—"] + cmp.OPERACIONES)
                m_rec = y2.selectbox("Momento de registro de recursos",
                                     ["—"] + cmp.MOMENTOS_RECURSOS)
                m_gas = y3.selectbox("Momento de registro de gastos", ["—"] + cmp.MOMENTOS_GASTOS,
                                     index=1 + cmp.MOMENTOS_GASTOS.index("DEVENGADO"))
                z1, z2 = st.columns(2)
                beneficiario = z1.text_input("Beneficiario", value="")
                respaldo = z2.text_input("Documento de respaldo",
                                         value="Planilla de sueldos y salarios")
                w1, w2, w3, w4 = st.columns(4)
                apertura = w1.text_input("Apertura programática", value="---/---/---")
                ff = w2.text_input("FF", value="")
                of = w3.text_input("OF", value="")
                et = w4.text_input("ET", value="")
            meta["comprobante"] = {
                "numero": numero, "areas": areas, "imputacion": imputacion,
                "operacion": None if operacion == "—" else operacion,
                "momento_recursos": None if m_rec == "—" else m_rec,
                "momento_gastos": None if m_gas == "—" else m_gas,
                "beneficiario": beneficiario, "respaldo": respaldo,
                "apertura": apertura, "ff": ff, "of": of, "et": et}
            asiento = asientos.asiento_publico(filas)
            vista = pd.DataFrame([{
                "Rubro": l["rubro"], "Partida presupuestaria": l["partida"],
                "Código": l["codigo"], "Cuenta contable": l["cuenta"],
                "Parcial": None if l["parcial"] is None else float(l["parcial"]),
                "Debe": None if l["debe"] is None else float(l["debe"]),
                "Haber": None if l["haber"] is None else float(l["haber"])}
                for l in asiento["lineas"]])
        else:
            asiento = asientos.asiento_privado(filas)
            vista = pd.DataFrame([{
                "Cuentas": " " * l["nivel"] + l["cuenta"],
                "Parcial": None if l["parcial"] is None else float(l["parcial"]),
                "Debe": None if l["debe"] is None else float(l["debe"]),
                "Haber": None if l["haber"] is None else float(l["haber"])}
                for l in asiento["lineas"]])
        st.dataframe(vista, hide_index=True)
        t1, t2 = st.columns(2)
        t1.metric("Total debe", f"Bs {float(asiento['total_debe']):,.2f}")
        t2.metric("Total haber", f"Bs {float(asiento['total_haber']):,.2f}")
        if not asiento["cuadra"]:
            st.error("El asiento no cuadra: revise los datos.")

# ---- Descargas ------------------------------------------------------------------
with tab_desc:
    if not filas:
        st.info("Primero ingrese al menos un empleado válido.")
    else:
        sufijo = f"{anio}-{mes:02d}"
        st.subheader("Planilla Laboral")
        a, b = st.columns(2)
        a.download_button("Descargar Excel (.xlsx)", docs.excel_planilla_laboral(filas, meta),
                          file_name=f"Planilla_Laboral_{sufijo}.xlsx", mime=MIME_XLSX,
                          key="dl_lab_xlsx")
        b.download_button("Descargar PDF", docs.pdf_planilla_laboral(filas, meta),
                          file_name=f"Planilla_Laboral_{sufijo}.pdf", mime="application/pdf",
                          key="dl_lab_pdf")
        st.subheader("Planilla Impositiva (RC-IVA)")
        a, b = st.columns(2)
        a.download_button("Descargar Excel (.xlsx)", docs.excel_planilla_impositiva(filas, meta),
                          file_name=f"Planilla_Impositiva_{sufijo}.xlsx", mime=MIME_XLSX,
                          key="dl_imp_xlsx")
        b.download_button("Descargar PDF", docs.pdf_planilla_impositiva(filas, meta),
                          file_name=f"Planilla_Impositiva_{sufijo}.pdf", mime="application/pdf",
                          key="dl_imp_pdf")
        st.subheader("Comprobante patrimonial" if es_publica else "Asiento contable")
        a, b = st.columns(2)
        if es_publica:
            xlsx_as = cmp.excel_comprobante_publico(asiento, meta)
            pdf_as = cmp.pdf_comprobante_publico(asiento, meta)
            nombre_as = "Comprobante_Patrimonial"
        else:
            xlsx_as = cmp.excel_asiento_privado(asiento, meta)
            pdf_as = cmp.pdf_asiento_privado(asiento, meta)
            nombre_as = "Asiento_Contable"
        a.download_button("Descargar Excel (.xlsx)", xlsx_as,
                          file_name=f"{nombre_as}_{sufijo}.xlsx", mime=MIME_XLSX,
                          key="dl_as_xlsx")
        b.download_button("Descargar PDF", pdf_as,
                          file_name=f"{nombre_as}_{sufijo}.pdf", mime="application/pdf",
                          key="dl_as_pdf")
