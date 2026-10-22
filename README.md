# Planillas de Sueldos y Salarios y RC-IVA (Bolivia)

Aplicación web en Streamlit que calcula la planilla de sueldos y salarios y la
planilla impositiva (RC-IVA) y genera ambas en **Excel** y **PDF** (carta apaisado).

## Qué hace
- Calcula bono de antigüedad, horas extra, recargo nocturno, domingos y feriados.
- Aportes laborales (cuenta personal, riesgo común, comisión, solidario) y Aporte Nacional Solidario (ANS).
- RC-IVA con el esquema de la planilla tributaria (mínimo no imponible de 2 SMN, tasa 13 %, crédito de
  dependiente de 13 % de 1 SMN, F-110 casilla 693, saldo a favor del Fisco y del dependiente, saldo del
  periodo anterior con mantenimiento de valor por UFV).
- Asiento contable según el tipo de empresa, elegido al inicio en la barra lateral:
  **privada** (Producción en Proceso / Gastos de Administración y Comercialización contra cuentas por pagar)
  o **pública** (comprobante patrimonial con partidas presupuestarias y cuentas contables).
  Cada empleado tiene un departamento (Producción o Administración y Comercialización); si no hay personal
  de un departamento, sus cuentas no aparecen.
- Aportes patronales, aguinaldo e indemnización.
- No guarda datos: todo ocurre en la sesión y se descarga al terminar.

## Estructura
| Archivo | Función |
|---|---|
| `app.py` | Interfaz Streamlit |
| `calculos.py` | Reglas de cálculo (solo biblioteca estándar, aritmética `Decimal`) |
| `planilla.py` | Lee el Excel base y calcula todos los empleados |
| `documentos.py` | Planilla Laboral y Planilla Impositiva (Excel y PDF) |
| `asientos.py` | Arma las líneas del asiento (privado y público) |
| `comprobantes.py` | Imprime el asiento y el comprobante patrimonial (Excel y PDF) |
| `datos/PlanillasCPA404_BD.xlsx` | Parámetros de ley, tabla de antigüedad y empleados de partida |
| `tests/test_calculos.py` | Pruebas con valores verificados a mano |

## Publicar con un solo link
1. Crea un repositorio en GitHub y sube todo el contenido de esta carpeta.
2. Entra a <https://share.streamlit.io>, inicia sesión con GitHub y elige **Create app**.
3. Selecciona el repositorio, la rama y el archivo principal `app.py`, y pulsa **Deploy**.
4. Streamlit entrega una URL pública para compartir.

## Ejecutar en local
```
pip install -r requirements.txt
streamlit run app.py
```

## Pruebas
```
python tests/test_calculos.py
```

## Parámetros
Se leen de la hoja de parámetros del Excel y se pueden editar en la app. Los que el Excel
no trae se completan desde `PARAMETROS_DEFECTO` en `calculos.py`: `SMN_CREDITO_DEP`
(crédito del dependiente, en SMN), umbrales y tasas del ANS (`ANS_T1..3`, `ANS_P1..3`)
y aportes patronales (`AP_CNS`, `AP_SOL`, `AP_RP`, `AP_PV`).
Para cambiarlos de forma permanente, agrega esas filas a la hoja de parámetros del Excel.

## Fórmulas
- Valor hora = haber básico / 240. Horas extra = valor hora × horas × factor. Nocturno = valor hora × horas × recargo.
- Domingo o feriado = (haber / 30) × días × factor.
- ANS = P1 × (TG − T1) + P2 × (TG − T2) + P3 × (TG − T3), sumando solo los excesos positivos.
- Base imponible = total ganado − aportes laborales (incluye ANS).
- Ingreso neto = total ganado − aportes laborales (incluye ANS). Importe sujeto = máx(ingreso neto − 2 SMN; 0).
- Impuesto RC-IVA = 13 % × importe sujeto. Impuesto neto = máx(impuesto − 13 % de 1 SMN; 0).
  El crédito de 1 SMN solo compensa impuesto: no genera saldo a favor.
- F-110 casilla 693 = 13 % × facturas. Saldo a favor del Fisco = máx(impuesto neto − F-110; 0);
  saldo a favor del dependiente = máx(F-110 − impuesto neto; 0).
- Saldo utilizado = mín(saldo anterior actualizado; saldo a favor del Fisco). RC-IVA retenido = saldo a favor del Fisco − saldo utilizado.
- Saldo para el mes siguiente = saldo a favor del dependiente + (saldo anterior actualizado − saldo utilizado).

## Asientos contables
**Privada:** debe = sueldos y salarios (total ganado), aportes patronales (17,21 %), aguinaldos e indemnizaciones (8,33 % c/u) por departamento.
Haber = sueldos por pagar (líquido), Gestora (aportes laborales + patronales sin CNS), C.N.S., RC-IVA dependientes por pagar, aguinaldos por pagar y previsión para indemnizaciones.

**Pública:** misma partida doble con rubros y cuentas: 11700 Sueldos, 11220 Bono de antigüedad, 11910 Horas extraordinarias
(horas extra + nocturno + domingos + feriados), 11400 Aguinaldo, 13110, 13120, 13131, 13200 y 94100 hacia 1155 (Producción) o 611, 6133, 6121, 6123, 6122 y 6131 (Administración);
haber en 2113, 21141, 21142, 21161, 2141 y 2251. Los códigos están en `asientos.py` (`PARTIDAS_PUBLICAS`, `CUENTAS_PUBLICAS_PAGAR`).
Las líneas con monto cero no se imprimen.
