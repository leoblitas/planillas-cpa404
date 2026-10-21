# Planillas de Sueldos y Salarios y RC-IVA (Bolivia)

Aplicación web en Streamlit que calcula la planilla de sueldos y salarios y la
planilla impositiva (RC-IVA) y genera ambas en **Excel** y **PDF** (carta apaisado).

## Qué hace
- Calcula bono de antigüedad, horas extra, recargo nocturno, domingos y feriados.
- Aportes laborales (cuenta personal, riesgo común, comisión, solidario) y Aporte Nacional Solidario (ANS).
- RC-IVA: mínimo no imponible de 2 SMN, tasa 13 %, crédito de dependiente (13 % de 1 SMN),
  crédito por facturas F-110 y saldo a favor actualizado por UFV.
- Aportes patronales, aguinaldo e indemnización.
- No guarda datos: todo ocurre en la sesión y se descarga al terminar.

## Estructura
| Archivo | Función |
|---|---|
| `app.py` | Interfaz Streamlit |
| `calculos.py` | Reglas de cálculo (solo biblioteca estándar, aritmética `Decimal`) |
| `planilla.py` | Lee el Excel base y calcula todos los empleados |
| `documentos.py` | Genera los 4 documentos (Excel y PDF) |
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
- RC-IVA neto = 13 % × (base − 2 SMN) − (13 % × 1 SMN + 13 % × facturas F-110 + saldo anterior actualizado).
  Si es negativo, no se retiene y queda como saldo a favor para el mes siguiente.
