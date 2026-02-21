"""
Generador: 05_Analisis_Nomina_XML_Pivot.xlsx
Modulo 2 -- Procesamiento Masivo y Analisis con Tablas Dinamicas

Hojas:
  - XML_Nomina: 500+ filas simulando datos de nomina extraidos de XML del SAT
    (20 empleados x 12 meses x ~4-5 conceptos por empleado/mes)
  - Instrucciones: guia para crear 3 tablas dinamicas
"""
import sys
from pathlib import Path
import random
import uuid
from datetime import date, timedelta
from typing import List, Tuple, Dict, Any

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK, SALARIO_MINIMO_MENSUAL
from scripts.config.isr_2026 import calcular_isr_mensual, SUBSIDIO_EMPLEO_MENSUAL
from scripts.config.styles import (
    FILL_AMARILLO, FILL_VERDE, FILL_LIGHT,
    FONT_NORMAL, FONT_SMALL,
    THIN_BORDER, ALIGN_CENTER, FMT_MONEY, FMT_DATE,
    style_title_cell, auto_width
)
from scripts.generators.xlsx_gen import ExcelGenerator

OUTPUT_DIR = PACK / "Modulo_2_Tablas_Dinamicas"

# ── 20 empleados con nombres realistas mexicanos ──────────────────
EMPLEADOS = [
    (1001, "Juan Carlos Hernandez Lopez", "Contador General"),
    (1002, "Maria Guadalupe Martinez Ramirez", "Auxiliar Contable"),
    (1003, "Jose Luis Garcia Flores", "Gerente de Finanzas"),
    (1004, "Ana Patricia Rodriguez Sanchez", "Analista de Nominas"),
    (1005, "Carlos Alberto Perez Torres", "Director Administrativo"),
    (1006, "Laura Elena Gonzalez Diaz", "Asistente de Direccion"),
    (1007, "Miguel Angel Lopez Morales", "Auditor Interno"),
    (1008, "Rosa Isela Fernandez Cruz", "Recepcionista"),
    (1009, "Francisco Javier Ramirez Ortiz", "Gerente de Operaciones"),
    (1010, "Patricia Alejandra Sanchez Vega", "Auxiliar Administrativo"),
    (1011, "Roberto Carlos Diaz Jimenez", "Ejecutivo de Ventas"),
    (1012, "Gabriela Fernanda Torres Ruiz", "Coordinadora de RH"),
    (1013, "Eduardo Daniel Castro Mendoza", "Desarrollador de Sistemas"),
    (1014, "Veronica Lizeth Morales Herrera", "Diseñadora Grafica"),
    (1015, "Alejandro Ivan Ortiz Guerrero", "Chofer Repartidor"),
    (1016, "Diana Paola Jimenez Vargas", "Asistente Legal"),
    (1017, "Fernando Arturo Ruiz Aguilar", "Jefe de Almacen"),
    (1018, "Karla Ivonne Mendoza Navarro", "Ejecutiva de Cobranza"),
    (1019, "Ricardo Enrique Herrera Rios", "Supervisor de Produccion"),
    (1020, "Claudia Berenice Guerrero Solis", "Contralora"),
]

# Salary ranges monthly (gross). Min wage ~ $8,364/month
SALARY_RANGES = {
    "Recepcionista": (8400, 12000),
    "Chofer Repartidor": (8400, 13000),
    "Auxiliar Contable": (10000, 16000),
    "Auxiliar Administrativo": (9500, 15000),
    "Asistente de Direccion": (12000, 18000),
    "Asistente Legal": (12000, 18000),
    "Ejecutiva de Cobranza": (11000, 17000),
    "Ejecutivo de Ventas": (14000, 22000),
    "Analista de Nominas": (15000, 22000),
    "Diseñadora Grafica": (14000, 20000),
    "Coordinadora de RH": (18000, 28000),
    "Jefe de Almacen": (16000, 24000),
    "Desarrollador de Sistemas": (20000, 35000),
    "Contador General": (18000, 30000),
    "Auditor Interno": (22000, 35000),
    "Supervisor de Produccion": (18000, 28000),
    "Gerente de Operaciones": (30000, 45000),
    "Gerente de Finanzas": (32000, 48000),
    "Director Administrativo": (38000, 55000),
    "Contralora": (35000, 50000),
}

MESES = [
    (1, "Enero"), (2, "Febrero"), (3, "Marzo"), (4, "Abril"),
    (5, "Mayo"), (6, "Junio"), (7, "Julio"), (8, "Agosto"),
    (9, "Septiembre"), (10, "Octubre"), (11, "Noviembre"), (12, "Diciembre"),
]

PERIODOS = {
    1: "01/01/2025-15/01/2025", 2: "01/02/2025-15/02/2025",
    3: "01/03/2025-15/03/2025", 4: "01/04/2025-15/04/2025",
    5: "01/05/2025-15/05/2025", 6: "01/06/2025-15/06/2025",
    7: "01/07/2025-15/07/2025", 8: "01/08/2025-15/08/2025",
    9: "01/09/2025-15/09/2025", 10: "01/10/2025-15/10/2025",
    11: "01/11/2025-15/11/2025", 12: "01/12/2025-15/12/2025",
}


def _get_subsidio(base_mensual):
    # type: (float) -> float
    """Busca subsidio al empleo mensual segun la tabla."""
    for rango in SUBSIDIO_EMPLEO_MENSUAL:
        if rango["desde"] <= base_mensual <= rango["hasta"]:
            return rango["subsidio"]
    return 0.0


def build():
    gen = ExcelGenerator("05_Analisis_Nomina_XML_Pivot.xlsx", OUTPUT_DIR)

    HEADERS = [
        "UUID", "NumEmpleado", "NombreEmpleado", "Puesto",
        "FechaPago", "Periodo", "Clase",
        "Concepto", "ImporteGravado", "ImporteExento"
    ]

    rng = random.Random(123)
    rows = []  # type: List[List]

    for num_emp, nombre, puesto in EMPLEADOS:
        sal_min, sal_max = SALARY_RANGES.get(puesto, (10000, 20000))
        sueldo_mensual = round(rng.uniform(sal_min, sal_max), 2)

        for mes_num, mes_nombre in MESES:
            fecha_pago = date(2025, mes_num, 15)
            periodo = PERIODOS[mes_num]

            # Pequena variacion mensual (+/- 2%)
            sueldo_mes = round(sueldo_mensual * rng.uniform(0.98, 1.02), 2)

            # Porcion gravada y exenta del sueldo
            # Todo el sueldo es gravado
            sueldo_gravado = sueldo_mes
            sueldo_exento = 0.0

            # 1) Percepcion: Sueldo
            rows.append([
                str(uuid.uuid5(uuid.NAMESPACE_DNS,
                               "{}-{}-sueldo".format(num_emp, mes_num))),
                num_emp, nombre, puesto,
                fecha_pago, periodo,
                "Percepcion", "001 - Sueldos, Salarios y Asimilados",
                round(sueldo_gravado, 2), round(sueldo_exento, 2)
            ])

            # 2) Percepcion: Vacaciones (marzo, julio, diciembre, o aleatoriamente ~25%)
            if mes_num in (3, 7, 12) or rng.random() < 0.08:
                dias_vac = rng.choice([6, 8, 10, 12, 14])
                sueldo_diario = sueldo_mes / 30.0
                prima_vac = round(sueldo_diario * dias_vac * 0.25, 2)
                # Prima vacacional: primeros 15 UMA diarios exentos
                uma_diario = 113.14  # UMA 2025 estimado
                exento_pv = min(prima_vac, round(uma_diario * 15, 2))
                gravado_pv = max(0, round(prima_vac - exento_pv, 2))

                rows.append([
                    str(uuid.uuid5(uuid.NAMESPACE_DNS,
                                   "{}-{}-vac".format(num_emp, mes_num))),
                    num_emp, nombre, puesto,
                    fecha_pago, periodo,
                    "Percepcion", "021 - Prima Vacacional",
                    gravado_pv, exento_pv
                ])

            # 3) Deduccion: ISR retenido
            isr_result = calcular_isr_mensual(sueldo_gravado)
            isr_mensual = isr_result.get("isr_total", 0.0) if isr_result else 0.0

            # Subsidio reduce el ISR a retener
            subsidio = _get_subsidio(sueldo_gravado)
            isr_a_retener = max(0, round(isr_mensual - subsidio, 2))

            rows.append([
                str(uuid.uuid5(uuid.NAMESPACE_DNS,
                               "{}-{}-isr".format(num_emp, mes_num))),
                num_emp, nombre, puesto,
                fecha_pago, periodo,
                "Deduccion", "002 - ISR",
                round(isr_a_retener, 2), 0.0
            ])

            # 4) Deduccion: IMSS (cuota obrera ~ 2.775% del SBC)
            sbc = sueldo_mes  # simplification: SBC ~ sueldo mensual
            imss_obrera = round(sbc * 0.02775, 2)

            rows.append([
                str(uuid.uuid5(uuid.NAMESPACE_DNS,
                               "{}-{}-imss".format(num_emp, mes_num))),
                num_emp, nombre, puesto,
                fecha_pago, periodo,
                "Deduccion", "001 - Seguridad Social (IMSS Obrero)",
                round(imss_obrera, 2), 0.0
            ])

            # 5) OtroPago: Subsidio al empleo (cuando aplica)
            if subsidio > 0:
                rows.append([
                    str(uuid.uuid5(uuid.NAMESPACE_DNS,
                                   "{}-{}-sub".format(num_emp, mes_num))),
                    num_emp, nombre, puesto,
                    fecha_pago, periodo,
                    "OtroPago", "002 - Subsidio para el Empleo",
                    0.0, round(subsidio, 2)
                ])

    # ── Hoja: XML_Nomina ────────────────────────────────────────
    ws = gen.add_sheet("XML_Nomina")
    style_title_cell(ws, 1, 1,
                     "Datos de Nomina Extraidos de XML CFDI - Ejercicio 2025", 10)
    ws.cell(row=2, column=1,
            value="Fuente simulada: XMLs de nomina timbrados en el SAT. {} registros.".format(
                len(rows))).font = FONT_SMALL

    gen.write_table(ws, HEADERS, rows, start_row=4,
                    table_name="xml_nomina",
                    money_cols=[9, 10], date_cols=[5])

    # ── Hoja: Instrucciones ─────────────────────────────────────
    gen.add_instructions_sheet([
        "Este archivo contiene {} registros de nomina simulando datos extraidos de XML CFDI del SAT.".format(
            len(rows)),
        "Los datos cubren 20 empleados, 12 meses (Ene-Dic 2025), con percepciones, deducciones y otros pagos.",
        "Los datos ya estan formateados como Tabla de Excel ('xml_nomina'). Puedes verificar en Diseno de Tabla.",
        "",
        "=== EJERCICIO 1: Tabla Dinamica de Percepciones ===",
        "1. Selecciona cualquier celda de la tabla y ve a Insertar > Tabla Dinamica.",
        "2. Filtra la columna 'Clase' para mostrar solo 'Percepcion'.",
        "3. Filas: NombreEmpleado. Columnas: Concepto. Valores: Suma de ImporteGravado.",
        "4. Agrega un Slicer por 'Periodo' para filtrar por mes.",
        "",
        "=== EJERCICIO 2: Tabla Dinamica de Deducciones ===",
        "1. Crea otra tabla dinamica en una hoja nueva.",
        "2. Filtra 'Clase' = 'Deduccion'.",
        "3. Filas: NombreEmpleado + Puesto. Columnas: Concepto. Valores: Suma de ImporteGravado.",
        "4. Pregunta: Quien paga mas ISR? Quien tiene el sueldo mas alto?",
        "",
        "=== EJERCICIO 3: Tabla Dinamica de Costo Total por Empleado ===",
        "1. Crea una tercera tabla dinamica.",
        "2. Filas: NombreEmpleado. Columnas: Clase. Valores: Suma de ImporteGravado + ImporteExento.",
        "3. Agrega un campo calculado o columna para 'Costo Total = Percepciones - Deducciones + OtrosPagos'.",
        "4. Ordena de mayor a menor para identificar los empleados con mayor costo.",
        "",
        "NOTA: Los montos de ISR y IMSS se calcularon con las tarifas 2026 del SAT.",
        "El subsidio al empleo aplica para sueldos bajos (menos de ~$7,382 mensuales).",
        "CONSEJO: Usa segmentadores (Slicers) para hacer el analisis interactivo.",
    ])

    gen.save()


if __name__ == "__main__":
    build()
