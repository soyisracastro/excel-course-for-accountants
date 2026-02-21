"""
Generador: 12_Dataset_Master_Copilot.xlsx
Modulo 5 -- Automatizacion Nativa con Microsoft 365 Copilot

Hojas:
  - Ventas_Gasolinera: 1200+ transacciones de combustible (3 sucursales)
  - Nomina_Empleados: 800+ registros de nomina (20 empleados x 12 meses + detalle)
  - Instrucciones: Requisitos para usar Copilot con este dataset
"""
import sys
import random
from pathlib import Path
from datetime import date, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from faker import Faker
from scripts.config.constants import PACK, COMBUSTIBLES
from scripts.config.isr_2026 import calcular_isr_mensual
from scripts.generators.xlsx_gen import ExcelGenerator

OUTPUT_DIR = PACK / "Modulo_5_Copilot_IA"

fake = Faker("es_MX")
Faker.seed(42)
random.seed(42)

# -- Catalogos ---------------------------------------------------------------
SUCURSALES = ["Centro", "Norte", "Sur"]

VENDEDORES = [
    "Carlos Martinez", "Ana Lopez", "Miguel Hernandez", "Vendedor_3",
    "Laura Garcia", "Roberto Sanchez", "Patricia Ramirez", "Jorge Torres",
]

METODOS_PAGO = ["Efectivo", "Tarjeta", "Transferencia"]
TURNOS = ["Matutino", "Vespertino", "Nocturno"]
TIPOS_COMBUSTIBLE = list(COMBUSTIBLES.keys())  # Magna, Premium, Diesel

PUESTOS = [
    "Gerente General", "Contador", "Auxiliar Contable", "Despachador",
    "Supervisor de Turno", "Cajero", "Almacenista", "Analista Administrativo",
    "Jefe de Recursos Humanos", "Auxiliar de RH",
    "Encargado de Sucursal", "Chofer de Pipa", "Vigilante",
    "Auxiliar de Limpieza", "Tecnico de Mantenimiento", "Asistente de Direccion",
    "Jefe de Compras", "Auxiliar de Compras", "Ejecutivo de Ventas", "Practicante",
]


def _gen_ventas():
    # type: () -> list
    """Genera 1200+ filas de ventas de gasolinera con patrones intencionales."""
    rows = []  # type: list
    start = date(2025, 1, 1)
    end = date(2025, 12, 31)
    delta = (end - start).days

    for _ in range(1250):
        dia = start + timedelta(days=random.randint(0, delta))
        sucursal = random.choices(
            SUCURSALES,
            weights=[50, 30, 20],  # Centro highest volume
            k=1,
        )[0]

        # Norte sells more Premium
        if sucursal == "Norte":
            tipo = random.choices(
                TIPOS_COMBUSTIBLE,
                weights=[25, 55, 20],
                k=1,
            )[0]
        else:
            tipo = random.choices(
                TIPOS_COMBUSTIBLE,
                weights=[55, 25, 20],
                k=1,
            )[0]

        vendedor = random.choice(VENDEDORES)

        litros = round(random.uniform(10, 500), 2)

        # Vendedor_3 consistently underperforms: smaller transactions
        if vendedor == "Vendedor_3":
            litros = round(random.uniform(10, 150), 2)

        precio_min = COMBUSTIBLES[tipo]["precio_min"]
        precio_max = COMBUSTIBLES[tipo]["precio_max"]
        precio = round(random.uniform(precio_min, precio_max), 2)

        venta_total = round(litros * precio, 2)
        metodo = random.choice(METODOS_PAGO)
        turno = random.choice(TURNOS)

        rows.append([
            dia, sucursal, vendedor, tipo, litros,
            precio, venta_total, metodo, turno,
        ])

    # Sort by date
    rows.sort(key=lambda r: r[0])
    return rows


def _gen_nomina():
    # type: () -> list
    """Genera 800+ filas de nomina con anomalias intencionales."""
    empleados = []  # type: list
    for i in range(20):
        nombre = fake.name()
        rfc = fake.bothify("????######???").upper()
        puesto = PUESTOS[i]
        sueldo_base = round(random.uniform(8000, 55000), 2)
        empleados.append({
            "nombre": nombre,
            "rfc": rfc,
            "puesto": puesto,
            "sueldo_base": sueldo_base,
        })

    meses = [
        "Ene-2025", "Feb-2025", "Mar-2025", "Abr-2025",
        "May-2025", "Jun-2025", "Jul-2025", "Ago-2025",
        "Sep-2025", "Oct-2025", "Nov-2025", "Dic-2025",
    ]

    rows = []  # type: list

    for idx, emp in enumerate(empleados):
        for m_idx, periodo in enumerate(meses):
            # Anomaly: employee 5 missing months 4-6
            if idx == 5 and m_idx in (3, 4, 5):
                continue

            sueldo = emp["sueldo_base"]

            # Anomaly: employee 2 gets sudden salary jump in month 7
            if idx == 2 and m_idx >= 6:
                sueldo = round(sueldo * 1.45, 2)

            # Anomaly: employee 10 gets sudden salary jump in month 10
            if idx == 10 and m_idx >= 9:
                sueldo = round(sueldo * 1.35, 2)

            # Overtime: spikes in December for everyone
            if m_idx == 11:
                horas_extra_val = round(random.uniform(10, 40) * 80, 2)
            else:
                horas_extra_val = round(random.uniform(0, 10) * 80, 2)

            bono = round(random.uniform(0, sueldo * 0.10), 2)
            total_percep = round(sueldo + horas_extra_val + bono, 2)

            # ISR calculation
            isr_data = calcular_isr_mensual(total_percep)
            isr = isr_data.get("isr_total", 0.0) if isr_data else 0.0

            imss = round(total_percep * 0.0275, 2)
            otras_ded = round(random.uniform(0, 500), 2)
            total_ded = round(isr + imss + otras_ded, 2)
            neto = round(total_percep - total_ded, 2)

            rows.append([
                emp["nombre"], emp["rfc"], emp["puesto"], periodo,
                sueldo, horas_extra_val, bono, total_percep,
                isr, imss, otras_ded, total_ded, neto,
            ])

    # Add extra detail rows to reach 800+
    # Duplicate some employees with "Aguinaldo" and "PTU" concepts
    for idx, emp in enumerate(empleados):
        sueldo = emp["sueldo_base"]
        # Aguinaldo (December)
        aguinaldo = round(sueldo * 15 / 30, 2)  # 15 days
        isr_data = calcular_isr_mensual(aguinaldo)
        isr = isr_data.get("isr_total", 0.0) if isr_data else 0.0
        imss = round(aguinaldo * 0.0275, 2)
        total_ded = round(isr + imss, 2)
        rows.append([
            emp["nombre"], emp["rfc"], emp["puesto"], "Aguinaldo-2025",
            aguinaldo, 0.0, 0.0, aguinaldo,
            isr, imss, 0.0, total_ded, round(aguinaldo - total_ded, 2),
        ])
        # PTU (May)
        ptu = round(random.uniform(2000, 15000), 2)
        isr_data_p = calcular_isr_mensual(ptu)
        isr_p = isr_data_p.get("isr_total", 0.0) if isr_data_p else 0.0
        imss_p = 0.0
        total_ded_p = round(isr_p, 2)
        rows.append([
            emp["nombre"], emp["rfc"], emp["puesto"], "PTU-2025",
            0.0, 0.0, ptu, ptu,
            isr_p, imss_p, 0.0, total_ded_p, round(ptu - total_ded_p, 2),
        ])

    # Add weekly detail rows for 10 employees (4 weeks each) to exceed 800
    for idx in range(10):
        emp = empleados[idx]
        sueldo_semanal = round(emp["sueldo_base"] / 4, 2)
        for sem in range(1, 5):
            periodo_s = "Sem{:02d}-Dic-2025".format(sem)
            horas = round(random.uniform(2, 8) * 80, 2)
            bono = round(random.uniform(0, sueldo_semanal * 0.05), 2)
            total_p = round(sueldo_semanal + horas + bono, 2)
            isr_data_s = calcular_isr_mensual(total_p)
            isr_s = isr_data_s.get("isr_total", 0.0) if isr_data_s else 0.0
            imss_s = round(total_p * 0.0275, 2)
            otras_s = round(random.uniform(0, 200), 2)
            total_d = round(isr_s + imss_s + otras_s, 2)
            rows.append([
                emp["nombre"], emp["rfc"], emp["puesto"], periodo_s,
                sueldo_semanal, horas, bono, total_p,
                isr_s, imss_s, otras_s, total_d, round(total_p - total_d, 2),
            ])

    return rows


def build():
    gen = ExcelGenerator("12_Dataset_Master_Copilot.xlsx", OUTPUT_DIR)

    # -- Hoja 1: Ventas_Gasolinera -------------------------------------------
    ws1 = gen.add_sheet("Ventas_Gasolinera")
    headers_ventas = [
        "Fecha", "Sucursal", "Vendedor", "TipoCombustible", "Litros",
        "PrecioUnitario", "Venta_Total", "MetodoPago", "Turno",
    ]
    data_ventas = _gen_ventas()
    gen.write_table(
        ws1, headers_ventas, data_ventas,
        table_name="Ventas_Gasolinera",
        money_cols=[6, 7], date_cols=[1],
    )

    # -- Hoja 2: Nomina_Empleados --------------------------------------------
    ws2 = gen.add_sheet("Nomina_Empleados")
    headers_nomina = [
        "Empleado", "RFC", "Puesto", "Periodo", "SueldoBase",
        "HorasExtra", "BonoProductividad", "TotalPercepcion",
        "ISR", "IMSS", "OtrasDeducciones", "TotalDeduccion", "NetoPagar",
    ]
    data_nomina = _gen_nomina()
    gen.write_table(
        ws2, headers_nomina, data_nomina,
        table_name="Nomina_Empleados",
        money_cols=[5, 6, 7, 8, 9, 10, 11, 12, 13],
    )

    # -- Hoja 3: Instrucciones -----------------------------------------------
    gen.add_instructions_sheet([
        "Este archivo contiene dos tablas de datos disenadas para practicar con Microsoft 365 Copilot.",
        "REQUISITOS para usar Copilot en Excel:",
        "   a) Licencia Microsoft 365 con Copilot habilitado (Business/Enterprise + add-on Copilot).",
        "   b) El archivo DEBE estar guardado en OneDrive o SharePoint (no funciona en escritorio local).",
        "   c) Los datos DEBEN estar en formato de Tabla con nombre (ya estan configurados).",
        "Tabla 1 - Ventas_Gasolinera: 1,250 transacciones de 3 sucursales durante 2025.",
        "   Patrones ocultos: Norte vende mas Premium, Centro tiene mayor volumen, Vendedor_3 tiene bajo rendimiento.",
        "Tabla 2 - Nomina_Empleados: 800+ registros de nomina con 20 empleados.",
        "   Anomalias intencionales: 2 empleados con incrementos subitos de sueldo, 1 con meses faltantes, horas extra disparadas en diciembre.",
        "COMO USAR: Sube este archivo a OneDrive, abrelo en Excel (web o desktop), y abre el panel de Copilot.",
        "Prueba los prompts de la Guia_Prompts_Copilot_Contadores.pdf incluida en este modulo.",
        "IMPORTANTE: Copilot puede cometer errores. Siempre valida sus respuestas con tu criterio profesional.",
    ], sheet_name="Instrucciones")

    gen.save()


if __name__ == "__main__":
    build()
