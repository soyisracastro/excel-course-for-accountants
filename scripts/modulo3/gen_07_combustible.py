"""
Generador: 07_Dashboard_Ventas_Combustible.xlsx
Modulo 3 -- Visualizacion de Impacto y Reportes Ejecutivos

Hojas:
  - Datos: Ventas mensuales de combustible Ene-Dic 2025 (Magna, Premium, Diesel)
  - Grafico: Columnas apiladas con colores corporativos
  - Instrucciones

Usa xlsxwriter para soporte nativo de graficos.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import random
import xlsxwriter
from scripts.config.constants import PACK, COMBUSTIBLES, CURSO_NOMBRE, INSTRUCTOR, ANIO

OUTPUT_DIR = PACK / "Modulo_3_Visualizacion"

# Meses del anio
MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# Patrones de estacionalidad (multiplicadores sobre base)
# Enero bajo ("cuesta de enero"), verano alto, diciembre baja ligeramente
ESTACIONALIDAD = [
    0.82,  # Ene - cuesta de enero
    0.88,  # Feb
    0.93,  # Mar
    0.97,  # Abr
    1.02,  # May
    1.08,  # Jun - verano inicia
    1.12,  # Jul - pico verano
    1.10,  # Ago
    1.03,  # Sep
    0.98,  # Oct
    0.95,  # Nov
    0.92,  # Dic - baja ligera
]


def _generar_datos():
    # type: () -> list
    """Genera datos mensuales realistas de venta de combustible."""
    random.seed(42)
    datos = []

    # Bases de litros (Magna mayor, luego Diesel, luego Premium)
    # Keys match COMBUSTIBLES dict from constants (Diesel uses accent)
    bases = {
        "Magna":   120000,
        "Premium":  85000,
        "Di\u00e9sel":  100000,
    }

    for i, mes in enumerate(MESES):
        factor = ESTACIONALIDAD[i]
        row = [mes]

        for tipo in ["Magna", "Premium", "Di\u00e9sel"]:
            info = COMBUSTIBLES[tipo]
            base_litros = bases[tipo]

            # Litros con variacion aleatoria +/- 5%
            litros = int(base_litros * factor * random.uniform(0.95, 1.05))

            # Precio aleatorio dentro del rango
            precio = round(random.uniform(info["precio_min"], info["precio_max"]), 2)
            monto = round(litros * precio, 2)

            row.extend([litros, monto])

        datos.append(row)

    return datos


def build():
    # type: () -> Path
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filepath = OUTPUT_DIR / "07_Dashboard_Ventas_Combustible.xlsx"

    wb = xlsxwriter.Workbook(str(filepath))

    # ── Formatos ──────────────────────────────────────────────────
    fmt_header = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#FFFFFF",
        "bg_color": "#2563EB",
        "border": 1,
        "border_color": "#CBD5E1",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
    })

    fmt_titulo = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 14,
        "font_color": "#2563EB",
    })

    fmt_normal = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "align": "left",
        "valign": "vcenter",
    })

    fmt_litros = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "num_format": "#,##0",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_monto = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "num_format": "$#,##0.00",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_zebra = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F8FAFC",
        "align": "left",
        "valign": "vcenter",
    })

    fmt_litros_z = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F8FAFC",
        "num_format": "#,##0",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_monto_z = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F8FAFC",
        "num_format": "$#,##0.00",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_instruccion = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "text_wrap": True,
        "valign": "vcenter",
    })

    fmt_instr_num = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "align": "center",
        "valign": "vcenter",
    })

    # ── Hoja 1: Datos ─────────────────────────────────────────────
    ws_datos = wb.add_worksheet("Datos")

    ws_datos.merge_range("A1:G1", "Ventas Mensuales de Combustible 2025 - Gasolinera Ejemplo S.A. de C.V.", fmt_titulo)

    headers = [
        "Mes",
        "Magna (litros)", "Magna (monto)",
        "Premium (litros)", "Premium (monto)",
        "Di\u00e9sel (litros)", "Di\u00e9sel (monto)",
    ]

    for col, h in enumerate(headers):
        ws_datos.write(2, col, h, fmt_header)

    datos = _generar_datos()
    for r, row_data in enumerate(datos):
        row_num = 3 + r
        is_zebra = r % 2 == 0

        for c, val in enumerate(row_data):
            if c == 0:  # Mes
                ws_datos.write(row_num, c, val, fmt_zebra if is_zebra else fmt_normal)
            elif c % 2 == 1:  # Litros (cols 1, 3, 5)
                ws_datos.write(row_num, c, val, fmt_litros_z if is_zebra else fmt_litros)
            else:  # Monto (cols 2, 4, 6)
                ws_datos.write(row_num, c, val, fmt_monto_z if is_zebra else fmt_monto)

    # Totales
    fmt_total = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F1F5F9",
        "align": "right",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    fmt_total_monto = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F1F5F9",
        "align": "right",
        "valign": "vcenter",
        "num_format": "$#,##0.00",
    })

    fmt_total_label = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F1F5F9",
        "align": "left",
        "valign": "vcenter",
    })

    total_row = 3 + len(datos)
    ws_datos.write(total_row, 0, "TOTAL", fmt_total_label)
    for col in range(1, 7):
        col_letter = chr(ord("B") + col - 1)
        formula = "=SUM({0}4:{0}15)".format(col_letter)
        if col % 2 == 1:  # Litros
            ws_datos.write_formula(total_row, col, formula, fmt_total)
        else:  # Monto
            ws_datos.write_formula(total_row, col, formula, fmt_total_monto)

    # Anchos de columna
    ws_datos.set_column("A:A", 14)
    ws_datos.set_column("B:B", 16)
    ws_datos.set_column("C:C", 18)
    ws_datos.set_column("D:D", 17)
    ws_datos.set_column("E:E", 18)
    ws_datos.set_column("F:F", 16)
    ws_datos.set_column("G:G", 18)

    # ── Hoja 2: Grafico ──────────────────────────────────────────
    ws_chart = wb.add_worksheet("Grafico")

    chart = wb.add_chart({"type": "column", "subtype": "stacked"})
    chart.set_title({"name": "Ventas Mensuales por Tipo de Combustible (Monto $)"})
    chart.set_x_axis({"name": "Mes"})
    chart.set_y_axis({
        "name": "Monto ($)",
        "num_format": "$#,##0",
    })
    chart.set_size({"width": 960, "height": 520})
    chart.set_style(10)

    # Series: Montos de cada tipo
    # Magna monto = col C (index 2), Premium monto = col E (index 4), Diesel monto = col G (index 6)
    series_config = [
        {"name": "Magna",   "col": 2, "color": "#10B981"},
        {"name": "Premium", "col": 4, "color": "#EF4444"},
        {"name": "Di\u00e9sel",  "col": 6, "color": "#64748B"},
    ]

    for s in series_config:
        chart.add_series({
            "name":       s["name"],
            "categories": ["Datos", 3, 0, 14, 0],   # A4:A15 (meses)
            "values":     ["Datos", 3, s["col"], 14, s["col"]],
            "fill":       {"color": s["color"]},
            "border":     {"color": s["color"]},
            "gap":        80,
        })

    chart.set_legend({"position": "bottom"})
    ws_chart.insert_chart("B2", chart)

    # Grafico secundario: litros
    chart2 = wb.add_chart({"type": "column", "subtype": "stacked"})
    chart2.set_title({"name": "Ventas Mensuales por Tipo de Combustible (Litros)"})
    chart2.set_x_axis({"name": "Mes"})
    chart2.set_y_axis({
        "name": "Litros",
        "num_format": "#,##0",
    })
    chart2.set_size({"width": 960, "height": 520})
    chart2.set_style(10)

    series_litros = [
        {"name": "Magna",   "col": 1, "color": "#10B981"},
        {"name": "Premium", "col": 3, "color": "#EF4444"},
        {"name": "Di\u00e9sel",  "col": 5, "color": "#64748B"},
    ]

    for s in series_litros:
        chart2.add_series({
            "name":       s["name"],
            "categories": ["Datos", 3, 0, 14, 0],
            "values":     ["Datos", 3, s["col"], 14, s["col"]],
            "fill":       {"color": s["color"]},
            "border":     {"color": s["color"]},
            "gap":        80,
        })

    chart2.set_legend({"position": "bottom"})
    ws_chart.insert_chart("B32", chart2)

    # ── Hoja 3: Instrucciones ─────────────────────────────────────
    ws_instr = wb.add_worksheet("Instrucciones")
    ws_instr.merge_range("A1:H1", "Instrucciones", fmt_titulo)

    instrucciones = [
        "Este archivo contiene datos de ventas de combustible de una gasolinera ficticia para el anio 2025.",
        "La hoja 'Datos' muestra litros vendidos y monto en pesos para Magna, Premium y Diesel.",
        "La hoja 'Grafico' contiene columnas apiladas: monto por mes y litros por mes.",
        "Colores: Magna = verde (#10B981), Premium = rojo (#EF4444), Diesel = gris (#64748B).",
        "Observa la estacionalidad: enero bajo por 'cuesta de enero', verano alto, diciembre baja.",
        "Ejercicio: Crea tu propia grafica a partir de los datos. Prueba otros tipos de grafico.",
        "Consejo: Para graficos dinamicos, primero convierte los datos en Tabla (Ctrl+T).",
        "Los precios por litro reflejan promedios nacionales 2025 (Magna ~$23-25, Premium ~$25-27, Diesel ~$25-27).",
    ]

    for i, text in enumerate(instrucciones):
        row = i + 2
        ws_instr.write(row, 0, i + 1, fmt_instr_num)
        ws_instr.write(row, 1, text, fmt_instruccion)

    ws_instr.set_column("A:A", 5)
    ws_instr.set_column("B:B", 90)

    wb.close()
    print("  \u2713 {}".format(filepath.name))
    return filepath


if __name__ == "__main__":
    build()
