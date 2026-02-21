"""
Generador: 08_Comparativa_Anual_Ventas_Gastos.xlsx
Modulo 3 -- Visualizacion de Impacto y Reportes Ejecutivos

Hojas:
  - Estado_Resultados: Cifras comparativas 2024 vs 2025 (en millones)
  - Grafico_Ingresos: Barras comparativas de ingresos
  - Grafico_Gastos: Barras comparativas de gastos
  - Instrucciones

Usa xlsxwriter para soporte nativo de graficos.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import xlsxwriter
from scripts.config.constants import PACK, CURSO_NOMBRE, INSTRUCTOR, ANIO

OUTPUT_DIR = PACK / "Modulo_3_Visualizacion"

# Estado de Resultados simplificado (cifras en millones de pesos)
# 2025 muestra ~5-8% crecimiento sobre 2024
CONCEPTOS = [
    # (concepto, valor_2024, valor_2025, es_subtotal)
    ("Ventas",              85.2,   91.7,   False),
    ("Otros Ingresos",       3.8,    4.1,   False),
    ("Total Ingresos",      89.0,   95.8,   True),
    ("Compras",             42.5,   44.9,   False),
    ("Gastos Generales",    18.3,   19.2,   False),
    ("Gastos de Nomina",    15.7,   16.9,   False),
    ("Total Gastos",        76.5,   81.0,   True),
    ("Utilidad Bruta",      12.5,   14.8,   True),
]


def build():
    # type: () -> Path
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filepath = OUTPUT_DIR / "08_Comparativa_Anual_Ventas_Gastos.xlsx"

    wb = xlsxwriter.Workbook(str(filepath))

    # ── Formatos ──────────────────────────────────────────────────
    fmt_titulo = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 14,
        "font_color": "#2563EB",
    })

    fmt_subtitulo = wb.add_format({
        "font_name": "Calibri",
        "font_size": 10,
        "font_color": "#475569",
    })

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
    })

    fmt_concepto = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "align": "left",
        "valign": "vcenter",
    })

    fmt_cifra = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "num_format": "$#,##0.0",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_variacion = wb.add_format({
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "num_format": "0.0%",
        "align": "center",
        "valign": "vcenter",
    })

    fmt_subtotal_concepto = wb.add_format({
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

    fmt_subtotal_cifra = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F1F5F9",
        "num_format": "$#,##0.0",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_subtotal_var = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 11,
        "font_color": "#1E293B",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#F1F5F9",
        "num_format": "0.0%",
        "align": "center",
        "valign": "vcenter",
    })

    fmt_utilidad_concepto = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 12,
        "font_color": "#FFFFFF",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#10B981",
        "align": "left",
        "valign": "vcenter",
    })

    fmt_utilidad_cifra = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 12,
        "font_color": "#FFFFFF",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#10B981",
        "num_format": "$#,##0.0",
        "align": "right",
        "valign": "vcenter",
    })

    fmt_utilidad_var = wb.add_format({
        "bold": True,
        "font_name": "Calibri",
        "font_size": 12,
        "font_color": "#FFFFFF",
        "border": 1,
        "border_color": "#CBD5E1",
        "bg_color": "#10B981",
        "num_format": "0.0%",
        "align": "center",
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

    # ── Hoja 1: Estado de Resultados ──────────────────────────────
    ws = wb.add_worksheet("Estado_Resultados")

    ws.merge_range("A1:E1",
                   "Estado de Resultados Comparativo - Empresa Ejemplo S.A. de C.V.",
                   fmt_titulo)
    ws.write(1, 0, "Cifras en millones de pesos (MDP)", fmt_subtitulo)

    # Headers
    headers = ["Concepto", "2024", "2025", "Variacion ($)", "Variacion (%)"]
    for col, h in enumerate(headers):
        ws.write(3, col, h, fmt_header)

    for r, (concepto, v2024, v2025, es_sub) in enumerate(CONCEPTOS):
        row_num = 4 + r
        variacion_abs = v2025 - v2024
        variacion_pct = (v2025 - v2024) / v2024 if v2024 != 0 else 0

        # Ultimo rubro = Utilidad Bruta (estilo especial)
        if concepto == "Utilidad Bruta":
            ws.write(row_num, 0, concepto, fmt_utilidad_concepto)
            ws.write(row_num, 1, v2024, fmt_utilidad_cifra)
            ws.write(row_num, 2, v2025, fmt_utilidad_cifra)
            ws.write(row_num, 3, variacion_abs, fmt_utilidad_cifra)
            ws.write(row_num, 4, variacion_pct, fmt_utilidad_var)
        elif es_sub:
            ws.write(row_num, 0, concepto, fmt_subtotal_concepto)
            ws.write(row_num, 1, v2024, fmt_subtotal_cifra)
            ws.write(row_num, 2, v2025, fmt_subtotal_cifra)
            ws.write(row_num, 3, variacion_abs, fmt_subtotal_cifra)
            ws.write(row_num, 4, variacion_pct, fmt_subtotal_var)
        else:
            ws.write(row_num, 0, concepto, fmt_concepto)
            ws.write(row_num, 1, v2024, fmt_cifra)
            ws.write(row_num, 2, v2025, fmt_cifra)
            ws.write(row_num, 3, variacion_abs, fmt_cifra)
            ws.write(row_num, 4, variacion_pct, fmt_variacion)

    ws.set_column("A:A", 22)
    ws.set_column("B:C", 14)
    ws.set_column("D:D", 16)
    ws.set_column("E:E", 16)

    # ── Hoja 2: Grafico Ingresos ─────────────────────────────────
    ws_ing = wb.add_worksheet("Grafico_Ingresos")

    chart_ing = wb.add_chart({"type": "bar"})
    chart_ing.set_title({"name": "Comparativa de Ingresos 2024 vs 2025 (MDP)"})
    chart_ing.set_y_axis({"name": "Concepto"})
    chart_ing.set_x_axis({
        "name": "Millones de pesos",
        "num_format": "$#,##0.0",
    })
    chart_ing.set_size({"width": 800, "height": 420})
    chart_ing.set_style(10)

    # Series de ingresos: filas 0, 1 (Ventas, Otros Ingresos) = rows 5, 6 in sheet
    # 2024 series
    chart_ing.add_series({
        "name":       "2024",
        "categories": ["Estado_Resultados", 4, 0, 5, 0],
        "values":     ["Estado_Resultados", 4, 1, 5, 1],
        "fill":       {"color": "#94A3B8"},
        "border":     {"color": "#64748B"},
        "gap":        100,
    })
    # 2025 series
    chart_ing.add_series({
        "name":       "2025",
        "categories": ["Estado_Resultados", 4, 0, 5, 0],
        "values":     ["Estado_Resultados", 4, 2, 5, 2],
        "fill":       {"color": "#2563EB"},
        "border":     {"color": "#1D4ED8"},
        "gap":        100,
    })

    chart_ing.set_legend({"position": "bottom"})
    ws_ing.insert_chart("B2", chart_ing)

    # ── Hoja 3: Grafico Gastos ───────────────────────────────────
    ws_gas = wb.add_worksheet("Grafico_Gastos")

    chart_gas = wb.add_chart({"type": "bar"})
    chart_gas.set_title({"name": "Comparativa de Gastos 2024 vs 2025 (MDP)"})
    chart_gas.set_y_axis({"name": "Concepto"})
    chart_gas.set_x_axis({
        "name": "Millones de pesos",
        "num_format": "$#,##0.0",
    })
    chart_gas.set_size({"width": 800, "height": 420})
    chart_gas.set_style(10)

    # Gastos: Compras (row 7), Gastos Generales (row 8), Gastos Nomina (row 9)
    chart_gas.add_series({
        "name":       "2024",
        "categories": ["Estado_Resultados", 7, 0, 9, 0],
        "values":     ["Estado_Resultados", 7, 1, 9, 1],
        "fill":       {"color": "#94A3B8"},
        "border":     {"color": "#64748B"},
        "gap":        100,
    })
    chart_gas.add_series({
        "name":       "2025",
        "categories": ["Estado_Resultados", 7, 0, 9, 0],
        "values":     ["Estado_Resultados", 7, 2, 9, 2],
        "fill":       {"color": "#EF4444"},
        "border":     {"color": "#DC2626"},
        "gap":        100,
    })

    chart_gas.set_legend({"position": "bottom"})
    ws_gas.insert_chart("B2", chart_gas)

    # ── Hoja 4: Instrucciones ─────────────────────────────────────
    ws_instr = wb.add_worksheet("Instrucciones")
    ws_instr.merge_range("A1:H1", "Instrucciones", fmt_titulo)

    instrucciones = [
        "Este archivo contiene un Estado de Resultados simplificado comparativo 2024 vs 2025.",
        "Las cifras estan expresadas en millones de pesos (MDP).",
        "La hoja 'Estado_Resultados' muestra conceptos de ingreso, gasto y utilidad bruta.",
        "La columna 'Variacion ($)' muestra el cambio absoluto entre anios.",
        "La columna 'Variacion (%)' muestra el cambio porcentual (crecimiento ~5-8%).",
        "La hoja 'Grafico_Ingresos' compara Ventas y Otros Ingresos en barras horizontales.",
        "La hoja 'Grafico_Gastos' compara Compras, Gastos Generales y Nomina.",
        "Ejercicio: Agrega un grafico combinado que muestre ingresos y gastos en la misma vista.",
        "Consejo: Usa colores consistentes - gris para anio anterior, color fuerte para anio actual.",
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
