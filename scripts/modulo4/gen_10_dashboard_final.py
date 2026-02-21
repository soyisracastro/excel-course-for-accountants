"""
Generador: 10_Dashboard_Final_Integrado.xlsx
Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

Hojas:
  - Datos_Nomina: 20 empleados x 12 meses con datos de nomina (tabla nombrada)
  - Tarifa_ISR: Tarifa mensual ISR 2026 completa (tabla nombrada)
  - Calculadora: Area para calculo ISR con VLOOKUP
  - Dashboard: Layout con KPIs, instrucciones para slicers y graficos
  - Instrucciones
"""
import sys
from pathlib import Path
import random

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from openpyxl.utils import get_column_letter
from scripts.config.constants import PACK, Color, SALARIO_MINIMO_MENSUAL
from scripts.config.isr_2026 import TARIFA_MENSUAL, calcular_isr_mensual
from scripts.config.styles import (
    FILL_HEADER, FILL_LIGHT, FILL_MEDIO, FILL_VERDE, FILL_AMARILLO, FILL_BLANCO,
    FONT_HEADER, FONT_TITULO_XL, FONT_SUBTITULO, FONT_NORMAL, FONT_SMALL,
    THIN_BORDER, ALIGN_CENTER, ALIGN_LEFT, ALIGN_RIGHT,
    FMT_MONEY, FMT_PCT,
    apply_header_style, apply_data_style, style_title_cell, auto_width,
    PatternFill, Font, Alignment, Border, Side
)
from scripts.generators.xlsx_gen import ExcelGenerator

OUTPUT_DIR = PACK / "Modulo_4_Dashboard"

# ---- Datos de empleados ----
EMPLEADOS = [
    ("Ana Lopez Martinez", "Contador Senior"),
    ("Carlos Ramirez Ortiz", "Auxiliar Contable"),
    ("Maria Fernanda Torres", "Gerente Administrativo"),
    ("Jorge Hernandez Cruz", "Analista Fiscal"),
    ("Patricia Gonzalez Ruiz", "Asistente Administrativo"),
    ("Roberto Sanchez Diaz", "Director Financiero"),
    ("Laura Garcia Mendoza", "Contador Junior"),
    ("Fernando Morales Rios", "Auditor Interno"),
    ("Gabriela Flores Castillo", "Nominas y IMSS"),
    ("Miguel Angel Vargas", "Gerente de Operaciones"),
    ("Diana Castro Perez", "Auxiliar de Nominas"),
    ("Alejandro Reyes Luna", "Contador General"),
    ("Sofia Martinez Aguilar", "Analista de Costos"),
    ("Ricardo Jimenez Torres", "Jefe de Contabilidad"),
    ("Claudia Romero Navarro", "Asistente Fiscal"),
    ("Eduardo Gutierrez Soto", "Tesorero"),
    ("Veronica Diaz Herrera", "Coordinadora RH"),
    ("Oscar Mendez Rojas", "Pasante Contable"),
    ("Isabel Juarez Medina", "Ejecutiva de Cobranza"),
    ("Raul Dominguez Espino", "Gerente General"),
]

PUESTOS_SALARIOS = {
    "Pasante Contable": (SALARIO_MINIMO_MENSUAL, 10000),
    "Auxiliar Contable": (10000, 14000),
    "Auxiliar de Nominas": (9500, 13000),
    "Asistente Administrativo": (10000, 14000),
    "Asistente Fiscal": (10500, 14500),
    "Contador Junior": (14000, 20000),
    "Contador Senior": (20000, 30000),
    "Contador General": (25000, 35000),
    "Analista Fiscal": (18000, 26000),
    "Analista de Costos": (17000, 25000),
    "Nominas y IMSS": (15000, 22000),
    "Jefe de Contabilidad": (28000, 38000),
    "Auditor Interno": (22000, 32000),
    "Coordinadora RH": (20000, 28000),
    "Ejecutiva de Cobranza": (12000, 18000),
    "Tesorero": (25000, 35000),
    "Gerente Administrativo": (32000, 42000),
    "Gerente de Operaciones": (35000, 45000),
    "Director Financiero": (40000, 55000),
    "Gerente General": (45000, 55000),
}

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


def _calcular_imss(sueldo):
    """Cuota IMSS trabajador simplificada (~2.775% del SBC tope 25 UMA)."""
    uma_diario = 113.14  # UMA 2026 estimado
    tope_sbc = uma_diario * 25 * 30
    base = min(sueldo, tope_sbc)
    return round(base * 0.02775, 2)


def _calcular_subsidio(sueldo):
    """Subsidio al empleo simplificado para sueldos bajos."""
    from scripts.config.isr_2026 import SUBSIDIO_EMPLEO_MENSUAL
    for rango in SUBSIDIO_EMPLEO_MENSUAL:
        if rango["desde"] <= sueldo <= rango["hasta"]:
            return rango["subsidio"]
    return 0.0


def build():
    gen = ExcelGenerator("10_Dashboard_Final_Integrado.xlsx", OUTPUT_DIR)

    random.seed(2026)

    # ==== Hoja 1: Datos_Nomina ============================================
    ws1 = gen.add_sheet("Datos_Nomina")
    style_title_cell(ws1, 1, 1, "Datos de Nomina 2026 -- 20 Empleados x 12 Meses", 8)
    ws1.cell(row=2, column=1,
             value="Base de datos para Tablas Dinamicas, Segmentadores y Dashboard.").font = FONT_SMALL

    headers = [
        "Empleado", "Puesto", "Periodo", "Sueldo",
        "ISR", "IMSS", "SubsidioEmpleo", "NetoPagar"
    ]

    data = []
    for nombre, puesto in EMPLEADOS:
        sal_min, sal_max = PUESTOS_SALARIOS.get(puesto, (15000, 25000))
        sueldo_base = round(random.uniform(sal_min, sal_max), 2)

        for mes in MESES:
            # Small monthly variation (+/- 3% for bonuses/overtime)
            variacion = random.uniform(-0.03, 0.03)
            sueldo = round(sueldo_base * (1 + variacion), 2)

            isr_result = calcular_isr_mensual(sueldo)
            isr = isr_result.get("isr_total", 0.0) if isr_result else 0.0
            imss = _calcular_imss(sueldo)
            subsidio = _calcular_subsidio(sueldo)

            neto = round(sueldo - isr - imss + subsidio, 2)
            if neto < 0:
                neto = round(sueldo * 0.70, 2)

            data.append([nombre, puesto, mes, sueldo, isr, imss, subsidio, neto])

    gen.write_table(
        ws1, headers, data, start_row=4,
        table_name="Nomina_Empleados",
        money_cols=[4, 5, 6, 7, 8]
    )

    ws1.column_dimensions["A"].width = 32
    ws1.column_dimensions["B"].width = 26
    ws1.column_dimensions["C"].width = 14

    # ==== Hoja 2: Tarifa_ISR ==============================================
    ws2 = gen.add_sheet("Tarifa_ISR")
    style_title_cell(ws2, 1, 1, "Tarifa ISR Mensual 2026 -- Art. 96 LISR (Anexo 8 RMF)", 6)
    ws2.cell(row=2, column=1,
             value="Tarifa oficial para retenciones de nomina. Usa BUSCARV para calcular el ISR.").font = FONT_SMALL

    isr_headers = ["Limite Inferior", "Limite Superior", "Cuota Fija", "% Sobre Excedente"]
    isr_data = []
    for r in TARIFA_MENSUAL:
        sup = "En adelante" if r["lim_sup"] > 999_999_999 else r["lim_sup"]
        isr_data.append([r["lim_inf"], sup, r["cuota"], r["pct"] / 100])

    gen.write_table(
        ws2, isr_headers, isr_data, start_row=4,
        table_name="Tarifa_ISR_Mensual",
        money_cols=[1, 2, 3], pct_cols=[4]
    )

    # ==== Hoja 3: Calculadora =============================================
    ws3 = gen.add_sheet("Calculadora")
    style_title_cell(ws3, 1, 1, "Calculadora ISR Mensual con BUSCARV", 5)
    ws3.cell(row=2, column=1,
             value="Ingresa un sueldo mensual en B4 y las formulas calculan automaticamente el ISR.").font = FONT_SMALL

    labels = [
        (4, "Sueldo Mensual Bruto (base gravable)"),
        (5, "Limite Inferior (BUSCARV)"),
        (6, "Excedente sobre Limite Inferior"),
        (7, "% Sobre Excedente (BUSCARV)"),
        (8, "ISR Marginal"),
        (9, "Cuota Fija (BUSCARV)"),
        (10, "ISR Causado del Periodo"),
        (11, ""),
        (12, "IMSS Trabajador (estimado)"),
        (13, "Subsidio al Empleo"),
        (14, "Neto a Pagar"),
    ]

    for row, label in labels:
        if label:
            ws3.cell(row=row, column=1, value=label).font = FONT_NORMAL
            ws3.cell(row=row, column=1).border = THIN_BORDER
            ws3.cell(row=row, column=2).border = THIN_BORDER
            ws3.cell(row=row, column=2).number_format = FMT_MONEY
            ws3.cell(row=row, column=2).alignment = ALIGN_RIGHT

    # Input cell
    ws3.cell(row=4, column=2, value=25000).font = FONT_SUBTITULO
    ws3.cell(row=4, column=2).fill = FILL_AMARILLO

    # Formulas
    ws3["B5"] = "=VLOOKUP(B4,Tarifa_ISR_Mensual,1,TRUE)"
    ws3["B6"] = "=B4-B5"
    ws3["B7"] = "=VLOOKUP(B4,Tarifa_ISR_Mensual,4,TRUE)"
    ws3["B7"].number_format = FMT_PCT
    ws3["B8"] = "=B6*B7"
    ws3["B9"] = "=VLOOKUP(B4,Tarifa_ISR_Mensual,3,TRUE)"
    ws3["B10"] = "=B8+B9"
    ws3["B10"].fill = FILL_VERDE
    ws3["B10"].font = FONT_SUBTITULO

    # Deductions
    ws3["B12"] = "=MIN(B4,113.14*25*30)*0.02775"  # IMSS estimate
    ws3["B13"] = 0.00  # Subsidio - user fills or formula
    ws3["B14"] = "=B4-B10-B12+B13"
    ws3["B14"].fill = FILL_VERDE
    ws3["B14"].font = FONT_SUBTITULO

    # Explanations
    explanations = {
        5: "BUSCARV busca el limite inferior que corresponde a tu sueldo",
        6: "Sueldo menos el limite inferior de tu rango",
        7: "Porcentaje marginal de impuesto sobre el excedente",
        8: "Excedente x Porcentaje = ISR marginal",
        9: "Cantidad fija que se suma segun tu rango",
        10: "ISR Marginal + Cuota Fija = ISR total del periodo",
        12: "Estimacion: ~2.775% del SBC (tope 25 UMA)",
        13: "Aplica solo a sueldos bajos (ver tabla de subsidio)",
        14: "Sueldo - ISR - IMSS + Subsidio = Neto a pagar",
    }
    for row, text in explanations.items():
        ws3.cell(row=row, column=3, value=text).font = FONT_SMALL

    ws3.column_dimensions["A"].width = 48
    ws3.column_dimensions["B"].width = 22
    ws3.column_dimensions["C"].width = 55

    # ==== Hoja 4: Dashboard ===============================================
    ws4 = gen.add_sheet("Dashboard")
    style_title_cell(ws4, 1, 1, "Dashboard de Nomina Integrado", 12)

    # Background
    bg_fill = PatternFill("solid", fgColor=Color.FONDO_CLARO)
    for r in range(1, 30):
        for c in range(1, 14):
            ws4.cell(row=r, column=c).fill = bg_fill

    ws4.cell(row=1, column=1).fill = PatternFill("solid", fgColor=Color.BLANCO)
    ws4.row_dimensions[1].height = 35

    # ---- KPI Row (row 3-4) ----
    kpi_defs = [
        ("B3", "Total Percepciones", Color.AZUL,
         "B4", '=SUBTOTAL(109,Nomina_Empleados[Sueldo])'),
        ("E3", "Total Deducciones (ISR+IMSS)", Color.ROJO,
         "E4", '=SUBTOTAL(109,Nomina_Empleados[ISR])+SUBTOTAL(109,Nomina_Empleados[IMSS])'),
        ("H3", "ISR del Periodo", Color.AMARILLO,
         "H4", '=SUBTOTAL(109,Nomina_Empleados[ISR])'),
        ("K3", "e.firma Status", Color.VERDE,
         "K4", "VIGENTE"),
    ]

    for title_cell_ref, title, color, val_cell_ref, formula in kpi_defs:
        # Title
        cell_t = ws4[title_cell_ref]
        cell_t.value = title
        cell_t.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        cell_t.fill = PatternFill("solid", fgColor=color)
        cell_t.alignment = ALIGN_CENTER
        cell_t.border = THIN_BORDER

        # Merge title across 2 cols
        t_row = cell_t.row
        t_col = cell_t.column
        ws4.merge_cells(start_row=t_row, start_column=t_col,
                        end_row=t_row, end_column=t_col + 1)
        ws4.cell(row=t_row, column=t_col + 1).fill = PatternFill("solid", fgColor=color)
        ws4.cell(row=t_row, column=t_col + 1).border = THIN_BORDER

        # Value
        cell_v = ws4[val_cell_ref]
        cell_v.value = formula
        cell_v.font = Font(name="Calibri", bold=True, size=14, color=color)
        cell_v.alignment = ALIGN_CENTER
        cell_v.border = THIN_BORDER
        if "SUBTOTAL" in str(formula):
            cell_v.number_format = FMT_MONEY

        # Merge value across 2 cols
        v_row = cell_v.row
        v_col = cell_v.column
        ws4.merge_cells(start_row=v_row, start_column=v_col,
                        end_row=v_row, end_column=v_col + 1)
        ws4.cell(row=v_row, column=v_col + 1).border = THIN_BORDER

    ws4.row_dimensions[3].height = 25
    ws4.row_dimensions[4].height = 35

    # ---- Instructions in Dashboard ----
    instructions_start = 7
    dash_instructions = [
        "INSTRUCCIONES PARA COMPLETAR EL DASHBOARD:",
        "",
        "1. TABLAS DINAMICAS:",
        "   a) Selecciona cualquier celda en Datos_Nomina > Insertar > Tabla Dinamica",
        "   b) Crea una TD con Periodo en filas, Sueldo/ISR/NetoPagar en valores (Suma)",
        "   c) Crea otra TD con Puesto en filas, Empleado en valores (Cuenta)",
        "",
        "2. SEGMENTADORES (SLICERS):",
        "   a) Clic en tu Tabla Dinamica > Insertar > Segmentacion de datos",
        "   b) Selecciona: Periodo, Puesto, Empleado",
        "   c) Conecta los slicers a TODAS las TDs: clic derecho > Conexiones de informe",
        "",
        "3. GRAFICOS:",
        "   a) Selecciona la TD > Insertar > Grafico > Barras agrupadas (para comparar puestos)",
        "   b) Inserta un grafico de lineas basado en la TD mensual (tendencia de nomina)",
        "   c) Mueve los graficos a esta hoja de Dashboard",
        "",
        "4. KPIs:",
        "   Los KPIs de arriba ya tienen formulas SUBTOTAL que respetan filtros de slicers.",
        "   Si usas Tablas Dinamicas, reemplaza con =GETPIVOTDATA() para mayor precision.",
        "",
        "5. FORMATO FINAL:",
        "   a) Oculta lineas de cuadricula: Vista > desmarcar Lineas de cuadricula",
        "   b) Ajusta el zoom a 85-90% para ver todo el dashboard",
        "   c) Usa Vista > Inmovilizar paneles en fila 5 para fijar los KPIs",
    ]

    for i, line in enumerate(dash_instructions):
        row = instructions_start + i
        ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        cell = ws4.cell(row=row, column=1, value=line)
        if i == 0:
            cell.font = FONT_SUBTITULO
        else:
            cell.font = FONT_NORMAL
        cell.alignment = ALIGN_LEFT

    # Column widths
    for c in range(1, 14):
        ws4.column_dimensions[get_column_letter(c)].width = 14

    ws4.sheet_properties.tabColor = Color.AZUL

    # ==== Hoja 5: Instrucciones ===========================================
    gen.add_instructions_sheet([
        "Este archivo integra todo el Modulo 4: Nomina + ISR + Dashboard.",
        "La hoja 'Datos_Nomina' tiene 240 registros (20 empleados x 12 meses) como tabla nombrada 'Nomina_Empleados'.",
        "La hoja 'Tarifa_ISR' contiene la tarifa mensual Art. 96 como tabla nombrada 'Tarifa_ISR_Mensual'.",
        "En 'Calculadora', cambia el sueldo amarillo (B4) para ver el calculo ISR paso a paso con BUSCARV.",
        "La hoja 'Dashboard' tiene KPIs con formulas SUBTOTAL que se actualizan con los segmentadores.",
        "PASO 1: Crea Tablas Dinamicas desde Datos_Nomina (una por meses, otra por puestos).",
        "PASO 2: Inserta Segmentadores vinculados a ambas Tablas Dinamicas.",
        "PASO 3: Crea graficos (barras y lineas) y muevalos a la hoja Dashboard.",
        "PASO 4: Protege las hojas de formulas y comparte como PDF o Excel protegido.",
        "Los datos simulan sueldos mexicanos reales desde salario minimo hasta $55,000/mes.",
    ])

    gen.save()


if __name__ == "__main__":
    build()
