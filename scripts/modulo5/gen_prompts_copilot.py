"""
Generador: Guia_Prompts_Copilot_Contadores.md
Modulo 5 -- Automatizacion Nativa con Microsoft 365 Copilot

Markdown con 20 prompts organizados en 5 categorias para usar
con Microsoft 365 Copilot en Excel.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.generators.md_gen import MarkdownGenerator

OUTPUT_DIR = PACK / "Modulo_5_Copilot_IA"

# -- Catálogo de prompts -----------------------------------------------------

PROMPTS = {
    "1. Análisis de Datos": [
        {
            "prompt": "Analiza las ventas por sucursal y dime cuál tiene mejor desempeño en los últimos 3 meses.",
            "esperar": "Copilot generará una tabla resumen con totales por sucursal (Centro, Norte, Sur) filtrando Oct-Dic 2025, indicando cuál tiene mayor volumen de ventas.",
            "validar": "Crea una tabla dinámica manual con filtro de fecha Oct-Dic y agrupa por Sucursal. Compara los totales con la respuesta de Copilot.",
        },
        {
            "prompt": "¿Cuál vendedor tiene el mejor desempeño en ventas totales? Muestra un ranking de los 5 mejores.",
            "esperar": "Un ranking ordenado de vendedores por Venta_Total acumulada. Vendedor_3 debería aparecer en los últimos lugares.",
            "validar": "Usa SUMAR.SI para acumular Venta_Total por vendedor y ordena de mayor a menor. Verifica que Vendedor_3 esté abajo.",
        },
        {
            "prompt": "Compara el volumen de litros vendidos por tipo de combustible entre las sucursales.",
            "esperar": "Una tabla cruzada Sucursal vs TipoCombustible mostrando suma de litros. Norte debería mostrar mayor proporción de Premium.",
            "validar": "Crea tabla dinámica con Sucursal en filas, TipoCombustible en columnas y Suma de Litros en valores.",
        },
        {
            "prompt": "¿Cuál es la tendencia de ventas mes a mes durante 2025? ¿Hay alguna estacionalidad?",
            "esperar": "Un análisis temporal con ventas mensuales mostrando si hay meses altos o bajos. Copilot puede identificar tendencias.",
            "validar": "Agrupa las fechas por mes con una tabla dinámica y grafica la serie temporal. Observa si coincide con el análisis de Copilot.",
        },
        {
            "prompt": "¿Qué porcentaje de las ventas se pagan con cada método de pago? Desglósalo por sucursal.",
            "esperar": "Porcentajes de Efectivo, Tarjeta y Transferencia por sucursal en formato tabla o gráfico.",
            "validar": "Usa CONTAR.SI.CONJUNTO para contar transacciones por MetodoPago y Sucursal. Calcula los porcentajes manualmente.",
        },
    ],
    "2. Detección de Errores y Anomalías": [
        {
            "prompt": "Identifica anomalías en la tabla de nómina. ¿Hay empleados con cambios inusuales de sueldo?",
            "esperar": "Copilot debería detectar los 2 empleados con incrementos súbitos de sueldo (empleado 3 en julio, empleado 11 en octubre).",
            "validar": "Filtra por cada empleado y grafica su SueldoBase por periodo. Los saltos serán visibles como picos en la línea.",
        },
        {
            "prompt": "¿Hay datos faltantes en la nómina? ¿Qué empleados tienen meses sin registro?",
            "esperar": "Debería identificar al empleado con 3 meses faltantes (abril, mayo, junio 2025).",
            "validar": "Usa CONTAR.SI para contar registros por empleado. El que tenga menos de 12 registros mensuales tiene meses faltantes.",
        },
        {
            "prompt": "Detecta si hay empleados con horas extra inusualmente altas. ¿En qué periodos ocurre?",
            "esperar": "Copilot debería señalar que diciembre tiene picos de horas extra en todos los empleados.",
            "validar": "Calcula el promedio de HorasExtra por periodo. Diciembre debería tener un promedio significativamente mayor.",
        },
        {
            "prompt": "Revisa si algún vendedor tiene un rendimiento consistentemente bajo comparado con el promedio.",
            "esperar": "Identificación de Vendedor_3 como el de menor rendimiento sistemático (transacciones pequeñas).",
            "validar": "Calcula promedio de Litros y Venta_Total por vendedor. Vendedor_3 tendrá promedios notablemente menores.",
        },
    ],
    "3. Cálculos y Fórmulas": [
        {
            "prompt": "Crea una columna que calcule el ISR marginal para cada empleado basado en su TotalPercepcion mensual.",
            "esperar": "Copilot agregará una columna con fórmula que aplique la tarifa Art. 96 LISR, ubicando el rango y aplicando el porcentaje correspondiente.",
            "validar": "Compara los valores de la nueva columna con la columna ISR existente. Deben ser iguales o muy cercanos.",
        },
        {
            "prompt": "Calcula una comisión del 2% sobre Venta_Total para cada vendedor y agrégala como nueva columna.",
            "esperar": "Una nueva columna 'Comisión' con la fórmula =Venta_Total*0.02 aplicada a todas las filas.",
            "validar": "Verifica manualmente: multiplica Venta_Total por 0.02 en algunas filas y compara.",
        },
        {
            "prompt": "Agrega una columna que clasifique cada venta como 'Alta' (>$5,000), 'Media' ($1,000-$5,000) o 'Baja' (<$1,000).",
            "esperar": "Copilot creará una columna con función SI anidada o IFS que clasifique por rango de Venta_Total.",
            "validar": "Filtra por cada categoría y verifica que los montos correspondan a los rangos definidos.",
        },
        {
            "prompt": "Calcula el sueldo neto promedio por puesto y ordena de mayor a menor.",
            "esperar": "Un resumen con el promedio de NetoPagar agrupado por Puesto, ordenado descendentemente.",
            "validar": "Usa PROMEDIO.SI para calcular el promedio de NetoPagar por cada puesto único.",
        },
    ],
    "4. Gráficos y Visualización": [
        {
            "prompt": "Crea un gráfico de barras que muestre las ventas totales por mes durante 2025.",
            "esperar": "Un gráfico de barras verticales con 12 barras (Ene-Dic) mostrando la suma de Venta_Total por mes.",
            "validar": "Crea tu propio gráfico con tabla dinámica de Fecha (agrupada por mes) vs Suma de Venta_Total.",
        },
        {
            "prompt": "Muestra la distribución de ventas por tipo de combustible con un gráfico de pastel.",
            "esperar": "Un gráfico circular con 3 segmentos (Magna, Premium, Diesel) mostrando proporción de ventas.",
            "validar": "Suma Venta_Total por TipoCombustible y crea un gráfico circular manual para comparar.",
        },
        {
            "prompt": "Genera un gráfico de líneas que muestre la evolución del sueldo base de los 5 empleados con mayor sueldo.",
            "esperar": "Un gráfico de líneas con 5 series temporales mostrando SueldoBase por periodo.",
            "validar": "Identifica los 5 empleados con mayor SueldoBase y grafícalos manualmente con tabla dinámica.",
        },
        {
            "prompt": "Crea un gráfico comparativo de ventas por turno (Matutino, Vespertino, Nocturno) para cada sucursal.",
            "esperar": "Un gráfico de barras agrupadas con 3 grupos (sucursales) y 3 barras cada uno (turnos).",
            "validar": "Tabla dinámica con Sucursal en filas, Turno en columnas y Suma de Venta_Total en valores.",
        },
    ],
    "5. Automatización y Resumen": [
        {
            "prompt": "Genera un resumen ejecutivo de la tabla de ventas: totales, promedios, mejor sucursal, mejor vendedor y tendencia.",
            "esperar": "Un párrafo o tabla con KPIs principales: venta total, promedio por transacción, sucursal líder, vendedor estrella.",
            "validar": "Calcula cada KPI manualmente con funciones SUMA, PROMEDIO, MAX, y verifica que coincidan.",
        },
        {
            "prompt": "Crea una tabla de frecuencia que muestre cuántas transacciones hay por rango de litros (0-50, 50-100, 100-200, 200-500).",
            "esperar": "Una tabla con 4 filas mostrando el conteo de transacciones en cada rango de litros.",
            "validar": "Usa CONTAR.SI.CONJUNTO con criterios de rango para contar transacciones en cada intervalo.",
        },
        {
            "prompt": "Resume las deducciones totales por tipo (ISR, IMSS, Otras) para toda la nómina y calcula el porcentaje que representa cada una.",
            "esperar": "Una tabla resumen con 3 filas: ISR total, IMSS total, OtrasDeducciones total, y su porcentaje del total de deducciones.",
            "validar": "Suma cada columna de deducciones y calcula el porcentaje de cada una sobre TotalDeduccion.",
        },
    ],
}


def build():
    pdf = MarkdownGenerator(
        "Guia_Prompts_Copilot_Contadores.md",
        OUTPUT_DIR,
        title="Guía de Prompts - Copilot para Contadores",
    )

    # -- Portada ---------------------------------------------------------------
    pdf.add_cover(
        title="Guía de Prompts para Copilot",
        subtitle="20 prompts listos para usar con datos contables en Excel",
        modulo="Módulo 5 — Automatización Nativa con Microsoft 365 Copilot",
    )

    # -- Introducción ----------------------------------------------------------
    pdf.add_section("Cómo usar esta guía")
    pdf.add_text(
        "Esta guía contiene 20 prompts diseñados específicamente para contadores y "
        "administrativos que usan Microsoft 365 Copilot en Excel. Cada prompt está "
        "pensado para trabajar con el archivo <b>12_Dataset_Master_Copilot.xlsx</b> "
        "incluido en este módulo."
    )
    pdf.add_spacer(0.15)
    pdf.add_text("<b>Requisitos previos:</b>")
    pdf.add_bullet("Licencia Microsoft 365 con Copilot habilitado.")
    pdf.add_bullet("Archivo guardado en OneDrive o SharePoint (obligatorio).")
    pdf.add_bullet("Datos en formato Tabla con nombre (ya configurado en el archivo).")
    pdf.add_spacer(0.15)
    pdf.add_text("<b>Estructura de cada prompt:</b>")
    pdf.add_bullet("<b>Prompt exacto:</b> Lo que escribirás en el panel de Copilot.")
    pdf.add_bullet("<b>Qué esperar:</b> La respuesta esperada de la IA.")
    pdf.add_bullet("<b>Cómo validar:</b> Cómo verificar que la respuesta sea correcta con funciones de Excel.")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Recuerda:</b> Copilot es una herramienta de apoyo, no un sustituto de tu "
        "criterio profesional. Siempre valida los resultados."
    )

    pdf.add_page_break()

    # -- Prompts por categoría -------------------------------------------------
    prompt_num = 0
    for cat_name, prompts in PROMPTS.items():
        pdf.add_section(cat_name)
        pdf.add_spacer(0.1)

        for p in prompts:
            prompt_num += 1
            pdf.add_subsection("Prompt {:d}".format(prompt_num))
            pdf.add_spacer(0.05)

            # Prompt box
            pdf.add_text("<b>Prompt exacto:</b>")
            pdf.add_code('"{}"'.format(p["prompt"]))
            pdf.add_spacer(0.08)

            pdf.add_text("<b>Qué esperar:</b>")
            pdf.add_text(p["esperar"])
            pdf.add_spacer(0.08)

            pdf.add_text("<b>Cómo validar:</b>")
            pdf.add_text(p["validar"])
            pdf.add_spacer(0.2)

        pdf.add_page_break()

    # -- Consejos finales ------------------------------------------------------
    pdf.add_section("Consejos para mejores resultados con Copilot")
    pdf.add_bullet("Sé específico: en lugar de 'analiza los datos', di exactamente qué columnas y qué operación.")
    pdf.add_bullet("Menciona el nombre de la tabla: 'En la tabla Ventas_Gasolinera, calcula...'")
    pdf.add_bullet("Pide un paso a la vez: no combines múltiples solicitudes en un solo prompt.")
    pdf.add_bullet("Si la respuesta no es correcta, reformula el prompt con más detalle.")
    pdf.add_bullet("Usa Copilot para explorar, pero siempre valida con fórmulas tradicionales.")
    pdf.add_bullet("Guarda las fórmulas útiles que Copilot genere para reutilizarlas.")
    pdf.add_spacer(0.3)

    pdf.add_section("Limitaciones actuales de Copilot en Excel")
    pdf.add_bullet("No puede acceder a archivos locales; el archivo debe estar en la nube (OneDrive/SharePoint).")
    pdf.add_bullet("Solo trabaja con datos en formato Tabla (Ctrl+T).")
    pdf.add_bullet("Puede generar fórmulas incorrectas; siempre verifica los resultados.")
    pdf.add_bullet("No reemplaza el criterio contable profesional (NIF, LISR, CFF).")
    pdf.add_bullet("Disponibilidad limitada a ciertos planes de Microsoft 365.")
    pdf.add_bullet("Las respuestas pueden variar si repites el mismo prompt.")

    pdf.save()


if __name__ == "__main__":
    build()
