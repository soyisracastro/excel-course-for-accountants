"""
Generador: Referencia_Modulo_5.md (~5 paginas)
Modulo 5 -- Automatizacion Nativa con Microsoft 365 Copilot

Contenido:
  - Copilot checklist
  - Top 15 prompts quick reference
  - Referencia de IA externa (ChatGPT, Gemini, Claude)
  - Resumen del curso M1-M5
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK, MODULOS
from scripts.generators.md_gen import MarkdownGenerator

inch = 72  # compat: MarkdownGenerator ignores col_widths

OUTPUT_DIR = PACK / "Modulo_5_Copilot_IA"


def build():
    pdf = MarkdownGenerator(
        "Referencia_Modulo_5.md",
        OUTPUT_DIR,
        title="Referencia Rápida - Módulo 5: Copilot e IA",
    )

    # -- Portada ---------------------------------------------------------------
    pdf.add_cover(
        title="Referencia Rápida - Módulo 5",
        subtitle="Copilot, IA externa y resumen del curso completo",
        modulo="Módulo 5 - Automatización Nativa con Microsoft 365 Copilot",
    )

    # =========================================================================
    # SECCIÓN 1: Copilot Checklist
    # =========================================================================
    pdf.add_section("1. Checklist para usar Copilot en Excel")
    pdf.add_text(
        "Antes de usar Microsoft 365 Copilot en Excel, verifica que cumples "
        "con todos los requisitos. Marca cada punto:"
    )
    pdf.add_spacer(0.1)

    checklist_items = [
        ("Licencia activa", "Microsoft 365 Business/Enterprise con add-on Copilot habilitado."),
        ("Archivo en la nube", "Guardado en OneDrive o SharePoint. No funciona con archivos locales."),
        ("Formato Tabla", "Datos convertidos a Tabla (Ctrl+T) con nombre descriptivo (ej. Ventas_Gasolinera)."),
        ("Sin celdas combinadas", "Eliminar celdas combinadas que interfieran con la estructura de la tabla."),
        ("Sin filas en blanco", "Remover filas vacías dentro del rango de datos."),
        ("Encabezados claros", "Nombres de columna descriptivos, sin espacios al inicio/final."),
        ("Conexión a internet", "Copilot procesa en la nube; necesitas conexión estable."),
        ("Panel de Copilot visible", "Clic en ícono de Copilot en pestaña Inicio. Si está gris, revisa requisitos."),
    ]

    table_data = [["#", "Requisito", "Detalle"]]
    for i, (req, det) in enumerate(checklist_items, 1):
        table_data.append([str(i), req, det])

    pdf.add_table(table_data, col_widths=[0.4 * inch, 1.8 * inch, 4.8 * inch])
    pdf.add_spacer(0.2)

    # =========================================================================
    # SECCIÓN 2: Top 15 Prompts Quick Reference
    # =========================================================================
    pdf.add_section("2. Top 15 Prompts - Referencia Rápida")
    pdf.add_text(
        "Los 15 prompts más útiles para contadores, organizados por categoría. "
        "Copia y pega directamente en el panel de Copilot."
    )
    pdf.add_spacer(0.1)

    prompts_ref = [
        ("Análisis", "Analiza las ventas por sucursal y dime cuál tiene mejor desempeño."),
        ("Análisis", "¿Cuál vendedor tiene el mejor desempeño? Muestra ranking de los 5 mejores."),
        ("Análisis", "¿Cuál es la tendencia de ventas mes a mes durante 2025?"),
        ("Anomalías", "Identifica anomalías en la tabla de nómina."),
        ("Anomalías", "¿Hay datos faltantes? ¿Qué empleados tienen meses sin registro?"),
        ("Anomalías", "Detecta empleados con horas extra inusualmente altas."),
        ("Cálculos", "Crea columna que calcule el ISR marginal por TotalPercepcion."),
        ("Cálculos", "Calcula comisión del 2% sobre Venta_Total como nueva columna."),
        ("Cálculos", "Clasifica cada venta como Alta (>$5,000), Media ($1,000-$5,000) o Baja (<$1,000)."),
        ("Gráficos", "Crea gráfico de barras de ventas totales por mes."),
        ("Gráficos", "Muestra distribución por tipo de combustible con gráfico de pastel."),
        ("Gráficos", "Genera gráfico comparativo de ventas por turno para cada sucursal."),
        ("Resumen", "Genera resumen ejecutivo: totales, promedios, mejor sucursal, mejor vendedor."),
        ("Resumen", "Crea tabla de frecuencia por rango de litros (0-50, 50-100, 100-200, 200-500)."),
        ("Resumen", "Resume deducciones totales por tipo (ISR, IMSS, Otras) con porcentajes."),
    ]

    prompt_table = [["#", "Categoría", "Prompt"]]
    for i, (cat, prompt) in enumerate(prompts_ref, 1):
        prompt_table.append([str(i), cat, prompt])

    pdf.add_table(prompt_table, col_widths=[0.35 * inch, 1.0 * inch, 5.65 * inch])

    pdf.add_page_break()

    # =========================================================================
    # SECCIÓN 3: Referencia de IA Externa
    # =========================================================================
    pdf.add_section("3. Referencia de IA Externa")
    pdf.add_text(
        "Además de Microsoft Copilot, estas herramientas de IA pueden potenciar tu trabajo contable:"
    )
    pdf.add_spacer(0.15)

    pdf.add_subsection("ChatGPT (OpenAI)")
    pdf.add_bullet("URL: chat.openai.com")
    pdf.add_bullet("Fortaleza: Generación de macros VBA, consultas Power Query M, explicación de fórmulas.")
    pdf.add_bullet("Uso típico: 'Genera macro VBA que exporte cada hoja como PDF individual.'")
    pdf.add_bullet("Costo: Gratuito (GPT-3.5) o $20 USD/mes (GPT-4, Plus).")
    pdf.add_spacer(0.1)

    pdf.add_subsection("Google Gemini")
    pdf.add_bullet("URL: gemini.google.com | Integrado en Google Sheets")
    pdf.add_bullet("Fortaleza: Análisis de datos en Google Sheets, búsqueda contextual.")
    pdf.add_bullet("Uso típico: Asistente en Google Sheets para fórmulas y análisis rápido.")
    pdf.add_bullet("Costo: Gratuito con cuenta Google; Gemini Advanced $20 USD/mes.")
    pdf.add_spacer(0.1)

    pdf.add_subsection("Claude (Anthropic)")
    pdf.add_bullet("URL: claude.ai")
    pdf.add_bullet("Fortaleza: Análisis de documentos extensos, precisión en razonamiento, código limpio.")
    pdf.add_bullet("Uso típico: 'Analiza este estado financiero y detecta inconsistencias.'")
    pdf.add_bullet("Costo: Gratuito (limitado) o $20 USD/mes (Pro).")
    pdf.add_spacer(0.15)

    pdf.add_subsection("Comparativa rápida")
    comp_table = [
        ["Herramienta", "Mejor para", "Integración Excel", "Costo base"],
        ["MS Copilot", "Análisis dentro de Excel", "Nativa (panel lateral)", "$30 USD/mes"],
        ["ChatGPT", "Macros VBA, Power Query", "Copiar/pegar código", "Gratis / $20"],
        ["Gemini", "Google Sheets, búsqueda", "Google Sheets nativo", "Gratis / $20"],
        ["Claude", "Documentos, razonamiento", "Copiar/pegar análisis", "Gratis / $20"],
    ]
    pdf.add_table(comp_table, col_widths=[1.2 * inch, 1.8 * inch, 2.0 * inch, 1.3 * inch])

    pdf.add_page_break()

    # =========================================================================
    # SECCIÓN 4: Resumen del Curso M1-M5
    # =========================================================================
    pdf.add_section("4. Resumen del Curso Completo: M1 a M5")
    pdf.add_text(
        "A lo largo de 5 módulos, este curso te llevó del dato crudo al insight accionable "
        "con herramientas modernas de Excel."
    )
    pdf.add_spacer(0.15)

    for mod_num in sorted(MODULOS.keys()):
        mod = MODULOS[mod_num]
        pdf.add_subsection("Módulo {}: {}".format(mod_num, mod["nombre"]))

    # Module 1 details
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Módulo 1 — Lógica Contable y Funciones de Control</b>")
    pdf.add_bullet("Funciones clave: BUSCARV, SI, SI.CONJUNTO, TRUNCAR")
    pdf.add_bullet("Cálculo de ISR con tarifa Art. 96 y Art. 152 LISR 2026")
    pdf.add_bullet("Factor de Actualización Art. 17-A CFF con TRUNCAR a 4 decimales")
    pdf.add_bullet("Extracción y validación de RFC con funciones de texto")
    pdf.add_spacer(0.1)

    pdf.add_text("<b>Módulo 2 — Procesamiento Masivo y Tablas Dinámicas</b>")
    pdf.add_bullet("Limpieza y preparación de datos masivos")
    pdf.add_bullet("Tablas dinámicas: agrupación, filtros, campos calculados")
    pdf.add_bullet("Análisis de datos reales de ventas y nómina")
    pdf.add_bullet("Segmentadores y líneas de tiempo para filtrado interactivo")
    pdf.add_spacer(0.1)

    pdf.add_text("<b>Módulo 3 — Visualización Profesional</b>")
    pdf.add_bullet("Gráficos de impacto: barras, líneas, combinados, cascada")
    pdf.add_bullet("Principios de diseño para reportes ejecutivos")
    pdf.add_bullet("Formateo condicional avanzado con barras, semáforos, íconos")
    pdf.add_bullet("Minigráficos (Sparklines) para tendencias en celdas")
    pdf.add_spacer(0.1)

    pdf.add_text("<b>Módulo 4 — El Dashboard Inteligente</b>")
    pdf.add_bullet("KPIs con formato profesional y semáforos de alerta")
    pdf.add_bullet("Controles interactivos: segmentadores, casillas, listas")
    pdf.add_bullet("Diseño de dashboard ejecutivo en una sola hoja")
    pdf.add_bullet("Exportación y entrega profesional (PDF, protección)")
    pdf.add_spacer(0.1)

    pdf.add_text("<b>Módulo 5 — Copilot e Inteligencia Artificial</b>")
    pdf.add_bullet("Microsoft 365 Copilot: requisitos, activación, panel lateral")
    pdf.add_bullet("Análisis con lenguaje natural, generación de fórmulas, gráficos")
    pdf.add_bullet("Detección de anomalías y columnas inteligentes")
    pdf.add_bullet("IA externa: ChatGPT, Gemini, Claude como complemento")
    pdf.add_spacer(0.2)

    # =========================================================================
    # SECCIÓN 5: Consejos finales
    # =========================================================================
    pdf.add_section("5. Consejos finales")
    pdf.add_bullet("Practica diariamente con los archivos del Pack Excel Pro.")
    pdf.add_bullet("Automatiza primero las tareas que haces todos los meses (nómina, conciliaciones, reportes).")
    pdf.add_bullet("Usa Copilot para explorar datos rápidamente, pero siempre valida con tu criterio.")
    pdf.add_bullet("Comparte lo aprendido con tu equipo: la productividad se multiplica.")
    pdf.add_bullet("Mantente actualizado: Excel y Copilot se actualizan constantemente.")
    pdf.add_bullet("Siguiente nivel: Power Query, Power BI, Power Automate, VBA.")
    pdf.add_spacer(0.2)

    pdf.add_text(
        "<b>Recuerda:</b> La tecnología es la herramienta, pero el profesional eres tú. "
        "Tu criterio contable, tu experiencia y tu ética profesional son insustituibles. "
        "La IA te hace más eficiente, pero no te reemplaza."
    )

    pdf.save()


if __name__ == "__main__":
    build()
