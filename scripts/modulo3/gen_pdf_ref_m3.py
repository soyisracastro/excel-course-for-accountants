"""
Generador: Referencia_Modulo_3.md
Modulo 3 -- Visualizacion de Impacto y Reportes Ejecutivos

Contenido (~5 paginas):
  - Guia de seleccion de graficos (arbol de decision en texto)
  - Checklist de limpieza visual
  - Referencia de paleta de colores
  - 3 ejercicios practicos
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK, MODULOS
from scripts.generators.md_gen import MarkdownGenerator

inch = 72  # compat: MarkdownGenerator ignores col_widths

OUTPUT_DIR = PACK / "Modulo_3_Visualizacion"
MOD = MODULOS[3]


def build():
    # type: () -> Path
    pdf = MarkdownGenerator(
        filename="Referencia_Modulo_3.md",
        output_dir=OUTPUT_DIR,
        title="Referencia Módulo 3 - Visualización de Impacto",
    )

    # ── Portada ───────────────────────────────────────────────────
    pdf.add_cover(
        title="Guía de Referencia",
        subtitle="Visualización de datos, selección de gráficos y mejores prácticas",
        modulo="Módulo 3: {}".format(MOD["nombre"]),
    )

    # ── Sección 1: Guía de selección de gráficos ─────────────────
    pdf.add_section("1. Guía de selección de gráficos")

    pdf.add_text(
        "Antes de crear cualquier gráfico, responde esta pregunta: "
        "<b>¿Qué quiero comunicar?</b> La respuesta determina el tipo de gráfico ideal."
    )
    pdf.add_spacer(0.15)

    pdf.add_subsection("Árbol de decisión")
    pdf.add_text("Sigue este flujo para elegir el gráfico correcto:")
    pdf.add_spacer(0.1)

    decision_tree = [
        ["Pregunta", "Respuesta", "Tipo de gráfico"],
        ["¿Quiero comparar cantidades entre categorías?", "Sí", "Barras (horizontales o verticales)"],
        ["¿Quiero mostrar tendencia en el tiempo?", "Sí", "Líneas"],
        ["¿Quiero mostrar proporción del total?", "Sí, 5-6 categorías máx", "Pastel o Dona"],
        ["¿Quiero composición a lo largo del tiempo?", "Sí", "Columnas apiladas"],
        ["¿Quiero relación entre dos variables?", "Sí", "Dispersión (scatter)"],
        ["¿Quiero mostrar progreso hacia una meta?", "Sí", "Indicador o barra de progreso"],
    ]

    pdf.add_table(decision_tree, col_widths=[2.5 * inch, 1.5 * inch, 2.5 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Errores comunes en la selección")
    pdf.add_bullet("Usar pastel con más de 6 categorías (se vuelve ilegible)")
    pdf.add_bullet("Usar líneas para datos categóricos sin orden temporal")
    pdf.add_bullet("Usar efecto 3D (distorsiona las proporciones)")
    pdf.add_bullet("Usar doble eje Y sin justificación clara")
    pdf.add_bullet("Usar gráficos complejos cuando uno simple basta")

    # ── Sección 2: Checklist de limpieza visual ──────────────────
    pdf.add_page_break()
    pdf.add_section("2. Checklist de limpieza visual")

    pdf.add_text(
        "Después de crear tu gráfico, revisa cada punto de esta lista. "
        "El objetivo es eliminar todo lo que no comunica información útil."
    )
    pdf.add_spacer(0.15)

    checklist = [
        ["#", "Verificación", "Acción"],
        ["1", "Título descriptivo", "Cambiar 'Gráfico 1' por título que responda '¿de qué es este gráfico?'"],
        ["2", "¿Leyenda necesaria?", "Si solo hay 1 serie, eliminar la leyenda"],
        ["3", "Botones de campo", "En gráficos dinámicos: clic derecho > Ocultar botones de campo"],
        ["4", "Líneas de cuadrícula", "Eliminar si el gráfico es simple y las etiquetas ya muestran valores"],
        ["5", "Etiquetas de datos", "Agregar valores sobre barras/puntos si son pocos datos"],
        ["6", "Formato de ejes", "Usar K (miles), M (millones). Quitar decimales innecesarios"],
        ["7", "Colores", "Máximo 3-4 colores. Consistentes con el significado"],
        ["8", "Tamaño de fuente", "Mínimo 10pts para etiquetas, 14pts para títulos"],
        ["9", "Orden de datos", "Ordenar barras de mayor a menor (o cronológico si aplica)"],
        ["10", "Fuente de datos", "Incluir periodo y unidad (ej: 'Ventas 2025, cifras en MDP')"],
    ]

    pdf.add_table(checklist, col_widths=[0.4 * inch, 1.8 * inch, 4.0 * inch])

    pdf.add_spacer(0.2)
    pdf.add_text(
        '<b>Principio de Edward Tufte:</b> "Maximiza la tinta de datos, '
        'minimiza la tinta de decoración." Cada píxel debe tener un propósito.'
    )

    # ── Sección 3: Paleta de colores ─────────────────────────────
    pdf.add_page_break()
    pdf.add_section("3. Paleta de colores de referencia")

    pdf.add_text(
        "Usa estos colores de forma consistente en todos tus reportes. "
        "Cada color tiene un significado intuitivo que facilita la lectura."
    )
    pdf.add_spacer(0.15)

    colores = [
        ["Color", "Código HEX", "Uso recomendado", "Ejemplo"],
        ["Azul", "#2563EB", "Color principal / institucional", "Títulos, barras principales"],
        ["Verde", "#10B981", "Positivo / crecimiento / Magna", "Utilidades, cumplimiento, Magna"],
        ["Rojo", "#EF4444", "Atención / negativo / Premium", "Gastos excesivos, alertas, Premium"],
        ["Amarillo", "#F59E0B", "Precaución / intermedio", "Advertencias, datos pendientes"],
        ["Gris", "#64748B", "Secundario / referencia / Diesel", "Datos de periodo anterior, Diesel"],
        ["Gris claro", "#CBD5E1", "Bordes / fondos", "Líneas de cuadrícula, bordes de tabla"],
    ]

    pdf.add_table(colores, col_widths=[0.9 * inch, 1.0 * inch, 2.3 * inch, 2.3 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Ejemplo aplicado: Gasolinera")
    pdf.add_bullet("Magna = Verde (#10B981) — la bomba verde que todos conocen")
    pdf.add_bullet("Premium = Rojo (#EF4444) — la bomba roja de alto octanaje")
    pdf.add_bullet("Diesel = Gris (#64748B) — la bomba gris/negra para vehículos pesados")

    pdf.add_spacer(0.15)
    pdf.add_subsection("Consideraciones de accesibilidad")
    pdf.add_bullet("No depender únicamente de rojo vs verde (daltonismo)")
    pdf.add_bullet("Usar texturas o patrones además de color cuando sea posible")
    pdf.add_bullet("Asegurar contraste suficiente entre texto y fondo")
    pdf.add_bullet("Probar el gráfico en escala de grises para verificar legibilidad")

    # ── Sección 4: Ejercicios ────────────────────────────────────
    pdf.add_page_break()
    pdf.add_section("4. Ejercicios prácticos")

    pdf.add_text(
        "Completa estos ejercicios usando los archivos del Módulo 3. "
        "Cada ejercicio refuerza un concepto diferente de visualización."
    )
    pdf.add_spacer(0.2)

    # Ejercicio 1
    pdf.add_subsection("Ejercicio 1: Gráfico de combustible personalizado")
    pdf.add_text("<b>Archivo:</b> 07_Dashboard_Ventas_Combustible.xlsx")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("Abre la hoja 'Datos' y selecciona los meses de Enero a Junio")
    pdf.add_bullet("Crea un gráfico de líneas con marcadores para los litros de cada combustible")
    pdf.add_bullet("Aplica los colores correctos: Magna=verde, Premium=rojo, Diesel=gris")
    pdf.add_bullet("Agrega etiquetas de datos solo en los puntos máximo y mínimo")
    pdf.add_bullet("Coloca un título descriptivo y elimina la cuadrícula")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Qué aprendes:</b> A crear gráficos de líneas con colores significativos "
        "y limpieza visual profesional."
    )

    pdf.add_spacer(0.3)

    # Ejercicio 2
    pdf.add_subsection("Ejercicio 2: Estado de resultados con gráfico combinado")
    pdf.add_text("<b>Archivo:</b> 08_Comparativa_Anual_Ventas_Gastos.xlsx")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("En la hoja 'Estado_Resultados', selecciona Total Ingresos y Total Gastos para ambos años")
    pdf.add_bullet("Crea un gráfico de barras agrupadas que muestre los 4 valores")
    pdf.add_bullet("Usa azul para ingresos y rojo para gastos")
    pdf.add_bullet("Agrega una línea que muestre la Utilidad Bruta de cada año (eje secundario)")
    pdf.add_bullet("Formatea el eje Y en millones (ej: $80M)")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Qué aprendes:</b> A crear gráficos combinados (barras + línea) "
        "y usar ejes secundarios para mostrar diferentes escalas."
    )

    pdf.add_spacer(0.3)

    # Ejercicio 3
    pdf.add_subsection("Ejercicio 3: Dashboard básico con segmentadores")
    pdf.add_text("<b>Archivo:</b> Crear desde cero usando los datos de combustible")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("Copia la hoja 'Datos' del archivo 07 a un libro nuevo")
    pdf.add_bullet("Convierte los datos en Tabla (Ctrl+T)")
    pdf.add_bullet("Crea una Tabla Dinámica en una hoja nueva")
    pdf.add_bullet("Inserta un Gráfico Dinámico de columnas apiladas")
    pdf.add_bullet("Agrega un segmentador por Trimestre (agrupa los meses)")
    pdf.add_bullet("Aplica la paleta de colores del curso y limpia el gráfico")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Qué aprendes:</b> La combinación Tabla Dinámica + Gráfico Dinámico + "
        "Segmentador, que es la base del dashboard del Módulo 4."
    )

    # ── Sección 5: Fórmulas útiles para gráficos ─────────────────
    pdf.add_page_break()
    pdf.add_section("5. Fórmulas útiles para gráficos")

    pdf.add_text(
        "Estas fórmulas de Excel te ayudan a preparar datos para gráficos más efectivos."
    )
    pdf.add_spacer(0.15)

    formulas = [
        ["Fórmula", "Descripción", "Ejemplo"],
        ["SUMAR.SI", "Suma condicional para agrupar categorías", "=SUMAR.SI(A:A,\"Magna\",B:B)"],
        ["CONTAR.SI", "Cuenta registros por categoría", "=CONTAR.SI(A:A,\"Enero\")"],
        ["PROMEDIO.SI", "Promedio condicional", "=PROMEDIO.SI(A:A,\"Premium\",C:C)"],
        ["MAX / MIN", "Identifica picos y valles para etiquetas", "=MAX(B2:B13)"],
        ["TEXTO", "Formatea números para etiquetas personalizadas", '=TEXTO(B2,"$#,##0")'],
        ["REDONDEAR", "Simplifica cifras para presentación", "=REDONDEAR(B2/1000000,1)"],
    ]

    pdf.add_table(formulas, col_widths=[1.5 * inch, 2.2 * inch, 2.8 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Atajos de teclado para gráficos en Excel")

    atajos = [
        ["Atajo", "Acción"],
        ["Alt + F1", "Insertar gráfico rápido en la hoja actual"],
        ["F11", "Insertar gráfico en hoja nueva"],
        ["Ctrl + T", "Convertir rango en Tabla (base para gráficos dinámicos)"],
        ["Alt + N + C", "Abrir menú de insertar gráfico"],
        ["Ctrl + 1", "Abrir formato de elemento seleccionado"],
        ["Supr", "Eliminar elemento seleccionado del gráfico"],
    ]

    pdf.add_table(atajos, col_widths=[1.5 * inch, 5.0 * inch])

    pdf.save()
    return pdf.filepath


if __name__ == "__main__":
    build()
