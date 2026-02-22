"""
Generador: Referencia_Modulo_3.pdf
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
        title="Referencia Modulo 3 - Visualizacion de Impacto",
    )

    # ── Portada ───────────────────────────────────────────────────
    pdf.add_cover(
        title="Guia de Referencia",
        subtitle="Visualizacion de datos, seleccion de graficos y mejores practicas",
        modulo="Modulo 3: {}".format(MOD["nombre"]),
    )

    # ── Seccion 1: Guia de seleccion de graficos ─────────────────
    pdf.add_section("1. Guia de seleccion de graficos")

    pdf.add_text(
        "Antes de crear cualquier grafico, responde esta pregunta: "
        "<b>Que quiero comunicar?</b> La respuesta determina el tipo de grafico ideal."
    )
    pdf.add_spacer(0.15)

    pdf.add_subsection("Arbol de decision")
    pdf.add_text("Sigue este flujo para elegir el grafico correcto:")
    pdf.add_spacer(0.1)

    decision_tree = [
        ["Pregunta", "Respuesta", "Tipo de grafico"],
        ["Quiero comparar cantidades\nentre categorias?", "Si", "Barras (horizontales\no verticales)"],
        ["Quiero mostrar tendencia\nen el tiempo?", "Si", "Lineas"],
        ["Quiero mostrar proporcion\ndel total?", "Si, 5-6 categorias max", "Pastel o Dona"],
        ["Quiero composicion\na lo largo del tiempo?", "Si", "Columnas apiladas"],
        ["Quiero relacion entre\ndos variables?", "Si", "Dispersion (scatter)"],
        ["Quiero mostrar progreso\nhacia una meta?", "Si", "Indicador o barra\nde progreso"],
    ]

    pdf.add_table(decision_tree, col_widths=[2.5 * inch, 1.5 * inch, 2.5 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Errores comunes en la seleccion")
    pdf.add_bullet("Usar pastel con mas de 6 categorias (se vuelve ilegible)")
    pdf.add_bullet("Usar lineas para datos categoricos sin orden temporal")
    pdf.add_bullet("Usar efecto 3D (distorsiona las proporciones)")
    pdf.add_bullet("Usar doble eje Y sin justificacion clara")
    pdf.add_bullet("Usar graficos complejos cuando uno simple basta")

    # ── Seccion 2: Checklist de limpieza visual ──────────────────
    pdf.add_page_break()
    pdf.add_section("2. Checklist de limpieza visual")

    pdf.add_text(
        "Despues de crear tu grafico, revisa cada punto de esta lista. "
        "El objetivo es eliminar todo lo que no comunica informacion util."
    )
    pdf.add_spacer(0.15)

    checklist = [
        ["#", "Verificacion", "Accion"],
        ["1", "Titulo descriptivo", "Cambiar 'Grafico 1' por titulo que responda\n'de que es este grafico?'"],
        ["2", "Leyenda necesaria?", "Si solo hay 1 serie, eliminar la leyenda"],
        ["3", "Botones de campo", "En graficos dinamicos: click derecho >\nOcultar botones de campo"],
        ["4", "Lineas de cuadricula", "Eliminar si el grafico es simple\ny las etiquetas ya muestran valores"],
        ["5", "Etiquetas de datos", "Agregar valores sobre barras/puntos\nsi son pocos datos"],
        ["6", "Formato de ejes", "Usar K (miles), M (millones).\nQuitar decimales innecesarios"],
        ["7", "Colores", "Maximo 3-4 colores.\nConsistentes con el significado"],
        ["8", "Tamano de fuente", "Minimo 10pts para etiquetas,\n14pts para titulos"],
        ["9", "Orden de datos", "Ordenar barras de mayor a menor\n(o cronologico si aplica)"],
        ["10", "Fuente de datos", "Incluir periodo y unidad\n(ej: 'Ventas 2025, cifras en MDP')"],
    ]

    pdf.add_table(checklist, col_widths=[0.4 * inch, 1.8 * inch, 4.0 * inch])

    pdf.add_spacer(0.2)
    pdf.add_text(
        '<b>Principio de Edward Tufte:</b> "Maximiza la tinta de datos, '
        'minimiza la tinta de decoracion." Cada pixel debe tener un proposito.'
    )

    # ── Seccion 3: Paleta de colores ─────────────────────────────
    pdf.add_page_break()
    pdf.add_section("3. Paleta de colores de referencia")

    pdf.add_text(
        "Usa estos colores de forma consistente en todos tus reportes. "
        "Cada color tiene un significado intuitivo que facilita la lectura."
    )
    pdf.add_spacer(0.15)

    colores = [
        ["Color", "Codigo HEX", "Uso recomendado", "Ejemplo"],
        ["Azul", "#2563EB", "Color principal / institucional", "Titulos, barras principales"],
        ["Verde", "#10B981", "Positivo / crecimiento / Magna", "Utilidades, cumplimiento, Magna"],
        ["Rojo", "#EF4444", "Atencion / negativo / Premium", "Gastos excesivos, alertas, Premium"],
        ["Amarillo", "#F59E0B", "Precaucion / intermedio", "Advertencias, datos pendientes"],
        ["Gris", "#64748B", "Secundario / referencia / Diesel", "Datos de periodo anterior, Diesel"],
        ["Gris claro", "#CBD5E1", "Bordes / fondos", "Lineas de cuadricula, bordes de tabla"],
    ]

    pdf.add_table(colores, col_widths=[0.9 * inch, 1.0 * inch, 2.3 * inch, 2.3 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Ejemplo aplicado: Gasolinera")
    pdf.add_bullet("Magna = Verde (#10B981) -- la bomba verde que todos conocen")
    pdf.add_bullet("Premium = Rojo (#EF4444) -- la bomba roja de alto octanaje")
    pdf.add_bullet("Diesel = Gris (#64748B) -- la bomba gris/negra para vehiculos pesados")

    pdf.add_spacer(0.15)
    pdf.add_subsection("Consideraciones de accesibilidad")
    pdf.add_bullet("No depender unicamente de rojo vs verde (daltonismo)")
    pdf.add_bullet("Usar texturas o patrones ademas de color cuando sea posible")
    pdf.add_bullet("Asegurar contraste suficiente entre texto y fondo")
    pdf.add_bullet("Probar el grafico en escala de grises para verificar legibilidad")

    # ── Seccion 4: Ejercicios ────────────────────────────────────
    pdf.add_page_break()
    pdf.add_section("4. Ejercicios practicos")

    pdf.add_text(
        "Completa estos ejercicios usando los archivos del Modulo 3. "
        "Cada ejercicio refuerza un concepto diferente de visualizacion."
    )
    pdf.add_spacer(0.2)

    # Ejercicio 1
    pdf.add_subsection("Ejercicio 1: Grafico de combustible personalizado")
    pdf.add_text("<b>Archivo:</b> 07_Dashboard_Ventas_Combustible.xlsx")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("Abre la hoja 'Datos' y selecciona los meses de Enero a Junio")
    pdf.add_bullet("Crea un grafico de lineas con marcadores para los litros de cada combustible")
    pdf.add_bullet("Aplica los colores correctos: Magna=verde, Premium=rojo, Diesel=gris")
    pdf.add_bullet("Agrega etiquetas de datos solo en los puntos maximo y minimo")
    pdf.add_bullet("Coloca un titulo descriptivo y elimina la cuadricula")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Que aprendes:</b> A crear graficos de lineas con colores significativos "
        "y limpieza visual profesional."
    )

    pdf.add_spacer(0.3)

    # Ejercicio 2
    pdf.add_subsection("Ejercicio 2: Estado de resultados con grafico combinado")
    pdf.add_text("<b>Archivo:</b> 08_Comparativa_Anual_Ventas_Gastos.xlsx")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("En la hoja 'Estado_Resultados', selecciona Total Ingresos y Total Gastos para ambos anios")
    pdf.add_bullet("Crea un grafico de barras agrupadas que muestre los 4 valores")
    pdf.add_bullet("Usa azul para ingresos y rojo para gastos")
    pdf.add_bullet("Agrega una linea que muestre la Utilidad Bruta de cada anio (eje secundario)")
    pdf.add_bullet("Formatea el eje Y en millones (ej: $80M)")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Que aprendes:</b> A crear graficos combinados (barras + linea) "
        "y usar ejes secundarios para mostrar diferentes escalas."
    )

    pdf.add_spacer(0.3)

    # Ejercicio 3
    pdf.add_subsection("Ejercicio 3: Dashboard basico con segmentadores")
    pdf.add_text("<b>Archivo:</b> Crear desde cero usando los datos de combustible")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Instrucciones:</b>")
    pdf.add_bullet("Copia la hoja 'Datos' del archivo 07 a un libro nuevo")
    pdf.add_bullet("Convierte los datos en Tabla (Ctrl+T)")
    pdf.add_bullet("Crea una Tabla Dinamica en una hoja nueva")
    pdf.add_bullet("Inserta un Grafico Dinamico de columnas apiladas")
    pdf.add_bullet("Agrega un segmentador por Trimestre (agrupa los meses)")
    pdf.add_bullet("Aplica la paleta de colores del curso y limpia el grafico")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Que aprendes:</b> La combinacion Tabla Dinamica + Grafico Dinamico + "
        "Segmentador, que es la base del dashboard del Modulo 4."
    )

    # ── Seccion 5: Formulas utiles para graficos ─────────────────
    pdf.add_page_break()
    pdf.add_section("5. Formulas utiles para graficos")

    pdf.add_text(
        "Estas formulas de Excel te ayudan a preparar datos para graficos mas efectivos."
    )
    pdf.add_spacer(0.15)

    formulas = [
        ["Formula", "Descripcion", "Ejemplo"],
        ["SUMAR.SI", "Suma condicional para\nagrupar categorias", "=SUMAR.SI(A:A,\"Magna\",B:B)"],
        ["CONTAR.SI", "Cuenta registros por\ncategoria", "=CONTAR.SI(A:A,\"Enero\")"],
        ["PROMEDIO.SI", "Promedio condicional", "=PROMEDIO.SI(A:A,\"Premium\",C:C)"],
        ["MAX / MIN", "Identifica picos y valles\npara etiquetas", "=MAX(B2:B13)"],
        ["TEXTO", "Formatea numeros para\netiquetas personalizadas", '=TEXTO(B2,"$#,##0")'],
        ["REDONDEAR", "Simplifica cifras para\npresentacion", "=REDONDEAR(B2/1000000,1)"],
    ]

    pdf.add_table(formulas, col_widths=[1.5 * inch, 2.2 * inch, 2.8 * inch])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Atajos de teclado para graficos en Excel")

    atajos = [
        ["Atajo", "Accion"],
        ["Alt + F1", "Insertar grafico rapido en la hoja actual"],
        ["F11", "Insertar grafico en hoja nueva"],
        ["Ctrl + T", "Convertir rango en Tabla (base para graficos dinamicos)"],
        ["Alt + N + C", "Abrir menu de insertar grafico"],
        ["Ctrl + 1", "Abrir formato de elemento seleccionado"],
        ["Supr", "Eliminar elemento seleccionado del grafico"],
    ]

    pdf.add_table(atajos, col_widths=[1.5 * inch, 5.0 * inch])

    pdf.save()
    return pdf.filepath


if __name__ == "__main__":
    build()
