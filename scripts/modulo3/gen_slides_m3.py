"""
Generador: Modulo_3_Visualizacion.pptx + Script_Modulo_3.md
Modulo 3 -- Visualizacion de Impacto y Reportes Ejecutivos

15 slides sobre visualizacion de datos y graficos en Excel.
Script de teleprompter ~30 min en espanol.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import (
    SLIDES_DIR, TELEPROMPTER_DIR, MODULOS, CURSO_NOMBRE, INSTRUCTOR, ANIO
)
from scripts.generators.pptx_gen import SlidesGenerator

MOD = MODULOS[3]


def build():
    # type: () -> Path
    gen = SlidesGenerator(
        filename=MOD["slide_nombre"],
        output_dir=SLIDES_DIR,
        script_filename=MOD["script_nombre"],
        script_dir=TELEPROMPTER_DIR,
    )

    # ── Slide 1: Portada ──────────────────────────────────────────
    gen.add_title_slide(
        modulo_num=3,
        modulo_nombre=MOD["nombre"],
        subtitulo="Graficos que cuentan historias, no solo muestran numeros",
    )
    gen.script_lines.append(
        "Bienvenidos al Modulo 3. Hasta ahora hemos aprendido a procesar datos con funciones "
        "y a analizarlos con tablas dinamicas. Ahora vamos a dar el salto a la comunicacion visual. "
        "Un grafico bien hecho puede cambiar una decision de negocio en segundos. "
        "Vamos a aprender a crear visualizaciones que impacten.\n"
    )

    # ── Slide 2: Una imagen vale mas ─────────────────────────────
    gen.add_content_slide(
        title='"Una imagen vale mas que mil palabras" en contabilidad',
        bullets=[
            "Un director no lee 500 filas de datos -- lee un grafico de 5 segundos",
            "El cerebro procesa imagenes 60,000x mas rapido que texto (MIT)",
            "En juntas directivas, los reportes visuales generan decisiones mas rapidas",
            "Error comun: mostrar TODOS los datos en vez del mensaje clave",
            "Tu trabajo: convertir datos en una historia clara y accionable",
        ],
        script_text=(
            "Piensen en esto: cuando su jefe les pide un reporte, no quiere ver una hoja "
            "de Excel con 500 filas. Quiere saber: estamos bien o estamos mal? Subieron o "
            "bajaron las ventas? El cerebro humano procesa imagenes sesenta mil veces mas "
            "rapido que el texto -- eso lo dice un estudio del MIT. Entonces, si yo les "
            "puedo mostrar en un grafico de barras que las ventas de julio fueron las mas "
            "altas del anio, eso vale mas que cualquier tabla. Nuestro trabajo como contadores "
            "y administrativos no es solo capturar datos; es comunicarlos de forma que "
            "alguien pueda tomar una decision en cinco segundos.\n"
        ),
    )

    # ── Slide 3: Arbol de decision ────────────────────────────────
    gen.add_content_slide(
        title="Psicologia del grafico: arbol de decision",
        bullets=[
            "Comparar cantidades entre categorias --> Barras (horizontal o vertical)",
            "Mostrar tendencia en el tiempo --> Lineas",
            "Mostrar composicion / proporcion --> Pastel o dona (maximo 5-6 categorias)",
            "Comparar partes de un todo a lo largo del tiempo --> Columnas apiladas",
            "Mostrar relacion entre dos variables --> Dispersion (scatter)",
            "Regla de oro: si el grafico necesita explicacion, escogiste mal el tipo",
        ],
        script_text=(
            "Antes de crear cualquier grafico, necesitan hacerse una pregunta clave: "
            "Que quiero comunicar? Si quiero comparar las ventas de cinco vendedores, uso "
            "barras. Si quiero ver como se comportaron las ventas mes a mes, uso lineas. "
            "Si quiero ver que porcentaje del total representa cada producto, uso un pastel -- "
            "pero ojo, maximo cinco o seis rebanadas, si no se vuelve ilegible. Y aqui viene "
            "la regla de oro que quiero que se tatuen: si tu grafico necesita una explicacion "
            "larga para que alguien lo entienda, escogiste mal el tipo de grafico. El grafico "
            "correcto se explica solo.\n"
        ),
    )

    # ── Slide 4: Graficos dinamicos desde tablas dinamicas ────────
    gen.add_content_slide(
        title="Graficos dinamicos desde Tablas Dinamicas",
        bullets=[
            "Selecciona tu Tabla Dinamica -> Insertar -> Grafico Dinamico",
            "El grafico se actualiza automaticamente al filtrar la tabla",
            "Puedes usar segmentadores (slicers) para filtrar visualmente",
            "Ventaja: no necesitas reconstruir el grafico cuando cambian los datos",
            "Combinacion poderosa: Tabla Dinamica + Grafico Dinamico + Segmentador",
        ],
        script_text=(
            "En el modulo anterior construimos tablas dinamicas. Ahora, imaginen que esa tabla "
            "dinamica cobra vida visual. Solo seleccionan su tabla dinamica, van a Insertar, "
            "Grafico Dinamico, y listo. Lo increible es que si filtran la tabla, el grafico se "
            "actualiza solo. Y si le agregan un segmentador -- esos botones bonitos que parecen "
            "filtros visuales -- pueden hacer un mini-dashboard interactivo sin programar nada. "
            "Esta combinacion es la base de lo que veremos en el Modulo 4.\n"
        ),
    )

    # ── Slide 5: Caso ventas por vendedor (barras) ────────────────
    gen.add_content_slide(
        title="Caso practico: Ventas por vendedor (grafico de barras)",
        bullets=[
            "Escenario: 5 vendedores, ventas trimestrales",
            "Tabla Dinamica: Filas = Vendedor, Valores = Suma de Ventas",
            "Grafico de barras horizontales --> facil comparar nombres largos",
            "Ordenar de mayor a menor para impacto visual inmediato",
            "Agregar etiquetas de datos para mostrar cifras exactas",
            "Color unico para todas las barras (evitar arcoiris innecesario)",
        ],
        script_text=(
            "Veamos un caso concreto. Tenemos cinco vendedores y queremos saber quien vendio "
            "mas. Hacemos una tabla dinamica: en filas ponemos al vendedor, en valores la suma "
            "de ventas. Insertamos un grafico de barras. Un tip importante: si los nombres de "
            "los vendedores son largos, usen barras horizontales -- se leen mejor. Y ordenan "
            "de mayor a menor para que de un vistazo sepan quien es el campeon. Otro error "
            "comun es usar un color diferente para cada barra -- eso no agrega informacion, "
            "solo distrae. Usen un solo color, y si quieren destacar al top vendedor, cambien "
            "solo esa barra a un color mas fuerte.\n"
        ),
    )

    # ── Slide 6: Caso productos mas vendidos (pastel) ─────────────
    gen.add_content_slide(
        title="Caso practico: Productos mas vendidos (grafico de pastel)",
        bullets=[
            "Escenario: participacion porcentual de 5 productos",
            "Tabla Dinamica: Filas = Producto, Valores = Suma de Ventas",
            "Grafico de pastel / dona con porcentajes en etiquetas",
            "Maximo 5-6 categorias (agrupar el resto como 'Otros')",
            "Evitar efecto 3D -- distorsiona las proporciones visualmente",
            "Alternativa: barras apiladas al 100% si hay muchas categorias",
        ],
        script_text=(
            "Ahora veamos cuando si tiene sentido usar un pastel. Imaginen que quieren mostrar "
            "que porcentaje del total de ventas representa cada producto. Si son cinco productos, "
            "perfecto: el pastel queda limpio. Pero si tienen quince productos, agrupenlos: "
            "los top cinco por nombre y el resto como Otros. Nunca usen el efecto 3D -- se ve "
            "bonito pero engana al ojo; las rebanadas de atras parecen mas chicas de lo que son. "
            "Y un ultimo tip: si el pastel no se ve claro, cambien a barras apiladas al cien "
            "por ciento -- misma informacion, mas facil de leer.\n"
        ),
    )

    # ── Slide 7: Limpieza visual ─────────────────────────────────
    gen.add_content_slide(
        title="Limpieza visual: menos es mas",
        bullets=[
            "Ocultar botones de campo en graficos dinamicos",
            "Quitar leyendas redundantes (si solo hay una serie)",
            "Eliminar lineas de cuadricula si no aportan informacion",
            "Usar formato de eje: quitar decimales innecesarios",
            "Titulo claro y descriptivo (no 'Grafico 1')",
            "Alinear el grafico con los datos de la hoja",
        ],
        script_text=(
            "Ahora hablemos de algo que separa a un amateur de un profesional: la limpieza visual. "
            "Cuando insertan un grafico dinamico, Excel pone unos botones de campo que son "
            "utiles para explorar pero feos en una presentacion. Click derecho, Ocultar botones "
            "de campo. Si su grafico solo tiene una serie de datos, la leyenda sobra -- eliminenla. "
            "Las lineas de cuadricula? Si el grafico es simple, quitenlas. Cada pixel de su "
            "grafico debe tener un proposito. Como decia Edward Tufte, el padre de la visualizacion "
            "de datos: maximiza la tinta de datos, minimiza la tinta de decoracion.\n"
        ),
    )

    # ── Slide 8: Colores con sentido ─────────────────────────────
    gen.add_content_slide(
        title="Colores con sentido: no al arcoiris",
        bullets=[
            "Cada color debe tener un significado o proposito",
            "Ejemplo gasolinera: Magna = verde, Premium = rojo, Diesel = gris",
            "Verde suele significar positivo / crecimiento / aprobado",
            "Rojo suele significar atencion / negativo / rechazo",
            "Gris para datos secundarios o de referencia",
            "Maximo 3-4 colores por grafico -- consistencia entre reportes",
            "Considerar daltonismo: no depender solo de rojo vs verde",
        ],
        script_text=(
            "Los colores no son decoracion -- son informacion. Veamos un ejemplo real: "
            "en la gasolinera, Magna es verde porque es la bomba verde. Premium es roja. "
            "Diesel es gris o negro. Si yo hago un grafico y pongo Magna en azul, Premium en "
            "amarillo y Diesel en rosa, el dueño de la gasolinera no lo va a entender "
            "intuitivamente. Pero si uso los colores que ya conoce, la lectura es instantanea. "
            "La regla es: maximo tres o cuatro colores por grafico, y que sean consistentes "
            "en todos sus reportes. Si Magna es verde en enero, debe ser verde en diciembre.\n"
        ),
    )

    # ── Slide 9: Caso combustible (columnas apiladas) ─────────────
    gen.add_content_slide(
        title="Caso: Ventas de combustible (columnas apiladas mensuales)",
        bullets=[
            "Archivo: 07_Dashboard_Ventas_Combustible.xlsx",
            "Datos: 12 meses x 3 tipos de combustible (litros y monto)",
            "Grafico de columnas apiladas: cada columna = mes, cada color = tipo",
            "Se ve la composicion (que tipo vende mas) Y la tendencia (meses altos/bajos)",
            "Colores: Magna (#10B981), Premium (#EF4444), Diesel (#64748B)",
            "La altura total de la columna = total de ventas del mes",
        ],
        script_text=(
            "Abran el archivo 07 de Dashboard de Ventas de Combustible. Este es un caso real "
            "simplificado de una gasolinera. Tenemos doce meses, tres tipos de combustible, "
            "litros y montos. El grafico de columnas apiladas es perfecto aqui porque nos muestra "
            "dos cosas al mismo tiempo: primero, la composicion -- cuanto vende cada tipo; "
            "segundo, la tendencia -- cuales meses venden mas. Cada columna es un mes, y los "
            "colores son los que ya definimos: verde para Magna, rojo para Premium, gris para "
            "Diesel. La altura total de la columna nos dice el total de ventas de ese mes. "
            "Pueden ver de un vistazo que julio es el mes mas fuerte.\n"
        ),
    )

    # ── Slide 10: Lectura de estacionalidad ──────────────────────
    gen.add_content_slide(
        title="Lectura de estacionalidad en el grafico",
        bullets=[
            "Enero bajo: 'cuesta de enero' -- la gente gasta menos despues de diciembre",
            "Febrero-marzo: recuperacion gradual",
            "Abril-mayo: estabilizacion cercana al promedio",
            "Junio-agosto: pico de verano (vacaciones, viajes, mas consumo)",
            "Septiembre-noviembre: descenso gradual, regreso a clases, ahorro",
            "Diciembre: baja ligera (gastos navidenos en otras cosas)",
            "Estas tendencias ayudan a PLANEAR compras y flujo de efectivo",
        ],
        script_text=(
            "Ahora leamos el grafico como un contador. Vean enero: es el mes mas bajo. "
            "Le llaman la cuesta de enero porque la gente gasto todo en diciembre y en enero "
            "no tiene dinero. Pero conforme avanzan los meses, las ventas suben. El pico esta "
            "en julio -- vacaciones de verano, la gente viaja, se mueve, consume gasolina. "
            "Luego baja en el ultimo trimestre. Diciembre baja un poco porque el gasto se "
            "desvia a regalos y cenas. Esta informacion no es solo curiosidad -- si yo soy el "
            "administrador de la gasolinera, necesito comprar mas combustible en junio y puedo "
            "negociar mejores precios con el proveedor sabiendo que en enero la demanda baja. "
            "El grafico me ayuda a planear mi flujo de efectivo.\n"
        ),
    )

    # ── Slide 11: Comparativa anual ──────────────────────────────
    gen.add_content_slide(
        title="Comparativa anual: 2024 vs 2025",
        bullets=[
            "Archivo: 08_Comparativa_Anual_Ventas_Gastos.xlsx",
            "Comparar el mismo rubro en dos periodos es basico en contabilidad",
            "Barras lado a lado: gris para 2024 (pasado), color fuerte para 2025 (actual)",
            "Se responde: crecimos o nos encogimos?",
            "La variacion porcentual da contexto: 5 millones mas suena diferente si es 5% o 50%",
            "Siempre incluir ambos ejes: absoluto y porcentual",
        ],
        script_text=(
            "Pasemos al archivo 08 de Comparativa Anual. En contabilidad, comparar periodos "
            "es pan de cada dia. El estado de resultados del anio actual sin compararlo con "
            "el anterior no dice mucho. Este archivo muestra un estado de resultados simplificado: "
            "ingresos y gastos de 2024 versus 2025. El grafico usa barras lado a lado: gris para "
            "2024 porque es el pasado, azul fuerte para 2025 porque es el presente. De un vistazo "
            "ven que 2025 crecio. Pero ojo, no se queden solo con los pesos absolutos -- "
            "la variacion porcentual les dice si el crecimiento es bueno o mediocre.\n"
        ),
    )

    # ── Slide 12: Estado de resultados comparativo ────────────────
    gen.add_content_slide(
        title="Estado de Resultados comparativo visual",
        bullets=[
            "Ingresos: Ventas + Otros Ingresos = Total Ingresos",
            "Gastos: Compras + Gastos Generales + Nomina = Total Gastos",
            "Utilidad Bruta = Total Ingresos - Total Gastos",
            "Un grafico para ingresos (azul), otro para gastos (rojo)",
            "La separacion en dos graficos evita confusion visual",
            "Insight clave: si gastos crecen mas rapido que ingresos, hay problema",
        ],
        script_text=(
            "Veamos la estructura del estado de resultados. Arriba van los ingresos: ventas "
            "y otros ingresos. Abajo van los gastos: lo que compramos, los gastos generales "
            "y la nomina. La diferencia es la utilidad bruta. En el archivo tenemos dos graficos "
            "separados: uno para ingresos en azul y otro para gastos en rojo. Por que separarlos? "
            "Porque si los pongo juntos, las escalas son diferentes y se ve confuso. Con graficos "
            "separados, cada uno cuenta su historia. El insight mas importante que deben buscar: "
            "si los gastos crecieron mas rapido que los ingresos, la utilidad se esta comprimiendo "
            "aunque los numeros absolutos se vean bien. Ese es el tipo de analisis que hace un "
            "buen contador.\n"
        ),
    )

    # ── Slide 13: Datos y etiquetas profesionales ─────────────────
    gen.add_content_slide(
        title="Datos y etiquetas profesionales",
        bullets=[
            "Etiquetas de datos: mostrar el valor exacto sobre la barra",
            "Formato del eje: usar miles (K) o millones (M) para cifras grandes",
            "No repetir informacion: si esta en la etiqueta, no necesita estar en el eje",
            "Titulo del grafico: debe responder 'de que es este grafico?'",
            "Fuente de datos: incluir siempre el periodo y la unidad",
            "Tamano legible: minimo 10 pts para etiquetas, 14 pts para titulos",
        ],
        script_text=(
            "Los detalles marcan la diferencia. Cuando su grafico va en un reporte o presentacion, "
            "necesita ser autonomo -- alguien debe poder entenderlo sin que ustedes lo expliquen. "
            "Primero, las etiquetas de datos: pongan el numero exacto sobre cada barra si el "
            "grafico es simple. Segundo, si las cifras son en millones, no pongan 85,200,000 -- "
            "pongan 85.2M. Tercero, el titulo: nada de 'Grafico 1'. Pongan 'Ventas Mensuales "
            "de Combustible 2025 (Monto en Pesos)'. Cuarto, la fuente: digan de donde vienen "
            "los datos y que periodo cubren. Y ultimo, el tamano: si su jefe tiene que entrecerrar "
            "los ojos para leer la etiqueta, esta muy chica.\n"
        ),
    )

    # ── Slide 14: Resumen ─────────────────────────────────────────
    gen.add_content_slide(
        title="Resumen del Modulo 3",
        bullets=[
            "Elegir el tipo de grafico correcto es el 80% del trabajo",
            "Graficos dinamicos + segmentadores = analisis interactivo",
            "Limpieza visual: quitar todo lo que no comunica",
            "Colores con significado, no con decoracion",
            "Columnas apiladas para composicion + tendencia temporal",
            "Comparativas anuales: siempre mostrar variacion porcentual",
            "Cada grafico debe contar UNA historia clara",
        ],
        script_text=(
            "Recapitulemos lo que aprendimos. Primero, elegir bien el grafico es el ochenta "
            "por ciento del trabajo -- si el tipo es correcto, casi se explica solo. Segundo, "
            "la combinacion de tabla dinamica con grafico dinamico y segmentador es poderosa "
            "para analisis interactivo. Tercero, la limpieza visual es clave: si algo no "
            "comunica, quitenlo. Cuarto, los colores tienen significado -- no son decoracion. "
            "Quinto, las columnas apiladas son ideales cuando quieren ver composicion y tendencia "
            "al mismo tiempo. Sexto, en comparativas anuales siempre muestren el porcentaje "
            "de variacion, no solo los numeros absolutos. Y septimo, cada grafico debe contar "
            "una sola historia -- si trata de contar tres, no cuenta ninguna.\n"
        ),
    )

    # ── Slide 15: Cierre / Recursos ──────────────────────────────
    gen.add_closing_slide(
        next_module="Modulo 4 -- El Dashboard Inteligente y Entrega Profesional",
        resources=[
            "Archivo: 07_Dashboard_Ventas_Combustible.xlsx (columnas apiladas)",
            "Archivo: 08_Comparativa_Anual_Ventas_Gastos.xlsx (barras comparativas)",
            "PDF: Referencia_Modulo_3.pdf (guia de seleccion de graficos y ejercicios)",
            "Referencia: Edward Tufte - The Visual Display of Quantitative Information",
        ],
    )
    gen.script_lines.append(
        "En el siguiente modulo vamos a integrar todo lo que hemos aprendido en un "
        "dashboard interactivo con segmentadores, graficos dinamicos y un diseno profesional. "
        "Nos vemos en el Modulo 4. Practiquen con los archivos de esta sesion.\n"
    )

    gen.save()
    return gen.filepath


if __name__ == "__main__":
    build()
