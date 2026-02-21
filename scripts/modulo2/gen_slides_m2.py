"""
Generador: Modulo_2_Tablas_Dinamicas.pptx + Script_Modulo_2.md
Modulo 2 -- Procesamiento Masivo y Analisis con Tablas Dinamicas

15-17 slides cubriendo:
  - Limpieza de datos masiva
  - Tablas vs Rangos
  - Tablas Dinamicas (pivot tables)
  - Papel de trabajo referenciado

Teleprompter: ~30-35 min, espanol, tono profesional pero accesible.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import (
    PACK, SLIDES_DIR, TELEPROMPTER_DIR, MODULOS
)
from scripts.generators.pptx_gen import SlidesGenerator

M2 = MODULOS[2]


def build():
    gen = SlidesGenerator(
        filename=M2["slide_nombre"],
        output_dir=SLIDES_DIR,
        script_filename=M2["script_nombre"],
        script_dir=TELEPROMPTER_DIR,
    )

    # ================================================================
    # SLIDE 1 -- Portada
    # ================================================================
    gen.add_title_slide(
        modulo_num=2,
        modulo_nombre=M2["nombre"],
        subtitulo="De datos crudos a analisis profesional en minutos"
    )
    gen.script_lines.append(
        "Bienvenidos al Modulo 2. En este modulo vamos a transformar datos crudos "
        "y desordenados en analisis profesionales usando las herramientas mas poderosas "
        "de Excel: las Tablas y las Tablas Dinamicas. Si en el Modulo 1 aprendimos a "
        "pensar con logica contable y a usar funciones de control, ahora vamos a procesar "
        "volumenes grandes de informacion de forma eficiente.\n"
    )

    # ================================================================
    # SLIDE 2 -- Agenda del modulo
    # ================================================================
    gen.add_content_slide(
        title="Agenda del Modulo 2",
        bullets=[
            "Limpieza masiva de datos: tecnicas y herramientas",
            "Tabla vs Rango: por que SIEMPRE usar Tablas",
            "Creacion de Tablas Dinamicas paso a paso",
            "Zonas de campo: Filas, Columnas, Valores, Filtros",
            "Caso practico: Analisis de nomina desde XML",
            "Papel de trabajo referenciado con BUSCARV",
            "Segmentadores (Slicers) para analisis interactivo",
        ],
        script_text=(
            "Esta es nuestra agenda. Vamos a cubrir siete temas principales. "
            "Empezamos con la limpieza de datos, que es donde el 80% de los contadores "
            "pasan su tiempo. Luego entendemos la diferencia entre un rango y una tabla, "
            "que es fundamental. Despues creamos tablas dinamicas paso a paso, "
            "analizamos datos de nomina reales, y terminamos vinculando todo con un "
            "papel de trabajo para declaracion anual. Son aproximadamente 30 a 35 minutos "
            "de contenido practico.\n"
        )
    )

    # ================================================================
    # SLIDE 3 -- El problema: datos sucios
    # ================================================================
    gen.add_content_slide(
        title="El Problema: Datos Sucios",
        bullets=[
            "El 80% del tiempo de analisis se gasta en LIMPIAR datos",
            "Problemas comunes: celdas vacias, espacios extra, formatos mixtos",
            "Texto almacenado como numero (subtotales con '$' y comas)",
            "Fechas en 5 formatos diferentes en la misma columna",
            "RFCs incompletos o con espacios al inicio/final",
            "Filas duplicadas que distorsionan los totales",
            "Sin limpieza, CUALQUIER analisis estara MAL",
        ],
        script_text=(
            "Antes de crear una tabla dinamica o cualquier reporte, necesitamos datos "
            "limpios. Y la realidad es que los datos nunca llegan limpios. "
            "Cuando descargas datos del SAT, de tu sistema contable, o de un ERP, "
            "siempre hay problemas. Celdas vacias, montos con signo de pesos que Excel "
            "interpreta como texto, fechas en formatos diferentes, RFCs con espacios... "
            "En el archivo 04 de su pack tienen 220 filas con errores intencionales "
            "para practicar. Les aseguro que si dominan la limpieza de datos, "
            "se ahorran horas de trabajo cada semana.\n"
        )
    )

    # ================================================================
    # SLIDE 4 -- Herramientas de limpieza
    # ================================================================
    gen.add_content_slide(
        title="Herramientas de Limpieza en Excel",
        bullets=[
            "Buscar y Reemplazar (Ctrl+H): quitar '$', comas, espacios",
            "ESPACIOS (TRIM): elimina espacios al inicio, final y dobles",
            "LIMPIAR (CLEAN): elimina caracteres no imprimibles",
            "SUSTITUIR (SUBSTITUTE): reemplaza texto especifico",
            "VALOR (VALUE): convierte texto que parece numero a numero real",
            "Texto en Columnas: separa datos concatenados",
            "Quitar Duplicados: elimina filas repetidas (Datos > Quitar duplicados)",
            "Pegado Especial > Valores: 'aplasta' formulas a valores fijos",
        ],
        script_text=(
            "Estas son las herramientas que van a usar todos los dias. "
            "Ctrl+H es su mejor amigo: pueden quitar signos de pesos, comas, "
            "y cualquier caracter en segundos. La funcion ESPACIOS quita los espacios "
            "invisibles que causan que BUSCARV no encuentre resultados. "
            "VALOR convierte un texto como '15,000.00' en el numero 15000. "
            "Y Pegado Especial con Valores es clave: despues de limpiar con formulas, "
            "pegan solo los valores para eliminar la dependencia de la formula. "
            "En el archivo de practica van a aplicar cada una de estas herramientas.\n"
        )
    )

    # ================================================================
    # SLIDE 5 -- Tabla vs Rango
    # ================================================================
    gen.add_content_slide(
        title="Tabla vs Rango: La Diferencia Clave",
        bullets=[
            "RANGO: celdas sueltas sin estructura (A1:H200)",
            "TABLA: estructura inteligente con nombre (Ctrl+T)",
            "Tabla agrega encabezados fijos al hacer scroll",
            "Tabla expande formulas automaticamente al agregar filas",
            "Tabla permite referencias estructuradas: =SUMA(Tabla[Ventas])",
            "Tabla es REQUISITO para Tablas Dinamicas eficientes",
            "Regla de oro: Si tiene encabezados, CONVIERTELO en Tabla",
        ],
        script_text=(
            "Esta es una de las lecciones mas importantes del curso. "
            "La diferencia entre un rango y una tabla parece sutil, pero cambia "
            "completamente como trabajan en Excel. Un rango es simplemente un grupo "
            "de celdas. Una tabla es una estructura inteligente que tiene nombre, "
            "que mantiene los encabezados visibles, que expande las formulas "
            "automaticamente cuando agregas filas, y que permite usar referencias "
            "como Tabla[Columna] en vez de C2:C500. Ademas, las tablas dinamicas "
            "funcionan mucho mejor cuando su fuente es una tabla. "
            "La regla es simple: si tus datos tienen encabezados, seleccionalos "
            "y presiona Ctrl+T. Siempre.\n"
        )
    )

    # ================================================================
    # SLIDE 6 -- Crear una tabla
    # ================================================================
    gen.add_content_slide(
        title="Crear una Tabla en 3 Pasos",
        bullets=[
            "1. Selecciona cualquier celda dentro de tus datos",
            "2. Presiona Ctrl+T (o Insertar > Tabla)",
            "3. Verifica que 'La tabla tiene encabezados' este marcado",
            "",
            "Excel detecta automaticamente el rango de datos",
            "Asigna un nombre por defecto (Tabla1, Tabla2...)",
            "CONSEJO: Renombra la tabla inmediatamente (pestana Diseno de Tabla)",
            "Nombres descriptivos: nomina_2025, gastos_deducibles, catalogo_cuentas",
        ],
        script_text=(
            "Crear una tabla es increiblemente facil. Solo necesitan tres pasos. "
            "Hagan clic en cualquier celda que tenga datos, presionen Ctrl+T, "
            "y confirmen que tiene encabezados. Listo. Ahora, el consejo mas "
            "importante: renombren la tabla inmediatamente. En la pestana Diseno "
            "de Tabla, arriba a la izquierda, cambien 'Tabla1' por un nombre "
            "descriptivo como 'nomina_2025' o 'gastos_marzo'. Esto hara que sus "
            "formulas y tablas dinamicas sean mucho mas legibles.\n"
        )
    )

    # ================================================================
    # SLIDE 7 -- Intro Tablas Dinamicas
    # ================================================================
    gen.add_content_slide(
        title="Tablas Dinamicas: El Superpoder de Excel",
        bullets=[
            "Resumen miles de filas en segundos",
            "Reorganiza datos arrastrando campos con el mouse",
            "Calcula sumas, promedios, conteos sin formulas",
            "Permite analisis multidimensional (por empleado, por mes, por concepto)",
            "Se actualizan con un clic derecho > Actualizar",
            "Son LA herramienta de analisis #1 para contadores",
        ],
        script_text=(
            "Si hay una sola herramienta que justifica este curso, son las tablas "
            "dinamicas. Con una tabla dinamica pueden tomar 500, 5,000 o 50,000 "
            "filas de datos y obtener un resumen en 10 segundos. Sin escribir "
            "una sola formula. Solo arrastrando campos. Quieren ver cuanto gasto "
            "cada departamento por mes? Arrastren 'Departamento' a filas, "
            "'Mes' a columnas, y 'Monto' a valores. Listo. Quieren filtrar solo "
            "un trimestre? Arrastren 'Periodo' a filtros. Es asi de poderoso.\n"
        )
    )

    # ================================================================
    # SLIDE 8 -- Crear Tabla Dinamica
    # ================================================================
    gen.add_content_slide(
        title="Crear una Tabla Dinamica",
        bullets=[
            "1. Haz clic en cualquier celda de tu tabla de datos",
            "2. Ve a Insertar > Tabla Dinamica",
            "3. Elige ubicacion: Nueva hoja (recomendado) o hoja existente",
            "4. Aparece el panel 'Campos de tabla dinamica' a la derecha",
            "5. Arrastra campos a las 4 zonas: Filtros, Columnas, Filas, Valores",
            "6. Excel calcula automaticamente los totales",
        ],
        script_text=(
            "El proceso de creacion es directo. Asegurense de que su fuente de datos "
            "sea una tabla con nombre. Hagan clic en cualquier celda, Insertar, "
            "Tabla Dinamica. Recomiendo siempre crearla en una hoja nueva para "
            "mantener la organizacion. Una vez creada, ven un panel a la derecha "
            "con todos los campos de su tabla. Ese panel es su centro de control. "
            "Ahora vamos a entender las cuatro zonas donde van esos campos.\n"
        )
    )

    # ================================================================
    # SLIDE 9 -- Las 4 zonas
    # ================================================================
    gen.add_content_slide(
        title="Las 4 Zonas de Campos",
        bullets=[
            "FILTROS: campos para filtrar toda la tabla (ej. Anio, Sucursal)",
            "COLUMNAS: campos que se muestran como encabezados horizontales",
            "FILAS: campos que se muestran como etiquetas verticales",
            "VALORES: campos numericos que se calculan (Suma, Promedio, Cuenta)",
            "",
            "Ejemplo nomina:",
            "  Filtros: Periodo | Filas: Empleado | Columnas: Concepto | Valores: Suma de Importe",
        ],
        script_text=(
            "Las cuatro zonas son Filtros, Columnas, Filas y Valores. "
            "Piensen en Filas como su eje Y, lo que quieren analizar. "
            "Columnas es su eje X, como quieren desglosarlo. "
            "Valores es lo que quieren calcular: suma, promedio, conteo. "
            "Y Filtros es para restringir el analisis sin cambiar la estructura. "
            "Por ejemplo, para nuestro caso de nomina: ponemos Empleado en filas "
            "para ver cada persona, Concepto en columnas para ver sueldo, ISR, IMSS "
            "por separado, y Suma de Importe en valores. Si queremos ver solo "
            "un mes especifico, ponemos Periodo en filtros.\n"
        )
    )

    # ================================================================
    # SLIDE 10 -- Caso practico: Nomina XML
    # ================================================================
    gen.add_content_slide(
        title="Caso Practico: Analisis de Nomina XML",
        bullets=[
            "Archivo: 05_Analisis_Nomina_XML_Pivot.xlsx",
            "500+ registros de 20 empleados, 12 meses",
            "Datos simulan extraccion de XML CFDI de nomina del SAT",
            "Columnas: UUID, Empleado, Puesto, Periodo, Clase, Concepto, Importes",
            "Clase: Percepcion, Deduccion, OtroPago",
            "",
            "Objetivo: Crear 3 tablas dinamicas para analizar la nomina completa",
        ],
        script_text=(
            "Vamos a trabajar con el archivo 05 de su pack. Tiene mas de 500 filas "
            "que simulan los datos que obtendrian al extraer XMLs de nomina del SAT. "
            "Cada fila es un concepto de pago: sueldo, prima vacacional, ISR retenido, "
            "IMSS, subsidio al empleo. Tenemos 20 empleados con puestos reales, "
            "desde recepcionista hasta director administrativo, con 12 meses de datos. "
            "Los montos se calcularon con las tarifas reales del ISR. "
            "Vamos a crear tres tablas dinamicas para analizar percepciones, "
            "deducciones y el costo total por empleado.\n"
        )
    )

    # ================================================================
    # SLIDE 11 -- Pivot 1: Percepciones
    # ================================================================
    gen.add_content_slide(
        title="Pivot 1: Percepciones por Empleado",
        bullets=[
            "Filtro de Clase: 'Percepcion'",
            "Filas: NombreEmpleado",
            "Columnas: Concepto (Sueldo, Prima Vacacional)",
            "Valores: Suma de ImporteGravado",
            "",
            "Preguntas que responde:",
            "  Quien gana mas? Cuanto se pago en prima vacacional?",
            "  Cual es el costo total de percepciones por puesto?",
        ],
        script_text=(
            "La primera tabla dinamica es para percepciones. Filtren por Clase igual "
            "a Percepcion. Pongan NombreEmpleado en filas, Concepto en columnas, "
            "y Suma de ImporteGravado en valores. Van a ver inmediatamente quien "
            "gana mas sueldo, quien recibio prima vacacional, y el total por empleado. "
            "Ahora agreguen un segmentador por Periodo para poder filtrar por mes "
            "y ver como se comporta la nomina a lo largo del anio.\n"
        )
    )

    # ================================================================
    # SLIDE 12 -- Pivot 2: Deducciones
    # ================================================================
    gen.add_content_slide(
        title="Pivot 2: Deducciones (ISR + IMSS)",
        bullets=[
            "Filtro de Clase: 'Deduccion'",
            "Filas: NombreEmpleado + Puesto (jerarquia)",
            "Columnas: Concepto (ISR, IMSS)",
            "Valores: Suma de ImporteGravado",
            "",
            "Insight clave: El ISR es proporcional al sueldo",
            "Director paga ~30% de ISR vs Recepcionista ~6%",
            "IMSS es relativamente parejo (~2.77% para todos)",
        ],
        script_text=(
            "La segunda tabla dinamica es para deducciones. Aqui van a ver algo "
            "muy interesante: la progresividad del ISR en accion. El director "
            "administrativo que gana 50,000 al mes paga alrededor de 30% de ISR, "
            "mientras que la recepcionista que gana el minimo paga apenas 6%. "
            "El IMSS, en cambio, es un porcentaje relativamente parejo para todos. "
            "Este tipo de analisis es exactamente lo que necesitan para revisar "
            "la nomina de sus clientes y detectar inconsistencias.\n"
        )
    )

    # ================================================================
    # SLIDE 13 -- Pivot 3: Costo total
    # ================================================================
    gen.add_content_slide(
        title="Pivot 3: Costo Total por Empleado",
        bullets=[
            "Sin filtro de Clase (todas las clases)",
            "Filas: NombreEmpleado",
            "Columnas: Clase (Percepcion, Deduccion, OtroPago)",
            "Valores: Suma de ImporteGravado + ImporteExento",
            "",
            "Costo neto = Percepciones - Deducciones + OtrosPagos",
            "Ideal para presupuestos y analisis de costo de nomina",
        ],
        script_text=(
            "La tercera tabla dinamica es la vision ejecutiva. Sin filtrar por clase, "
            "pongan NombreEmpleado en filas y Clase en columnas. "
            "En valores pongan tanto ImporteGravado como ImporteExento. "
            "Esto les da el panorama completo: cuanto percibe cada empleado, "
            "cuanto se le retiene, y cuanto sale en otros pagos como subsidio. "
            "Con esta tabla dinamica pueden calcular el costo real de nomina "
            "y presentar un reporte ejecutivo a la direccion.\n"
        )
    )

    # ================================================================
    # SLIDE 14 -- Segmentadores (Slicers)
    # ================================================================
    gen.add_content_slide(
        title="Segmentadores: Filtros Visuales",
        bullets=[
            "Insertar > Segmentacion de datos (Slicer)",
            "Botones visuales para filtrar la tabla dinamica",
            "Pueden conectar UN slicer a MULTIPLES tablas dinamicas",
            "Ideal para dashboards interactivos",
            "",
            "Recomendacion para nomina:",
            "  Slicer de Periodo (filtrar por mes)",
            "  Slicer de Puesto (filtrar por cargo)",
            "  Slicer de Clase (cambiar entre percepciones/deducciones)",
        ],
        script_text=(
            "Los segmentadores son filtros visuales que se ven como botones. "
            "Son mucho mas intuitivos que los filtros del panel de campos. "
            "Y lo mejor: pueden conectar un solo segmentador a varias tablas "
            "dinamicas. Entonces si tienen tres pivots en la misma hoja, "
            "un clic en el segmentador de Periodo filtra las tres simultaneamente. "
            "Para nuestro caso de nomina, recomiendo crear un segmentador de "
            "Periodo para analizar mes a mes, y uno de Puesto para comparar "
            "por nivel jerarquico.\n"
        )
    )

    # ================================================================
    # SLIDE 15 -- Papel de trabajo referenciado
    # ================================================================
    gen.add_content_slide(
        title="Papel de Trabajo Referenciado con Tarifa ISR",
        bullets=[
            "Archivo: 06_Papel_Trabajo_Referenciado.xlsx",
            "Formato de declaracion anual ISR con formulas automaticas",
            "Usa IFERROR(BUSCARV(...)) para buscar en tarifa ISR 2026",
            "Secciones: Ingresos, Deducciones, Base Gravable, Calculo ISR",
            "Celdas amarillas = captura manual, Verdes = calculo automatico",
            "",
            "VINCULACION: Los totales de tus Pivots alimentan este papel",
            "Percepciones -> Ingresos | ISR retenido -> Retenciones",
        ],
        script_text=(
            "El cierre de este modulo es el papel de trabajo referenciado. "
            "Este archivo replica el calculo de la declaracion anual del ISR. "
            "Tiene la tarifa anual 2026 como tabla, y usa BUSCARV con IFERROR "
            "para calcular automaticamente el limite inferior, el excedente, "
            "el porcentaje marginal, la cuota fija y el ISR causado. "
            "Lo mas importante es la vinculacion: los totales que obtuvieron "
            "en sus tablas dinamicas de nomina se capturan aqui. "
            "Las percepciones van a ingresos, el ISR retenido va a retenciones, "
            "y el papel calcula si el contribuyente tiene ISR a cargo o a favor. "
            "Este flujo de trabajo es exactamente como lo hacen en un despacho real.\n"
        )
    )

    # ================================================================
    # SLIDE 16 -- Tips profesionales
    # ================================================================
    gen.add_content_slide(
        title="Tips Profesionales para Tablas Dinamicas",
        bullets=[
            "SIEMPRE nombra tus tablas fuente (no dejes 'Tabla1')",
            "Actualiza pivots despues de cambiar datos: clic derecho > Actualizar",
            "Usa 'Mostrar valores como' para porcentajes y diferencias",
            "Agrupa fechas por mes/trimestre/anio automaticamente",
            "Crea campos calculados para metricas personalizadas",
            "Formato: ajusta numeros a moneda/porcentaje dentro del pivot",
            "Un pivot por pregunta: no sobrecargues una sola tabla",
        ],
        script_text=(
            "Para cerrar, aqui van siete tips que van a mejorar su trabajo diario. "
            "Primero, siempre nombren sus tablas. Segundo, recuerden actualizar "
            "los pivots porque NO se actualizan solos. Tercero, exploren la opcion "
            "Mostrar valores como, que les permite ver porcentajes del total, "
            "diferencias con el periodo anterior, y mas. Cuarto, si tienen fechas, "
            "Excel puede agruparlas automaticamente por mes, trimestre o anio. "
            "Y el tip mas importante: una tabla dinamica por pregunta. "
            "No intenten meter todo en un solo pivot, es mejor tener tres pivots "
            "claros que uno sobrecargado.\n"
        )
    )

    # ================================================================
    # SLIDE 17 -- Cierre
    # ================================================================
    gen.add_closing_slide(
        next_module="Modulo 3 - Visualizacion de Impacto y Reportes Ejecutivos",
        resources=[
            "Archivo 04: Practica de limpieza masiva (220 filas con errores)",
            "Archivo 05: Datos de nomina para crear 3 tablas dinamicas",
            "Archivo 06: Papel de trabajo ISR con BUSCARV y tarifa 2026",
            "PDF de Referencia: Modulo 2 (checklist, diagramas, ejercicios)",
        ]
    )
    gen.script_lines.append(
        "Esto es todo para el Modulo 2. Ahora tienen las herramientas para limpiar "
        "datos, organizarlos en tablas, y analizarlos con tablas dinamicas. "
        "Practiquen con los tres archivos de su pack. En el Modulo 3 vamos a "
        "tomar estos analisis y convertirlos en graficos profesionales y reportes "
        "ejecutivos que impresionen a sus clientes y jefes. Nos vemos ahi.\n"
    )

    gen.save()


if __name__ == "__main__":
    build()
