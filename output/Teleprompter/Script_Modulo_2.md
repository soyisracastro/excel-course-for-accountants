# Modulo 2: Procesamiento Masivo y Analisis con Tablas Dinamicas

## Slide 1 — Portada

Bienvenidos al Modulo 2. En este modulo vamos a transformar datos crudos y desordenados en analisis profesionales usando las herramientas mas poderosas de Excel: las Tablas y las Tablas Dinamicas. Si en el Modulo 1 aprendimos a pensar con logica contable y a usar funciones de control, ahora vamos a procesar volumenes grandes de informacion de forma eficiente.

## Slide 2 — Agenda del Modulo 2

Esta es nuestra agenda. Vamos a cubrir siete temas principales. Empezamos con la limpieza de datos, que es donde el 80% de los contadores pasan su tiempo. Luego entendemos la diferencia entre un rango y una tabla, que es fundamental. Despues creamos tablas dinamicas paso a paso, analizamos datos de nomina reales, y terminamos vinculando todo con un papel de trabajo para declaracion anual. Son aproximadamente 30 a 35 minutos de contenido practico.


## Slide 3 — El Problema: Datos Sucios

Antes de crear una tabla dinamica o cualquier reporte, necesitamos datos limpios. Y la realidad es que los datos nunca llegan limpios. Cuando descargas datos del SAT, de tu sistema contable, o de un ERP, siempre hay problemas. Celdas vacias, montos con signo de pesos que Excel interpreta como texto, fechas en formatos diferentes, RFCs con espacios... En el archivo 04 de su pack tienen 220 filas con errores intencionales para practicar. Les aseguro que si dominan la limpieza de datos, se ahorran horas de trabajo cada semana.


## Slide 4 — Herramientas de Limpieza en Excel

Estas son las herramientas que van a usar todos los dias. Ctrl+H es su mejor amigo: pueden quitar signos de pesos, comas, y cualquier caracter en segundos. La funcion ESPACIOS quita los espacios invisibles que causan que BUSCARV no encuentre resultados. VALOR convierte un texto como '15,000.00' en el numero 15000. Y Pegado Especial con Valores es clave: despues de limpiar con formulas, pegan solo los valores para eliminar la dependencia de la formula. En el archivo de practica van a aplicar cada una de estas herramientas.


## Slide 5 — Tabla vs Rango: La Diferencia Clave

Esta es una de las lecciones mas importantes del curso. La diferencia entre un rango y una tabla parece sutil, pero cambia completamente como trabajan en Excel. Un rango es simplemente un grupo de celdas. Una tabla es una estructura inteligente que tiene nombre, que mantiene los encabezados visibles, que expande las formulas automaticamente cuando agregas filas, y que permite usar referencias como Tabla[Columna] en vez de C2:C500. Ademas, las tablas dinamicas funcionan mucho mejor cuando su fuente es una tabla. La regla es simple: si tus datos tienen encabezados, seleccionalos y presiona Ctrl+T. Siempre.


## Slide 6 — Crear una Tabla en 3 Pasos

Crear una tabla es increiblemente facil. Solo necesitan tres pasos. Hagan clic en cualquier celda que tenga datos, presionen Ctrl+T, y confirmen que tiene encabezados. Listo. Ahora, el consejo mas importante: renombren la tabla inmediatamente. En la pestana Diseno de Tabla, arriba a la izquierda, cambien 'Tabla1' por un nombre descriptivo como 'nomina_2025' o 'gastos_marzo'. Esto hara que sus formulas y tablas dinamicas sean mucho mas legibles.


## Slide 7 — Tablas Dinamicas: El Superpoder de Excel

Si hay una sola herramienta que justifica este curso, son las tablas dinamicas. Con una tabla dinamica pueden tomar 500, 5,000 o 50,000 filas de datos y obtener un resumen en 10 segundos. Sin escribir una sola formula. Solo arrastrando campos. Quieren ver cuanto gasto cada departamento por mes? Arrastren 'Departamento' a filas, 'Mes' a columnas, y 'Monto' a valores. Listo. Quieren filtrar solo un trimestre? Arrastren 'Periodo' a filtros. Es asi de poderoso.


## Slide 8 — Crear una Tabla Dinamica

El proceso de creacion es directo. Asegurense de que su fuente de datos sea una tabla con nombre. Hagan clic en cualquier celda, Insertar, Tabla Dinamica. Recomiendo siempre crearla en una hoja nueva para mantener la organizacion. Una vez creada, ven un panel a la derecha con todos los campos de su tabla. Ese panel es su centro de control. Ahora vamos a entender las cuatro zonas donde van esos campos.


## Slide 9 — Las 4 Zonas de Campos

Las cuatro zonas son Filtros, Columnas, Filas y Valores. Piensen en Filas como su eje Y, lo que quieren analizar. Columnas es su eje X, como quieren desglosarlo. Valores es lo que quieren calcular: suma, promedio, conteo. Y Filtros es para restringir el analisis sin cambiar la estructura. Por ejemplo, para nuestro caso de nomina: ponemos Empleado en filas para ver cada persona, Concepto en columnas para ver sueldo, ISR, IMSS por separado, y Suma de Importe en valores. Si queremos ver solo un mes especifico, ponemos Periodo en filtros.


## Slide 10 — Caso Practico: Analisis de Nomina XML

Vamos a trabajar con el archivo 05 de su pack. Tiene mas de 500 filas que simulan los datos que obtendrian al extraer XMLs de nomina del SAT. Cada fila es un concepto de pago: sueldo, prima vacacional, ISR retenido, IMSS, subsidio al empleo. Tenemos 20 empleados con puestos reales, desde recepcionista hasta director administrativo, con 12 meses de datos. Los montos se calcularon con las tarifas reales del ISR. Vamos a crear tres tablas dinamicas para analizar percepciones, deducciones y el costo total por empleado.


## Slide 11 — Pivot 1: Percepciones por Empleado

La primera tabla dinamica es para percepciones. Filtren por Clase igual a Percepcion. Pongan NombreEmpleado en filas, Concepto en columnas, y Suma de ImporteGravado en valores. Van a ver inmediatamente quien gana mas sueldo, quien recibio prima vacacional, y el total por empleado. Ahora agreguen un segmentador por Periodo para poder filtrar por mes y ver como se comporta la nomina a lo largo del anio.


## Slide 12 — Pivot 2: Deducciones (ISR + IMSS)

La segunda tabla dinamica es para deducciones. Aqui van a ver algo muy interesante: la progresividad del ISR en accion. El director administrativo que gana 50,000 al mes paga alrededor de 30% de ISR, mientras que la recepcionista que gana el minimo paga apenas 6%. El IMSS, en cambio, es un porcentaje relativamente parejo para todos. Este tipo de analisis es exactamente lo que necesitan para revisar la nomina de sus clientes y detectar inconsistencias.


## Slide 13 — Pivot 3: Costo Total por Empleado

La tercera tabla dinamica es la vision ejecutiva. Sin filtrar por clase, pongan NombreEmpleado en filas y Clase en columnas. En valores pongan tanto ImporteGravado como ImporteExento. Esto les da el panorama completo: cuanto percibe cada empleado, cuanto se le retiene, y cuanto sale en otros pagos como subsidio. Con esta tabla dinamica pueden calcular el costo real de nomina y presentar un reporte ejecutivo a la direccion.


## Slide 14 — Segmentadores: Filtros Visuales

Los segmentadores son filtros visuales que se ven como botones. Son mucho mas intuitivos que los filtros del panel de campos. Y lo mejor: pueden conectar un solo segmentador a varias tablas dinamicas. Entonces si tienen tres pivots en la misma hoja, un clic en el segmentador de Periodo filtra las tres simultaneamente. Para nuestro caso de nomina, recomiendo crear un segmentador de Periodo para analizar mes a mes, y uno de Puesto para comparar por nivel jerarquico.


## Slide 15 — Papel de Trabajo Referenciado con Tarifa ISR

El cierre de este modulo es el papel de trabajo referenciado. Este archivo replica el calculo de la declaracion anual del ISR. Tiene la tarifa anual 2026 como tabla, y usa BUSCARV con IFERROR para calcular automaticamente el limite inferior, el excedente, el porcentaje marginal, la cuota fija y el ISR causado. Lo mas importante es la vinculacion: los totales que obtuvieron en sus tablas dinamicas de nomina se capturan aqui. Las percepciones van a ingresos, el ISR retenido va a retenciones, y el papel calcula si el contribuyente tiene ISR a cargo o a favor. Este flujo de trabajo es exactamente como lo hacen en un despacho real.


## Slide 16 — Tips Profesionales para Tablas Dinamicas

Para cerrar, aqui van siete tips que van a mejorar su trabajo diario. Primero, siempre nombren sus tablas. Segundo, recuerden actualizar los pivots porque NO se actualizan solos. Tercero, exploren la opcion Mostrar valores como, que les permite ver porcentajes del total, diferencias con el periodo anterior, y mas. Cuarto, si tienen fechas, Excel puede agruparlas automaticamente por mes, trimestre o anio. Y el tip mas importante: una tabla dinamica por pregunta. No intenten meter todo en un solo pivot, es mejor tener tres pivots claros que uno sobrecargado.


## Slide 17 — Cierre

Esto es todo para el Modulo 2. Ahora tienen las herramientas para limpiar datos, organizarlos en tablas, y analizarlos con tablas dinamicas. Practiquen con los tres archivos de su pack. En el Modulo 3 vamos a tomar estos analisis y convertirlos en graficos profesionales y reportes ejecutivos que impresionen a sus clientes y jefes. Nos vemos ahi.
