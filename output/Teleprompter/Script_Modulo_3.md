# Módulo 3: Visualización de Impacto y Reportes Ejecutivos

## Slide 1 — Portada

Bienvenidos al Modulo 3. Hasta ahora hemos aprendido a procesar datos con funciones y a analizarlos con tablas dinamicas. Ahora vamos a dar el salto a la comunicacion visual. Un grafico bien hecho puede cambiar una decision de negocio en segundos. Vamos a aprender a crear visualizaciones que impacten.

## Slide 2 — "Una imagen vale mas que mil palabras" en contabilidad

Piensen en esto: cuando su jefe les pide un reporte, no quiere ver una hoja de Excel con 500 filas. Quiere saber: estamos bien o estamos mal? Subieron o bajaron las ventas? El cerebro humano procesa imagenes sesenta mil veces mas rapido que el texto -- eso lo dice un estudio del MIT. Entonces, si yo les puedo mostrar en un grafico de barras que las ventas de julio fueron las mas altas del anio, eso vale mas que cualquier tabla. Nuestro trabajo como contadores y administrativos no es solo capturar datos; es comunicarlos de forma que alguien pueda tomar una decision en cinco segundos.


## Slide 3 — Psicologia del grafico: arbol de decision

Antes de crear cualquier grafico, necesitan hacerse una pregunta clave: Que quiero comunicar? Si quiero comparar las ventas de cinco vendedores, uso barras. Si quiero ver como se comportaron las ventas mes a mes, uso lineas. Si quiero ver que porcentaje del total representa cada producto, uso un pastel -- pero ojo, maximo cinco o seis rebanadas, si no se vuelve ilegible. Y aqui viene la regla de oro que quiero que se tatuen: si tu grafico necesita una explicacion larga para que alguien lo entienda, escogiste mal el tipo de grafico. El grafico correcto se explica solo.


## Slide 4 — Graficos dinamicos desde Tablas Dinamicas

En el modulo anterior construimos tablas dinamicas. Ahora, imaginen que esa tabla dinamica cobra vida visual. Solo seleccionan su tabla dinamica, van a Insertar, Grafico Dinamico, y listo. Lo increible es que si filtran la tabla, el grafico se actualiza solo. Y si le agregan un segmentador -- esos botones bonitos que parecen filtros visuales -- pueden hacer un mini-dashboard interactivo sin programar nada. Esta combinacion es la base de lo que veremos en el Modulo 4.


## Slide 5 — Caso practico: Ventas por vendedor (grafico de barras)

Veamos un caso concreto. Tenemos cinco vendedores y queremos saber quien vendio mas. Hacemos una tabla dinamica: en filas ponemos al vendedor, en valores la suma de ventas. Insertamos un grafico de barras. Un tip importante: si los nombres de los vendedores son largos, usen barras horizontales -- se leen mejor. Y ordenan de mayor a menor para que de un vistazo sepan quien es el campeon. Otro error comun es usar un color diferente para cada barra -- eso no agrega informacion, solo distrae. Usen un solo color, y si quieren destacar al top vendedor, cambien solo esa barra a un color mas fuerte.


## Slide 6 — Caso practico: Productos mas vendidos (grafico de pastel)

Ahora veamos cuando si tiene sentido usar un pastel. Imaginen que quieren mostrar que porcentaje del total de ventas representa cada producto. Si son cinco productos, perfecto: el pastel queda limpio. Pero si tienen quince productos, agrúpenlos: los top cinco por nombre y el resto como Otros. Nunca usen el efecto 3D -- se ve bonito pero engana al ojo; las rebanadas de atras parecen mas chicas de lo que son. Y un ultimo tip: si el pastel no se ve claro, cambien a barras apiladas al cien por ciento -- misma informacion, mas facil de leer.


## Slide 7 — Limpieza visual: menos es mas

Ahora hablemos de algo que separa a un amateur de un profesional: la limpieza visual. Cuando insertan un grafico dinamico, Excel pone unos botones de campo que son utiles para explorar pero feos en una presentacion. Click derecho, Ocultar botones de campo. Si su grafico solo tiene una serie de datos, la leyenda sobra -- eliminenla. Las lineas de cuadricula? Si el grafico es simple, quitenlas. Cada pixel de su grafico debe tener un proposito. Como decia Edward Tufte, el padre de la visualizacion de datos: maximiza la tinta de datos, minimiza la tinta de decoracion.


## Slide 8 — Colores con sentido: no al arcoiris

Los colores no son decoracion -- son informacion. Veamos un ejemplo real: en la gasolinera, Magna es verde porque es la bomba verde. Premium es roja. Diesel es gris o negro. Si yo hago un grafico y pongo Magna en azul, Premium en amarillo y Diesel en rosa, el dueño de la gasolinera no lo va a entender intuitivamente. Pero si uso los colores que ya conoce, la lectura es instantanea. La regla es: maximo tres o cuatro colores por grafico, y que sean consistentes en todos sus reportes. Si Magna es verde en enero, debe ser verde en diciembre.


## Slide 9 — Caso: Ventas de combustible (columnas apiladas mensuales)

Abran el archivo 07 de Dashboard de Ventas de Combustible. Este es un caso real simplificado de una gasolinera. Tenemos doce meses, tres tipos de combustible, litros y montos. El grafico de columnas apiladas es perfecto aqui porque nos muestra dos cosas al mismo tiempo: primero, la composicion -- cuanto vende cada tipo; segundo, la tendencia -- cuales meses venden mas. Cada columna es un mes, y los colores son los que ya definimos: verde para Magna, rojo para Premium, gris para Diesel. La altura total de la columna nos dice el total de ventas de ese mes. Pueden ver de un vistazo que julio es el mes mas fuerte.


## Slide 10 — Lectura de estacionalidad en el grafico

Ahora leamos el grafico como un contador. Vean enero: es el mes mas bajo. Le llaman la cuesta de enero porque la gente gasto todo en diciembre y en enero no tiene dinero. Pero conforme avanzan los meses, las ventas suben. El pico esta en julio -- vacaciones de verano, la gente viaja, se mueve, consume gasolina. Luego baja en el ultimo trimestre. Diciembre baja un poco porque el gasto se desvía a regalos y cenas. Esta informacion no es solo curiosidad -- si yo soy el administrador de la gasolinera, necesito comprar mas combustible en junio y puedo negociar mejores precios con el proveedor sabiendo que en enero la demanda baja. El grafico me ayuda a planear mi flujo de efectivo.


## Slide 11 — Comparativa anual: 2024 vs 2025

Pasemos al archivo 08 de Comparativa Anual. En contabilidad, comparar periodos es pan de cada dia. El estado de resultados del anio actual sin compararlo con el anterior no dice mucho. Este archivo muestra un estado de resultados simplificado: ingresos y gastos de 2024 versus 2025. El grafico usa barras lado a lado: gris para 2024 porque es el pasado, azul fuerte para 2025 porque es el presente. De un vistazo ven que 2025 crecio. Pero ojo, no se queden solo con los pesos absolutos -- la variacion porcentual les dice si el crecimiento es bueno o mediocre.


## Slide 12 — Estado de Resultados comparativo visual

Veamos la estructura del estado de resultados. Arriba van los ingresos: ventas y otros ingresos. Abajo van los gastos: lo que compramos, los gastos generales y la nomina. La diferencia es la utilidad bruta. En el archivo tenemos dos graficos separados: uno para ingresos en azul y otro para gastos en rojo. Por que separarlos? Porque si los pongo juntos, las escalas son diferentes y se ve confuso. Con graficos separados, cada uno cuenta su historia. El insight mas importante que deben buscar: si los gastos crecieron mas rapido que los ingresos, la utilidad se esta comprimiendo aunque los numeros absolutos se vean bien. Ese es el tipo de analisis que hace un buen contador.


## Slide 13 — Datos y etiquetas profesionales

Los detalles marcan la diferencia. Cuando su grafico va en un reporte o presentacion, necesita ser autonomo -- alguien debe poder entenderlo sin que ustedes lo expliquen. Primero, las etiquetas de datos: pongan el numero exacto sobre cada barra si el grafico es simple. Segundo, si las cifras son en millones, no pongan 85,200,000 -- pongan 85.2M. Tercero, el titulo: nada de 'Grafico 1'. Pongan 'Ventas Mensuales de Combustible 2025 (Monto en Pesos)'. Cuarto, la fuente: digan de donde vienen los datos y que periodo cubren. Y ultimo, el tamano: si su jefe tiene que entrecerrar los ojos para leer la etiqueta, esta muy chica.


## Slide 14 — Resumen del Modulo 3

Recapitulemos lo que aprendimos. Primero, elegir bien el grafico es el ochenta por ciento del trabajo -- si el tipo es correcto, casi se explica solo. Segundo, la combinacion de tabla dinamica con grafico dinamico y segmentador es poderosa para analisis interactivo. Tercero, la limpieza visual es clave: si algo no comunica, quitenlo. Cuarto, los colores tienen significado -- no son decoracion. Quinto, las columnas apiladas son ideales cuando quieren ver composicion y tendencia al mismo tiempo. Sexto, en comparativas anuales siempre muestren el porcentaje de variacion, no solo los numeros absolutos. Y septimo, cada grafico debe contar una sola historia -- si trata de contar tres, no cuenta ninguna.


## Slide 15 — Cierre

En el siguiente modulo vamos a integrar todo lo que hemos aprendido en un dashboard interactivo con segmentadores, graficos dinamicos y un diseno profesional. Nos vemos en el Modulo 4. Practiquen con los archivos de esta sesion.
