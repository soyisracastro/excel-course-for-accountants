# Módulo 4: El Dashboard Inteligente y Entrega Profesional

## Slide 1 — Portada

Bienvenidos al Modulo 4, el modulo donde todo lo que hemos construido se une en un dashboard profesional. Vamos a tomar la calculadora ISR del Modulo 1, los datos de nomina del Modulo 2, los graficos del Modulo 3, y vamos a integrarlos en un panel de control que impresione. Ademas, veremos como proteger y distribuir nuestro trabajo de forma profesional.

## Slide 2 — Que es un Dashboard?

Un dashboard es basicamente un tablero de control, como el tablero de tu auto. No necesitas abrir el motor para saber si todo esta bien -- solo ves los indicadores. En contabilidad, un dashboard te muestra de un vistazo cuanto llevas de nomina, cuanto de ISR, si alguna e.firma esta por vencer. Es un resumen ejecutivo visual, no un reporte completo. Y lo mejor: en Excel podemos construirlo sin programar, usando las herramientas que ya conocemos.


## Slide 3 — Principios de Diseno: Menos es Mas

Antes de construir, hablemos de diseno. La regla mas importante es: menos es mas. Tu dashboard debe comunicar el mensaje principal en 5 segundos. Si alguien necesita 30 segundos para entender que pasa, has fallado. Maximo 4 a 6 KPIs. Colores consistentes: nosotros usamos azul como color principal, rojo para alertas, verde para 'todo bien'. Los graficos deben ser simples: barras para comparar y lineas para tendencias. Y algo que muchos olvidan: oculta las lineas de cuadricula. Eso solo da un aspecto mucho mas profesional.


## Slide 4 — Segmentadores de Datos (Slicers)

Los segmentadores, o slicers en ingles, son la magia detras de un dashboard interactivo. Son filtros visuales que aparecen como botones. El usuario simplemente hace clic en 'Enero' y todo el dashboard se actualiza. Para insertarlos, primero necesitas una Tabla Dinamica. Luego vas a Insertar, Segmentacion de datos, y seleccionas los campos que quieres filtrar. Lo mas poderoso es que puedes conectar un solo slicer a multiples Tablas Dinamicas. Asi, un clic filtra todos los graficos y KPIs al mismo tiempo.


## Slide 5 — Caso Practico: Slicers Vinculados a TDs

Vamos a verlo paso a paso. Primero creamos dos Tablas Dinamicas desde nuestra tabla de nomina. La primera muestra sueldos e ISR por periodo. La segunda muestra empleados por puesto. Luego insertamos un segmentador de Periodo. Por defecto solo esta conectado a TD1. Pero si hacemos clic derecho en el slicer y vamos a Conexiones de informe, podemos marcar tambien TD2. Ahora, cuando seleccionamos 'Marzo', ambas tablas y sus graficos se filtran. Esto es lo que hace que un dashboard en Excel sea realmente interactivo.


## Slide 6 — Proyecto Final: La Gran Integracion

Este es el momento de la verdad. Todo lo que hemos aprendido converge aqui. Del Modulo 1 tomamos la logica de calculo ISR con BUSCARV. Del Modulo 2, los datos masivos de nomina y las Tablas Dinamicas. Del Modulo 3, los graficos y la visualizacion. Y ahora en el Modulo 4, los unimos en un dashboard interactivo. El archivo de trabajo es el Dashboard Final Integrado. Tiene 4 hojas: datos de nomina, tarifa ISR, calculadora y el dashboard. Vamos a ver como se construye.


## Slide 7 — KPIs: Los Numeros que Importan

Los KPIs son los numeros clave que van en la parte superior del dashboard. Nosotros usamos 4: Total Percepciones, que es la suma de todos los sueldos. Total Deducciones, que incluye ISR e IMSS. ISR del Periodo como desglose. Y el estatus de la e.firma, que conecta con lo que vimos en el Modulo 1. Un detalle tecnico importante: usamos SUBTOTAL con el codigo 109, que es SUMA pero ignorando filas ocultas por filtros. Asi, cuando el usuario filtra con un slicer, los KPIs se recalculan. Si quieres aun mas precision, puedes usar GETPIVOTDATA para leer directamente desde la Tabla Dinamica.


## Slide 8 — Construyendo el Layout del Dashboard

El layout es la estructura visual de tu dashboard. Piensa en el como el plano de una casa antes de construir. Arriba va el titulo. Debajo, los KPIs en cuadros de colores. A la izquierda, la zona de segmentadores. En el centro, los graficos principales -- normalmente dos: uno de barras y uno de linea. Abajo, espacio para una tabla de detalle o un grafico adicional. En el archivo template que les di, ya tienen esta estructura lista. Solo necesitan llenarla con sus Tablas Dinamicas y graficos.


## Slide 9 — Proteccion de Celdas y Hojas

Ahora que tu dashboard funciona, hay que protegerlo. No quieres que alguien borre una formula por accidente. El truco es que en Excel todas las celdas ya vienen bloqueadas, pero el bloqueo no se activa hasta que proteges la hoja. Entonces el flujo es: primero, desbloqueas las celdas donde el usuario SI debe escribir. Segundo, las marcas con fondo amarillo para que sepa donde ir. Tercero, proteges la hoja con contrasena. Y en las opciones de proteccion, asegurate de permitir: seleccionar celdas, usar filtros y usar Tablas Dinamicas. Asi el dashboard sigue siendo interactivo pero las formulas quedan protegidas.


## Slide 10 — Distribucion: PDF vs Excel Protegido

La pregunta final es: como lo comparto? Hay dos opciones principales. PDF es ideal para reportes finales. No se puede editar, no requiere Excel, y puedes firmarlo digitalmente. Excel protegido es para cuando el destinatario necesita interactuar: cambiar filtros, usar slicers, ingresar datos. Si necesitas seguridad real, usa la contrasena de apertura de archivo, que usa cifrado AES. Eso es como ponerle llave al archivo. La contrasena de escritura permite que lo abran pero no lo modifiquen. Y 'Marcar como final' es solo una sugerencia, el usuario la puede quitar. Mi recomendacion: para el cliente, PDF. Para tu equipo, Excel protegido.


## Slide 11 — Demo: Dashboard Final en Accion

Vamos a hacer la demo paso a paso. Abran el archivo Dashboard Final Integrado. Primero, revisen la hoja Datos Nomina. Son 240 registros: 20 empleados por 12 meses. Los sueldos van desde salario minimo hasta 55 mil pesos. Luego vayan a la Calculadora y cambien el sueldo en B4. Vean como BUSCARV calcula automaticamente el ISR. Ahora, creen una Tabla Dinamica con Periodo en filas y Sueldo en valores. Inserten un segmentador de Periodo. Creen un grafico de barras. Hagan lo mismo con otra TD por Puesto. Muevan todo a la hoja Dashboard. Ajusten tamanos. Finalmente, protejan las hojas y exporten un PDF. Eso es su proyecto final del modulo.


## Slide 12 — Resumen del Modulo 4

Hagamos un resumen rapido. Un dashboard es un resumen visual. Los slicers lo hacen interactivo. Los KPIs con SUBTOTAL se actualizan con filtros. El layout sigue una jerarquia clara. Protege tus formulas y comparte correctamente. Y lo mas importante: este modulo no es un tema aislado. Es la integracion de todo el curso. La calculadora ISR, los datos de nomina, los graficos, todo converge en un dashboard profesional. Esto es lo que puedes presentar a tu jefe o a tu cliente.


## Slide 13 — Cierre

Estos son los recursos del modulo. Tienen el template de layout, el archivo integrado con datos, la guia de proteccion en PDF y la referencia rapida del modulo. Mi recomendacion es que practiquen con datos reales de su propia empresa. Tomen su nomina real, cambien los nombres si quieren por privacidad, y construyan su propio dashboard. En el siguiente modulo veremos como Microsoft 365 Copilot puede ayudarnos a automatizar aun mas nuestro trabajo con inteligencia artificial. Nos vemos ahi!
