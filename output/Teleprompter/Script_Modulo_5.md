# Modulo 5: Automatizacion Nativa con Microsoft 365 Copilot

## Slide 1 — Portada

Bienvenidos al Modulo 5, el ultimo modulo de nuestro curso. Hoy vamos a explorar como la Inteligencia Artificial, especificamente Microsoft 365 Copilot, puede transformar la forma en que trabajamos con Excel. Este modulo no es sobre reemplazar al contador, sino sobre potenciar sus capacidades con herramientas de nueva generacion.

## Slide 2 — El futuro ya esta aqui

La inteligencia artificial generativa no es ciencia ficcion. Desde 2023, Microsoft integro modelos de lenguaje avanzados directamente en las aplicaciones de Office. Esto significa que puedes hablarle a Excel en espanol, pedirle que analice tus datos, y obtener resultados en segundos. No necesitas saber programar, no necesitas escribir formulas complejas. Le describes lo que necesitas y Copilot lo ejecuta. El profesional que domine estas herramientas hoy tendra una ventaja enorme en el mercado laboral de manana.

## Slide 3 — Requisitos para usar Copilot en Excel

Antes de empezar, veamos los requisitos. Primero, necesitas una licencia de Microsoft 365 con el add-on de Copilot. Segundo, y esto es muy importante: tu archivo debe estar guardado en OneDrive o SharePoint. Si lo tienes en el escritorio local, Copilot simplemente no aparece. Tercero, tus datos deben estar en formato de Tabla de Excel, con Ctrl+T, y con un nombre descriptivo. Esto es porque Copilot necesita entender la estructura de tus datos para poder analizarlos. Si tus datos estan en un rango suelto sin formato de tabla, Copilot no podra trabajar con ellos.

## Slide 4 — Activando el panel de Copilot

Vamos paso a paso. Primero guardas en OneDrive. Luego seleccionas tus datos y presionas Ctrl+T para convertirlos en tabla. Le pones un nombre descriptivo como 'Ventas_Gasolinera' o 'Nomina_Empleados'. Despues, en la pestana Inicio, veras el icono de Copilot, que parece un asistente con destellos. Lo presionas y se abre un panel lateral a la derecha. Ahi es donde escribiras tus solicitudes en lenguaje natural. Si el boton esta gris o no aparece, revisa que tu licencia este activa y que el archivo este en la nube.

## Slide 5 — Analisis con lenguaje natural

Veamos el primer caso de uso: analisis con lenguaje natural. Imagina que tienes una tabla con 1,250 transacciones de ventas de gasolinera. En lugar de crear una tabla dinamica, filtrar, agrupar y formatear, simplemente le escribes a Copilot: 'Analiza las ventas por sucursal y dime cual tiene mejor desempeno'. En segundos, Copilot genera una tabla resumen con totales, promedios y te dice cual sucursal lidera. Lo que normalmente te toma 5-10 minutos, se hace en 15 segundos. Eso si, a veces la interpretacion no es exactamente lo que esperabas, asi que hay que aprender a formular buenos prompts.

## Slide 6 — Generacion de formulas con Copilot

Otro caso de uso poderoso es la generacion de formulas. Le puedes pedir a Copilot que calcule el ISR marginal basandose en la percepcion total, y el genera una formula compleja que aplica la tarifa del articulo 96. Lo interesante es que entiende el contexto contable mexicano. Le puedes decir 'calcula comision del 2 por ciento' o 'clasifica como Alta, Media o Baja' y genera la formula correspondiente. Pero atencion: siempre debes revisar la formula antes de aceptar. He visto casos donde Copilot genera una formula que se ve correcta pero tiene un error sutil en los rangos o en la logica.

## Slide 7 — Columnas inteligentes

Las columnas inteligentes son una de las funciones mas practicas. Le pides a Copilot que agregue una nueva columna con cierta logica, y la crea automaticamente dentro de tu tabla existente. Por ejemplo, le pides que clasifique cada venta como Alta si supera 5 mil pesos, Media entre mil y 5 mil, y Baja si es menor a mil. Copilot crea la columna, le pone nombre y aplica un SI anidado o IFS. Luego puedes pedir otra columna encima: 'ahora agrega comision del 2 por ciento'. Las columnas se van sumando a tu tabla y puedes seguir construyendo analisis.

## Slide 8 — Visualizacion instantanea

Copilot tambien puede crear graficos al instante. Le dices 'crea un grafico de barras de ventas por mes' y genera el grafico directamente en tu hoja. Soporta barras, lineas, pastel, dispersion y combinados. Los graficos se crean con colores y etiquetas predeterminadas. Para los que vimos en el Modulo 3 sobre visualizacion profesional, quiero ser honesto: los graficos de Copilot son funcionales pero basicos. Para un reporte ejecutivo que impacte, vas a querer ajustar colores, fuentes y formato manualmente. Copilot te da el punto de partida rapido.

## Slide 9 — Deteccion de anomalias (Insights)

Una de las funciones mas valiosas para contadores es la deteccion de anomalias. En nuestro dataset de nomina, hay anomalias intencionales: dos empleados con incrementos subitos de sueldo y uno con meses faltantes. Cuando le pides a Copilot que identifique anomalias, puede detectar estos patrones. Esto es increiblemente util para auditorias internas: en lugar de revisar 800 registros linea por linea, Copilot te senala donde estan las banderas rojas. Pero ojo: esto no sustituye una auditoria formal. Es una herramienta de deteccion inicial que te ahorra horas de trabajo manual.

## Slide 10 — Caso practico: 3 prompts contables en secuencia

Vamos a ver un caso practico completo con 3 prompts en secuencia. Primero le preguntamos cual vendedor tiene peor desempeno. Copilot analiza los datos y nos dice que Vendedor_3 tiene las transacciones mas bajas de forma consistente. Luego le pedimos un grafico comparativo de ese vendedor versus el promedio general. Copilot genera un grafico de barras que muestra visualmente la brecha. Finalmente, le pedimos un resumen ejecutivo con recomendacion. Copilot genera un parrafo con los KPIs clave y sugiere acciones. En 3 minutos, tienes un mini-analisis que normalmente tomaria media hora entre tablas dinamicas, graficos y redaccion.

## Slide 11 — Limitaciones: "El criterio contable es tuyo"

Ahora hablemos de lo que Copilot NO puede hacer. Y esto es crucial. Copilot no conoce las Normas de Informacion Financiera en detalle. No sabe en que regimen fiscal estas. Puede generar una formula de ISR que se vea perfecta pero tenga un error sutil en el porcentaje o en el rango. No puede firmar tu declaracion anual ni sustituir tu criterio profesional. La regla de oro es simple: usa Copilot para acelerar el trabajo mecanico, para explorar datos rapidamente, para generar borradores de analisis. Pero SIEMPRE valida los resultados con tu conocimiento. El criterio contable es tuyo y solo tuyo.

## Slide 12 — IA Externa: ChatGPT, Gemini, Claude

Microsoft Copilot no es la unica opcion. Existen otras herramientas de IA que pueden complementar tu trabajo como contador. ChatGPT de OpenAI es excelente para generar codigo de macros VBA o consultas de Power Query. Le describes lo que quieres automatizar y te genera el codigo listo para pegar. Google Gemini esta integrado en Google Sheets y es una buena alternativa gratuita. Claude de Anthropic es muy preciso para analizar documentos y textos largos. Un truco que uso frecuentemente: cuando tengo una formula compleja que no entiendo, la copio y la pego en ChatGPT y le pido que me la explique paso a paso. O cuando necesito una macro, le describo a Claude lo que quiero y me genera el codigo.

## Slide 13 — Caso: Macro sencilla generada por IA

Veamos un ejemplo concreto de como usar IA externa con Excel. Supongamos que necesitas una macro que formatee tu reporte mensual automaticamente: encabezados con color, bordes, ancho de columnas, pie de pagina con fecha. En lugar de aprender VBA desde cero, le dices a ChatGPT o Claude: 'Genera una macro VBA que aplique formato profesional a mi reporte'. La IA genera el codigo completo con comentarios. Tu lo copias, abres el Editor de VBA con Alt+F11, pegas el codigo y lo ejecutas. Otros ejemplos utiles: una macro que exporte cada hoja como PDF individual, o una que consolide datos de multiples archivos en uno solo. El consejo clave es pedirle a la IA que agregue comentarios al codigo para que tu entiendas que hace cada linea.

## Slide 14 — Resumen del curso completo: M1 a M5

Hagamos un recorrido rapido por todo lo que cubrimos en este curso. En el Modulo 1 sentamos las bases con logica contable y funciones clave: BUSCARV, SI, calculo de ISR con la tarifa oficial, factor de actualizacion del CFF. En el Modulo 2 pasamos a procesamiento masivo con tablas dinamicas, aprendiendo a analizar miles de registros en minutos. El Modulo 3 nos enseno a comunicar resultados con graficos profesionales y reportes ejecutivos que impactan. En el Modulo 4 construimos dashboards interactivos con KPIs y semaforos para la toma de decisiones. Y hoy, en el Modulo 5, incorporamos la inteligencia artificial como aceleradora de todo lo anterior. Juntos, estos 5 modulos te llevan del dato crudo al insight accionable con herramientas modernas.

## Slide 15 — Comunidad y siguiente paso

Este no es el final, es el comienzo. Te invito a unirte a nuestra comunidad de contadores y administrativos que usan Excel de forma avanzada. Practica con los archivos del Pack Excel Pro que incluimos en cada modulo. Comparte tus logros: toma un proceso que antes te tomaba horas y muestra como lo haces ahora en minutos. Microsoft actualiza Copilot constantemente, asi que mantente al dia con las nuevas funciones. Y si quieres seguir creciendo, el siguiente nivel es Power Query para transformacion de datos, Power BI para dashboards avanzados, y Power Automate para automatizar flujos completos. Tu inversion en aprendizaje hoy se traduce directamente en eficiencia y valor profesional manana.

## Slide 16 — Cierre

Muchas gracias por acompanarme en este curso. Fue un placer compartir estos conocimientos con ustedes. Recuerden: la tecnologia es la herramienta, pero el profesional son ustedes. Sigan aprendiendo, sigan practicando, y sigan creciendo. Exito en todo lo que emprendan. Nos vemos en la comunidad.
