# Módulo : Claude en Excel: Tu Segundo Cerebro Contable

## Slide 1 — Portada

Bienvenidos al segundo bonus. Esta es probablemente la sesion que mas les va a cambiar la forma de trabajar. Vamos a ver como usar Claude, la inteligencia artificial de Anthropic, directamente dentro de Excel. No en una ventana aparte, no copiando y pegando: directamente en su hoja de calculo.

## Slide 2 — Que es Claude? Anthropic y la IA Conversacional

Claude es la inteligencia artificial de Anthropic. Si conocen ChatGPT, Claude es su competencia directa, pero con un enfoque diferente: esta disenado para ser especialmente bueno en razonamiento y analisis de datos.

Tiene tres niveles: Haiku que es el mas rapido, Sonnet que es el equilibrado, y Opus que es el mas potente. Para trabajo en Excel, Sonnet les va a funcionar perfecto en la mayoria de los casos.

Lo mas importante para nosotros: desde enero de 2026, Claude esta disponible como add-in oficial para suscriptores de Microsoft 365 Pro. Eso significa que lo pueden instalar directamente desde el Marketplace de Microsoft y usarlo sin salir de Excel.


## Slide 3 — Claude vs Copilot: Diferencias y Complementos

La pregunta que todos hacen: y entonces, uso Copilot o uso Claude?

La respuesta es: ambos. No son competencia, son complementos. Copilot esta integrado nativamente en Excel y es excelente para tareas rapidas: 'dame la suma de esta columna', 'crea un grafico de barras', 'ordena por fecha'.

Claude brilla en analisis mas profundo. Pueden decirle: 'Analiza esta tabla de nomina, identifica patrones anomalos en las horas extra de los ultimos 6 meses, y sugiere si hay empleados que sistematicamente exceden el presupuesto.' Eso es razonamiento, y ahi Claude es muy fuerte.

Mi recomendacion: usen Copilot para lo rapido y Claude para lo complejo. Son como tener dos asistentes con habilidades distintas.


## Slide 4 — Instalacion del Add-in desde Microsoft Marketplace

Vamos a instalar el add-in paso a paso. Necesitan tener Microsoft 365 Pro activo. Si tienen la version gratuita o la version basica, no les va a aparecer.

Abran Excel, vayan a la pestana Insertar, y busquen 'Obtener complementos' o 'Get Add-ins'. En el Marketplace busquen 'Claude for Excel'. Lo van a ver con el logo de Anthropic. Clic en Agregar.

Les va a pedir permisos. Basicamente Claude necesita acceso a los datos de su hoja para poder analizarlos. Acepten los permisos.

Una vez instalado, van a ver un panel lateral nuevo. Para verificar que funciona, escriban algo simple como 'Hola, que puedes hacer?' y Claude les responde.

Desde enero 2026 el add-in viene como complemento oficial de Microsoft 365 Pro. La instalacion es directa desde el Marketplace, sin necesidad de configuraciones adicionales o claves API.


## Slide 5 — Caso 1: Analizar Tabla de Nomina con Lenguaje Natural

Primer caso practico. Tenemos nuestra tabla de nomina con 200 empleados. En vez de crear tablas dinamicas y graficos manualmente, le decimos a Claude:

'Analiza esta tabla de nomina. Dime cuales empleados ganan mas que el promedio de su departamento, que departamento tiene el mayor gasto total en horas extra, y si hay alguna anomalia en las deducciones de IMSS.'

Claude analiza toda la tabla y responde con: los nombres, los montos, y las anomalias. Si le piden, genera un resumen ejecutivo en una hoja nueva.

Esto no reemplaza su criterio como contadores. Claude les da los datos; ustedes deciden que hacer con ellos. Pero les ahorra el tiempo de encontrar esos datos manualmente.


## Slide 6 — Caso 2: Generar Formulas Complejas

Segundo caso. Las formulas complejas son el dolor de cabeza de todo contador. Especialmente cuando necesitan anidar funciones.

Le dicen a Claude: 'Necesito una formula que busque el RFC del proveedor en la columna A de la tabla Proveedores, traiga el nombre de la columna 3, y si no lo encuentra muestre No encontrado.'

Claude genera:
=SIERROR(BUSCARV(A2,Proveedores,3,0),"No encontrado")

Pero lo mejor es que tambien les explica: 'BUSCARV busca el valor de A2 en la primera columna de Proveedores. El 3 indica que devuelve la tercera columna. El 0 indica coincidencia exacta. SIERROR captura el error si no existe.'

Pueden pedirle formulas mucho mas complejas: INDICE/COINCIDIR con multiples criterios, SUMAR.SI.CONJUNTO con rangos dinamicos, o funciones LET para formulas mas legibles. Claude las genera y las explica.


## Slide 7 — Caso 3: Explicacion de Formula Heredada

Este caso es mi favorito. Quien no ha recibido un archivo de Excel del contador anterior con formulas que parecen escritas en jeroglificos?

Seleccionan la celda, copian la formula, y le dicen a Claude: 'Explicame que hace esta formula paso a paso, en espaniol, como si yo fuera contador y no programador.'

Claude descompone cada funcion: 'Primero, SI evalua si B2 es mayor a 10000. Si es verdadero, aplica BUSCARV en la tabla de tarifas. El resultado se multiplica por el factor de la celda H1. Si es falso, devuelve 0.'

Tambien pueden pedirle: 'Reescribe esta formula de forma mas legible usando la funcion LET.' Y Claude la reorganiza para que sea mas facil de entender y mantener.

Esto es un cambio de juego para auditorias y transiciones de puesto.


## Slide 8 — MCP Connectors: Fuentes Externas

Esto es mas avanzado pero quiero que lo conozcan. MCP significa Model Context Protocol. Es un protocolo que creo Anthropic para que Claude pueda conectarse a fuentes de datos externas.

Que significa en la practica? Que Claude no solo puede analizar los datos que estan en su hoja de Excel, sino que puede ir a buscar datos a otras fuentes: una base de datos SQL, archivos en SharePoint, o incluso APIs externas.

Imaginen esto: estan en Excel y le dicen a Claude 'Trae las ventas del ultimo trimestre de la base de datos SQL y comparalas con el presupuesto que tengo en esta hoja.' Claude va a la base de datos, trae los datos, y los analiza contra su hoja. Todo sin salir de Excel.

Esto requiere configuracion por parte del area de TI. Los conectores MCP se configuran a nivel empresa. Pero el potencial es enorme.


## Slide 9 — Claude Code: Automatizacion desde Terminal

Claude Code es otra forma de usar Claude, pero desde la terminal o linea de comandos. Esto es para los que quieren llevar la automatizacion al siguiente nivel.

Imaginen que cada mes reciben 20 archivos de Excel de diferentes sucursales y necesitan consolidarlos. En vez de abrir cada uno, copiar, pegar... le dicen a Claude Code: 'Abre todos los archivos .xlsx de la carpeta Sucursales, toma la hoja Ventas de cada uno, y consolida todo en un archivo nuevo.'

Claude Code genera un script en Python, lo ejecuta, y les entrega el archivo consolidado. Todo automatico.

No es necesario que sepan Python. Claude Code lo genera por ustedes. Solo necesitan describir lo que quieren lograr.

El add-in en Excel es para analisis interactivo. Claude Code es para automatizacion en lote. Juntos cubren el 95%% de las necesidades.


## Slide 10 — Mejores Practicas: Validacion y Privacidad

Puntos importantes de mejores practicas. Primero: SIEMPRE validen lo que Claude les dice. La IA es muy buena pero no es perfecta. Puede equivocarse en calculos, puede malinterpretar datos, puede inventar algo que suena correcto pero no lo es. Ustedes son los contadores; ustedes validan.

Segundo: privacidad. Verifiquen con su area de TI o con su organizacion que politicas tienen sobre enviar datos a servicios de IA en la nube. Claude no almacena los datos de las hojas que procesa, pero es importante que su organizacion este de acuerdo.

Tercero: creen una biblioteca de prompts. Los prompts que les funcionan bien, guardenlos en un documento. Asi no tienen que reinventar la rueda cada vez.

Y cuarto: capaciten a su equipo. La IA no reemplaza contadores; amplifica su productividad. El que la sabe usar tiene ventaja competitiva.


## Slide 11 — El Futuro del Contador: IA como Herramienta, Criterio Humano como Brujula

Quiero terminar con una reflexion. Hay mucho miedo de que la IA va a reemplazar a los contadores. No es cierto. La IA reemplaza tareas, no profesionales.

Las tareas repetitivas: captura de datos, formateo de reportes, busqueda de informacion... esas si se automatizan. Y esta bien. Esas tareas no son lo que hace valioso a un contador.

Lo que hace valioso a un contador es el juicio profesional: saber interpretar los numeros, identificar riesgos, tomar decisiones eticas, comunicar resultados a la gerencia. Eso la IA no lo hace.

El contador del futuro es alguien que sabe usar la IA como herramienta pero mantiene su criterio profesional como brujula. Ustedes, por estar aqui, ya estan en ese camino.

Excel mas IA es un superpoder contable. Y ahora ustedes lo tienen.


## Slide 12 — Cierre

Y con esto terminamos el segundo bonus y cerramos la seccion de bonuses del curso. Repasamos que es Claude, como instalarlo en Excel, vimos tres casos practicos de uso contable, conocimos MCP y Claude Code, y hablamos de mejores practicas y el futuro de la profesion.

La invitacion es: instalen el add-in, prueben los tres casos que vimos, y empiecen a construir su biblioteca de prompts. Cada prompt que funciona es tiempo que ahorran el proximo mes.

Gracias por acompanarme en todo el curso. Espero que Excel haya dejado de ser solo una hoja de calculo para ustedes y se haya convertido en una herramienta de poder para su carrera. Nos vemos en todoconta.com. Exito.
