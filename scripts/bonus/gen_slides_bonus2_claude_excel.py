"""
Generador: Bonus_2_Claude_en_Excel.pptx + Script_Bonus_2_Claude.md
Bonus 2 — Claude en Excel: Tu Segundo Cerebro Contable

Slides:
  1. Portada
  2. Que es Claude? Anthropic y la IA conversacional
  3. Claude vs Copilot: diferencias y complementos
  4. Instalacion del add-in desde Microsoft Marketplace
  5. Caso 1 - Analizar tabla de nomina con lenguaje natural
  6. Caso 2 - Generar formulas complejas (BUSCARV anidado + SIERROR)
  7. Caso 3 - Explicacion de formula heredada
  8. MCP Connectors: fuentes externas
  9. Claude Code: automatizacion desde terminal
  10. Mejores practicas: validacion, privacidad
  11. El futuro del contador: IA como herramienta, criterio humano como brujula
  12. Recursos
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import SLIDES_DIR, TELEPROMPTER_DIR
from scripts.generators.pptx_gen import SlidesGenerator


def build():
    gen = SlidesGenerator(
        filename="Bonus_2_Claude_en_Excel.md",
        output_dir=SLIDES_DIR,
        script_filename="Script_Bonus_2_Claude.md",
        script_dir=TELEPROMPTER_DIR,
    )

    # ── Slide 1 — Portada ─────────────────────────────────────────
    gen.add_title_slide(
        modulo_num="",
        modulo_nombre="Claude en Excel: Tu Segundo Cerebro Contable",
        subtitulo="Bonus 2 — Inteligencia Artificial integrada en tu hoja de calculo",
    )
    gen.script_lines.append(
        "Bienvenidos al segundo bonus. Esta es probablemente la sesion que mas "
        "les va a cambiar la forma de trabajar. Vamos a ver como usar Claude, "
        "la inteligencia artificial de Anthropic, directamente dentro de Excel. "
        "No en una ventana aparte, no copiando y pegando: directamente en su "
        "hoja de calculo.\n"
    )

    # ── Slide 2 — Que es Claude? ──────────────────────────────────
    gen.add_content_slide(
        title="Que es Claude? Anthropic y la IA Conversacional",
        bullets=[
            "Claude es un asistente de IA creado por Anthropic",
            "Especializado en razonamiento, analisis de datos y codigo",
            "Modelos disponibles: Claude Haiku (rapido), Sonnet (equilibrado), Opus (potente)",
            "Diferencia clave: enfoque en seguridad y respuestas utiles",
            "Desde enero 2026: add-in oficial para Microsoft 365 Pro",
            "Gratuito para funciones basicas; plan Pro para volumen alto",
        ],
        script_text=(
            "Claude es la inteligencia artificial de Anthropic. Si conocen ChatGPT, Claude es "
            "su competencia directa, pero con un enfoque diferente: esta disenado para ser "
            "especialmente bueno en razonamiento y analisis de datos.\n\n"
            "Tiene tres niveles: Haiku que es el mas rapido, Sonnet que es el equilibrado, "
            "y Opus que es el mas potente. Para trabajo en Excel, Sonnet les va a funcionar "
            "perfecto en la mayoria de los casos.\n\n"
            "Lo mas importante para nosotros: desde enero de 2026, Claude esta disponible "
            "como add-in oficial para suscriptores de Microsoft 365 Pro. Eso significa que "
            "lo pueden instalar directamente desde el Marketplace de Microsoft y usarlo "
            "sin salir de Excel.\n"
        ),
    )

    # ── Slide 3 — Claude vs Copilot ───────────────────────────────
    gen.add_content_slide(
        title="Claude vs Copilot: Diferencias y Complementos",
        bullets=[
            "Copilot (Microsoft): integrado nativamente, bueno en tareas rapidas",
            "Claude (Anthropic): superior en analisis profundo y razonamiento",
            "Copilot: 'Dame la suma de ventas' -> rapido pero superficial",
            "Claude: 'Analiza tendencias de las ultimas 12 nominas y sugiere ajustes' -> profundo",
            "No son competencia: son complementos",
            "Copilot para automatizacion rapida + Claude para analisis complejo",
            "Ambos funcionan con lenguaje natural en espanol",
        ],
        script_text=(
            "La pregunta que todos hacen: y entonces, uso Copilot o uso Claude?\n\n"
            "La respuesta es: ambos. No son competencia, son complementos. Copilot "
            "esta integrado nativamente en Excel y es excelente para tareas rapidas: "
            "'dame la suma de esta columna', 'crea un grafico de barras', 'ordena por fecha'.\n\n"
            "Claude brilla en analisis mas profundo. Pueden decirle: 'Analiza esta tabla de "
            "nomina, identifica patrones anomalos en las horas extra de los ultimos 6 meses, "
            "y sugiere si hay empleados que sistematicamente exceden el presupuesto.' Eso es "
            "razonamiento, y ahi Claude es muy fuerte.\n\n"
            "Mi recomendacion: usen Copilot para lo rapido y Claude para lo complejo. "
            "Son como tener dos asistentes con habilidades distintas.\n"
        ),
    )

    # ── Slide 4 — Instalacion ─────────────────────────────────────
    gen.add_content_slide(
        title="Instalacion del Add-in desde Microsoft Marketplace",
        bullets=[
            "Requisito: Microsoft 365 Pro (suscripcion activa)",
            "Pasos:",
            "   1. Abrir Excel -> Insertar -> Obtener complementos",
            "   2. Buscar 'Claude for Excel' en el Marketplace",
            "   3. Clic en Agregar -> Aceptar permisos",
            "   4. Aparece panel lateral 'Claude' en la pestana Inicio",
            "Configuracion inicial: ingresar con cuenta Anthropic o M365",
            "Verificar: escribir en el panel '=Hola, Claude' -> debe responder",
        ],
        script_text=(
            "Vamos a instalar el add-in paso a paso. Necesitan tener Microsoft 365 Pro "
            "activo. Si tienen la version gratuita o la version basica, no les va a aparecer.\n\n"
            "Abran Excel, vayan a la pestana Insertar, y busquen 'Obtener complementos' o "
            "'Get Add-ins'. En el Marketplace busquen 'Claude for Excel'. Lo van a ver con el "
            "logo de Anthropic. Clic en Agregar.\n\n"
            "Les va a pedir permisos. Basicamente Claude necesita acceso a los datos de su hoja "
            "para poder analizarlos. Acepten los permisos.\n\n"
            "Una vez instalado, van a ver un panel lateral nuevo. Para verificar que funciona, "
            "escriban algo simple como 'Hola, que puedes hacer?' y Claude les responde.\n\n"
            "Desde enero 2026 el add-in viene como complemento oficial de Microsoft 365 Pro. "
            "La instalacion es directa desde el Marketplace, sin necesidad de configuraciones "
            "adicionales o claves API.\n"
        ),
    )

    # ── Slide 5 — Caso 1: Analizar nomina ─────────────────────────
    gen.add_content_slide(
        title="Caso 1: Analizar Tabla de Nomina con Lenguaje Natural",
        bullets=[
            "Seleccionar tabla de nomina (Ctrl+T si no es tabla)",
            "En panel Claude: 'Analiza esta tabla de nomina e identifica:'",
            "   - Empleados con sueldo mayor al promedio",
            "   - Departamentos con mayor gasto",
            "   - Anomalias en deducciones",
            "Claude responde con analisis + sugerencias",
            "Puede generar resumen ejecutivo directamente en nueva hoja",
            "Ventaja: analisis que tomaria 30 min, en 30 segundos",
        ],
        script_text=(
            "Primer caso practico. Tenemos nuestra tabla de nomina con 200 empleados. "
            "En vez de crear tablas dinamicas y graficos manualmente, le decimos a Claude:\n\n"
            "'Analiza esta tabla de nomina. Dime cuales empleados ganan mas que el promedio "
            "de su departamento, que departamento tiene el mayor gasto total en horas extra, "
            "y si hay alguna anomalia en las deducciones de IMSS.'\n\n"
            "Claude analiza toda la tabla y responde con: los nombres, los montos, y las anomalias. "
            "Si le piden, genera un resumen ejecutivo en una hoja nueva.\n\n"
            "Esto no reemplaza su criterio como contadores. Claude les da los datos; ustedes "
            "deciden que hacer con ellos. Pero les ahorra el tiempo de encontrar esos datos "
            "manualmente.\n"
        ),
    )

    # ── Slide 6 — Caso 2: Generar formulas ────────────────────────
    gen.add_content_slide(
        title="Caso 2: Generar Formulas Complejas",
        bullets=[
            "Prompt: 'Necesito una formula que busque el RFC en la tabla Proveedores'",
            "Claude genera: =SIERROR(BUSCARV(A2,Proveedores,3,0),\"No encontrado\")",
            "Formulas anidadas que tomarian 15 min de armar manualmente",
            "Claude explica cada parte de la formula",
            "Soporta: BUSCARV, INDICE/COINCIDIR, SUMAR.SI.CONJUNTO, LET",
            "Tip: pedir 'con SIERROR para manejar errores'",
            "La formula se puede insertar directamente en la celda seleccionada",
        ],
        script_text=(
            "Segundo caso. Las formulas complejas son el dolor de cabeza de todo contador. "
            "Especialmente cuando necesitan anidar funciones.\n\n"
            "Le dicen a Claude: 'Necesito una formula que busque el RFC del proveedor en "
            "la columna A de la tabla Proveedores, traiga el nombre de la columna 3, y si "
            "no lo encuentra muestre No encontrado.'\n\n"
            "Claude genera:\n"
            "=SIERROR(BUSCARV(A2,Proveedores,3,0),\"No encontrado\")\n\n"
            "Pero lo mejor es que tambien les explica: 'BUSCARV busca el valor de A2 en la "
            "primera columna de Proveedores. El 3 indica que devuelve la tercera columna. "
            "El 0 indica coincidencia exacta. SIERROR captura el error si no existe.'\n\n"
            "Pueden pedirle formulas mucho mas complejas: INDICE/COINCIDIR con multiples "
            "criterios, SUMAR.SI.CONJUNTO con rangos dinamicos, o funciones LET para "
            "formulas mas legibles. Claude las genera y las explica.\n"
        ),
    )

    # ── Slide 7 — Caso 3: Explicar formula heredada ──────────────
    gen.add_content_slide(
        title="Caso 3: Explicacion de Formula Heredada",
        bullets=[
            "Escenario: reciben archivo de contador anterior con formulas complejas",
            "Seleccionar celda con formula compleja",
            "Prompt: 'Explicame que hace esta formula paso a paso'",
            "Claude descompone cada funcion anidada",
            "Identifica posibles errores o mejoras",
            "Tip: 'Reescribe esta formula de forma mas legible'",
            "Nunca mas tener miedo de archivos heredados",
        ],
        script_text=(
            "Este caso es mi favorito. Quien no ha recibido un archivo de Excel del "
            "contador anterior con formulas que parecen escritas en jeroglificos?\n\n"
            "Seleccionan la celda, copian la formula, y le dicen a Claude: 'Explicame "
            "que hace esta formula paso a paso, en espanol, como si yo fuera contador "
            "y no programador.'\n\n"
            "Claude descompone cada funcion: 'Primero, SI evalua si B2 es mayor a 10000. "
            "Si es verdadero, aplica BUSCARV en la tabla de tarifas. El resultado se "
            "multiplica por el factor de la celda H1. Si es falso, devuelve 0.'\n\n"
            "Tambien pueden pedirle: 'Reescribe esta formula de forma mas legible usando "
            "la funcion LET.' Y Claude la reorganiza para que sea mas facil de entender "
            "y mantener.\n\n"
            "Esto es un cambio de juego para auditorias y transiciones de puesto.\n"
        ),
    )

    # ── Slide 8 — MCP Connectors ──────────────────────────────────
    gen.add_content_slide(
        title="MCP Connectors: Fuentes Externas",
        bullets=[
            "MCP = Model Context Protocol (protocolo de Anthropic)",
            "Permite a Claude conectarse a fuentes externas de datos",
            "Ejemplos de conectores:",
            "   - Base de datos SQL del sistema contable",
            "   - Archivos en SharePoint / OneDrive",
            "   - APIs del SAT (via conector personalizado)",
            "Flujo: Claude lee datos externos -> los analiza -> responde en Excel",
            "Configuracion por administrador de TI",
            "Potencial: Excel como hub central conectado a todo",
        ],
        script_text=(
            "Esto es mas avanzado pero quiero que lo conozcan. MCP significa Model Context "
            "Protocol. Es un protocolo que creo Anthropic para que Claude pueda conectarse "
            "a fuentes de datos externas.\n\n"
            "Que significa en la practica? Que Claude no solo puede analizar los datos que "
            "estan en su hoja de Excel, sino que puede ir a buscar datos a otras fuentes: "
            "una base de datos SQL, archivos en SharePoint, o incluso APIs externas.\n\n"
            "Imaginen esto: estan en Excel y le dicen a Claude 'Trae las ventas del ultimo "
            "trimestre de la base de datos SQL y comparalas con el presupuesto que tengo "
            "en esta hoja.' Claude va a la base de datos, trae los datos, y los analiza "
            "contra su hoja. Todo sin salir de Excel.\n\n"
            "Esto requiere configuracion por parte del area de TI. Los conectores MCP "
            "se configuran a nivel empresa. Pero el potencial es enorme.\n"
        ),
    )

    # ── Slide 9 — Claude Code ─────────────────────────────────────
    gen.add_content_slide(
        title="Claude Code: Automatizacion desde Terminal",
        bullets=[
            "Claude Code = Claude en la linea de comandos (terminal)",
            "Para usuarios avanzados: automatizar procesamiento de archivos Excel",
            "Ejemplo: 'Abre todos los .xlsx de esta carpeta y consolida ventas'",
            "Genera scripts Python (openpyxl, pandas) automaticamente",
            "Ideal para: procesamiento batch, ETL contable, reportes automaticos",
            "Instalacion: npm install -g @anthropic-ai/claude-code",
            "Complementa el add-in: lo que no se puede en Excel, se hace en terminal",
        ],
        script_text=(
            "Claude Code es otra forma de usar Claude, pero desde la terminal o linea de "
            "comandos. Esto es para los que quieren llevar la automatizacion al siguiente nivel.\n\n"
            "Imaginen que cada mes reciben 20 archivos de Excel de diferentes sucursales "
            "y necesitan consolidarlos. En vez de abrir cada uno, copiar, pegar... le dicen "
            "a Claude Code: 'Abre todos los archivos .xlsx de la carpeta Sucursales, toma "
            "la hoja Ventas de cada uno, y consolida todo en un archivo nuevo.'\n\n"
            "Claude Code genera un script en Python, lo ejecuta, y les entrega el archivo "
            "consolidado. Todo automatico.\n\n"
            "No es necesario que sepan Python. Claude Code lo genera por ustedes. Solo "
            "necesitan describir lo que quieren lograr.\n\n"
            "El add-in en Excel es para analisis interactivo. Claude Code es para "
            "automatizacion en lote. Juntos cubren el 95% de las necesidades.\n"
        ),
    )

    # ── Slide 10 — Mejores practicas ──────────────────────────────
    gen.add_content_slide(
        title="Mejores Practicas: Validacion y Privacidad",
        bullets=[
            "SIEMPRE validar las respuestas de Claude (la IA puede equivocarse)",
            "Datos sensibles: verificar politica de privacidad de su organizacion",
            "Claude NO almacena datos de las hojas de calculo procesadas",
            "Para datos altamente confidenciales: usar Claude en modo local/privado",
            "No depender 100% de la IA: el criterio contable es insustituible",
            "Documentar prompts utiles para reutilizarlos (crear biblioteca de prompts)",
            "Capacitar al equipo: la IA amplifica la productividad de todos",
        ],
        script_text=(
            "Puntos importantes de mejores practicas. Primero: SIEMPRE validen lo que "
            "Claude les dice. La IA es muy buena pero no es perfecta. Puede equivocarse "
            "en calculos, puede malinterpretar datos, puede inventar algo que suena correcto "
            "pero no lo es. Ustedes son los contadores; ustedes validan.\n\n"
            "Segundo: privacidad. Verifiquen con su area de TI o con su organizacion "
            "que politicas tienen sobre enviar datos a servicios de IA en la nube. Claude "
            "no almacena los datos de las hojas que procesa, pero es importante que su "
            "organizacion este de acuerdo.\n\n"
            "Tercero: creen una biblioteca de prompts. Los prompts que les funcionan bien, "
            "guardenlos en un documento. Asi no tienen que reinventar la rueda cada vez.\n\n"
            "Y cuarto: capaciten a su equipo. La IA no reemplaza contadores; amplifica "
            "su productividad. El que la sabe usar tiene ventaja competitiva.\n"
        ),
    )

    # ── Slide 11 — El futuro del contador ─────────────────────────
    gen.add_content_slide(
        title="El Futuro del Contador: IA como Herramienta, Criterio Humano como Brujula",
        bullets=[
            "La IA NO reemplaza contadores; reemplaza tareas repetitivas",
            "El contador del futuro: analista + estratega + verificador",
            "Habilidades clave: saber preguntar (prompting) + saber validar",
            "Excel + IA = superpoder contable",
            "Lo que la IA no puede hacer: juicio profesional, etica, relaciones",
            "Su ventaja competitiva: dominar ambos mundos",
            "El mejor momento para aprender IA fue ayer; el segundo mejor es hoy",
        ],
        script_text=(
            "Quiero terminar con una reflexion. Hay mucho miedo de que la IA va a reemplazar "
            "a los contadores. No es cierto. La IA reemplaza tareas, no profesionales.\n\n"
            "Las tareas repetitivas: captura de datos, formateo de reportes, busqueda de "
            "informacion... esas si se automatizan. Y esta bien. Esas tareas no son lo que "
            "hace valioso a un contador.\n\n"
            "Lo que hace valioso a un contador es el juicio profesional: saber interpretar "
            "los numeros, identificar riesgos, tomar decisiones eticas, comunicar resultados "
            "a la gerencia. Eso la IA no lo hace.\n\n"
            "El contador del futuro es alguien que sabe usar la IA como herramienta pero "
            "mantiene su criterio profesional como brujula. Ustedes, por estar aqui, "
            "ya estan en ese camino.\n\n"
            "Excel mas IA es un superpoder contable. Y ahora ustedes lo tienen.\n"
        ),
    )

    # ── Slide 12 — Recursos y Cierre ──────────────────────────────
    gen.add_closing_slide(
        next_module="",
        resources=[
            "Add-in Claude para Excel: Marketplace de Microsoft 365",
            "Documentacion Claude: docs.anthropic.com",
            "Claude Code: npm install -g @anthropic-ai/claude-code",
            "MCP Protocol: modelcontextprotocol.io",
            "Practica: instalar el add-in y probar los 3 casos de esta sesion",
            "Comunidad: todoconta.com para preguntas y recursos adicionales",
        ],
    )
    gen.script_lines.append(
        "Y con esto terminamos el segundo bonus y cerramos la seccion de bonuses "
        "del curso. Repasamos que es Claude, como instalarlo en Excel, vimos tres "
        "casos practicos de uso contable, conocimos MCP y Claude Code, y hablamos "
        "de mejores practicas y el futuro de la profesion.\n\n"
        "La invitacion es: instalen el add-in, prueben los tres casos que vimos, "
        "y empiecen a construir su biblioteca de prompts. Cada prompt que funciona "
        "es tiempo que ahorran el proximo mes.\n\n"
        "Gracias por acompanarme en todo el curso. Espero que Excel haya dejado de "
        "ser solo una hoja de calculo para ustedes y se haya convertido en una "
        "herramienta de poder para su carrera. Nos vemos en todoconta.com. Exito.\n"
    )

    gen.save()
    print("Bonus 2 - Claude en Excel generado correctamente.")


if __name__ == "__main__":
    build()
