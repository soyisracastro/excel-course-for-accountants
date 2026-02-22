# Claude en Excel: Tu Segundo Cerebro Contable

*Bonus 2 — Inteligencia Artificial integrada en tu hoja de calculo*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## Que es Claude? Anthropic y la IA Conversacional

- Claude es un asistente de IA creado por Anthropic
- Especializado en razonamiento, analisis de datos y codigo
- Modelos disponibles: Claude Haiku (rapido), Sonnet (equilibrado), Opus (potente)
- Diferencia clave: enfoque en seguridad y respuestas utiles
- Desde enero 2026: add-in oficial para Microsoft 365 Pro
- Gratuito para funciones basicas; plan Pro para volumen alto

## Claude vs Copilot: Diferencias y Complementos

- Copilot (Microsoft): integrado nativamente, bueno en tareas rapidas
- Claude (Anthropic): superior en analisis profundo y razonamiento
- Copilot: 'Dame la suma de ventas' -> rapido pero superficial
- Claude: 'Analiza tendencias de las ultimas 12 nominas y sugiere ajustes' -> profundo
- No son competencia: son complementos
- Copilot para automatizacion rapida + Claude para analisis complejo
- Ambos funcionan con lenguaje natural en espaniol

## Instalacion del Add-in desde Microsoft Marketplace

- Requisito: Microsoft 365 Pro (suscripcion activa)
- Pasos:
-    1. Abrir Excel -> Insertar -> Obtener complementos
-    2. Buscar 'Claude for Excel' en el Marketplace
-    3. Clic en Agregar -> Aceptar permisos
-    4. Aparece panel lateral 'Claude' en la pestana Inicio
- Configuracion inicial: ingresar con cuenta Anthropic o M365
- Verificar: escribir en el panel '=Hola, Claude' -> debe responder

## Caso 1: Analizar Tabla de Nomina con Lenguaje Natural

- Seleccionar tabla de nomina (Ctrl+T si no es tabla)
- En panel Claude: 'Analiza esta tabla de nomina e identifica:'
-    - Empleados con sueldo mayor al promedio
-    - Departamentos con mayor gasto
-    - Anomalias en deducciones
- Claude responde con analisis + sugerencias
- Puede generar resumen ejecutivo directamente en nueva hoja
- Ventaja: analisis que tomaria 30 min, en 30 segundos

## Caso 2: Generar Formulas Complejas

- Prompt: 'Necesito una formula que busque el RFC en la tabla Proveedores'
- Claude genera: =SIERROR(BUSCARV(A2,Proveedores,3,0),"No encontrado")
- Formulas anidadas que tomarian 15 min de armar manualmente
- Claude explica cada parte de la formula
- Soporta: BUSCARV, INDICE/COINCIDIR, SUMAR.SI.CONJUNTO, LET
- Tip: pedir 'con SIERROR para manejar errores'
- La formula se puede insertar directamente en la celda seleccionada

## Caso 3: Explicacion de Formula Heredada

- Escenario: reciben archivo de contador anterior con formulas crípticas
- Seleccionar celda con formula compleja
- Prompt: 'Explicame que hace esta formula paso a paso'
- Claude descompone cada funcion anidada
- Identifica posibles errores o mejoras
- Tip: 'Reescribe esta formula de forma mas legible'
- Nunca mas tener miedo de archivos heredados

## MCP Connectors: Fuentes Externas

- MCP = Model Context Protocol (protocolo de Anthropic)
- Permite a Claude conectarse a fuentes externas de datos
- Ejemplos de conectores:
-    - Base de datos SQL del sistema contable
-    - Archivos en SharePoint / OneDrive
-    - APIs del SAT (via conector personalizado)
- Flujo: Claude lee datos externos -> los analiza -> responde en Excel
- Configuracion por administrador de TI
- Potencial: Excel como hub central conectado a todo

## Claude Code: Automatizacion desde Terminal

- Claude Code = Claude en la linea de comandos (terminal)
- Para usuarios avanzados: automatizar procesamiento de archivos Excel
- Ejemplo: 'Abre todos los .xlsx de esta carpeta y consolida ventas'
- Genera scripts Python (openpyxl, pandas) automaticamente
- Ideal para: procesamiento batch, ETL contable, reportes automaticos
- Instalacion: npm install -g @anthropic-ai/claude-code
- Complementa el add-in: lo que no se puede en Excel, se hace en terminal

## Mejores Practicas: Validacion y Privacidad

- SIEMPRE validar las respuestas de Claude (la IA puede equivocarse)
- Datos sensibles: verificar politica de privacidad de su organizacion
- Claude NO almacena datos de las hojas de calculo procesadas
- Para datos altamente confidenciales: usar Claude en modo local/privado
- No depender 100%% de la IA: el criterio contable es insustituible
- Documentar prompts utiles para reutilizarlos (crear biblioteca de prompts)
- Capacitar al equipo: la IA amplifica la productividad de todos

## El Futuro del Contador: IA como Herramienta, Criterio Humano como Brujula

- La IA NO reemplaza contadores; reemplaza tareas repetitivas
- El contador del futuro: analista + estratega + verificador
- Habilidades clave: saber preguntar (prompting) + saber validar
- Excel + IA = superpoder contable
- Lo que la IA no puede hacer: juicio profesional, etica, relaciones
- Su ventaja competitiva: dominar ambos mundos
- El mejor momento para aprender IA fue ayer; el segundo mejor es hoy

## Recursos y Siguiente Paso

- Add-in Claude para Excel: Marketplace de Microsoft 365
- Documentacion Claude: docs.anthropic.com
- Claude Code: npm install -g @anthropic-ai/claude-code
- MCP Protocol: modelcontextprotocol.io
- Practica: instalar el add-in y probar los 3 casos de esta sesion
- Comunidad: todoconta.com para preguntas y recursos adicionales

*Excel para Contadores y Administrativos — Israel Castro*
