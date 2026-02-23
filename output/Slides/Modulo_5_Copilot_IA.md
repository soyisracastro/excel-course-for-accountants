# MODULO 5: Automatizacion Nativa con Microsoft 365 Copilot

*Inteligencia Artificial aplicada a datos contables en Excel*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## El futuro ya esta aqui

- La IA generativa llego a las herramientas de oficina en 2023-2024
- Microsoft 365 Copilot integra GPT-4 directamente en Excel, Word, PowerPoint
- No necesitas saber programar: le hablas en espanol y ejecuta
- El contador que domine estas herramientas tendra ventaja competitiva
- Hoy: 80% de las tareas repetitivas pueden acelerarse con IA

## Requisitos para usar Copilot en Excel

- Licencia Microsoft 365 Business/Enterprise + add-on Copilot ($30 USD/mes)
- El archivo DEBE estar en OneDrive o SharePoint (no funciona en local)
- Los datos DEBEN estar en formato Tabla (Ctrl+T) con nombre
- Conexion a internet activa (Copilot procesa en la nube)
- Funciona en Excel escritorio (Windows/Mac) y Excel Web
- Recomendacion: tablas limpias, sin celdas combinadas, sin filas en blanco

## Activando el panel de Copilot

- Paso 1: Guarda tu archivo en OneDrive (Archivo > Guardar como > OneDrive)
- Paso 2: Selecciona tus datos y convierte a Tabla (Ctrl+T)
- Paso 3: Nombra tu tabla (Diseno de Tabla > Nombre de tabla)
- Paso 4: En la pestana Inicio, clic en el icono de Copilot
- Paso 5: Se abre un panel lateral donde escribes tu solicitud
- Importante: si el boton de Copilot esta gris, verifica los requisitos

## Analisis con lenguaje natural

- Prompt: "Analiza las ventas por sucursal y dime cual tiene mejor desempeno"
- Copilot interpreta tu solicitud y genera tablas resumen automaticamente
- Puede crear calculos que tomarian minutos con tablas dinamicas
- Responde en espanol con explicaciones claras
- Ejemplo real: analisis de 1,250 transacciones en 15 segundos
- Limitacion: a veces interpreta diferente a lo que esperabas

## Generacion de formulas con Copilot

- Prompt: "Calcula el ISR marginal basado en TotalPercepcion"
- Copilot sugiere la formula y la aplica a toda la columna
- Entiende contexto contable mexicano (ISR, IMSS, CFF)
- Genera formulas complejas: BUSCARV, SI anidados, SUMAR.SI.CONJUNTO
- Tu decides si aceptas, modificas o rechazas cada sugerencia
- Siempre revisa la formula antes de aceptar: la IA no es infalible

## Columnas inteligentes

- Prompt: "Agrega columna que clasifique ventas como Alta, Media o Baja"
- Copilot crea la columna con formula y la nombra automaticamente
- Puede crear multiples columnas en secuencia
- Ejemplo: "Ahora agrega una columna de comision del 2%"
- Las columnas se integran a la Tabla existente
- Puedes pedir que elimine o modifique columnas creadas

## Visualizacion instantanea

- Prompt: "Crea grafico de barras de ventas por mes"
- Copilot genera graficos directamente en la hoja de calculo
- Tipos: barras, lineas, pastel, dispersion, combinados
- Ajusta colores y etiquetas automaticamente
- Prompt: "Muestra distribucion por tipo de combustible con grafico de pastel"
- Limitacion: los graficos generados son basicos; para reportes ejecutivos, ajustalos tu

## Deteccion de anomalias (Insights)

- Prompt: "Identifica anomalias en la tabla de nomina"
- Copilot detecta: valores atipicos, cambios subitos, datos faltantes
- Ejemplo: identifica empleados con incrementos inusuales de sueldo
- Ejemplo: detecta meses sin registros para un empleado
- Util para auditorias internas y revision de nomina
- No reemplaza una auditoria formal, pero acelera la deteccion inicial

## Caso practico: 3 prompts contables en secuencia

- Prompt 1: "Cual vendedor tiene el peor desempeno en ventas?"
-    -> Copilot identifica a Vendedor_3 con datos de soporte
- Prompt 2: "Crea grafico comparando ventas de ese vendedor vs el promedio"
-    -> Grafico de barras comparativo generado automaticamente
- Prompt 3: "Genera resumen ejecutivo con recomendacion"
-    -> Resumen con KPIs y sugerencia de accion
- En 3 minutos tienes un mini-analisis que antes tomaba 30 min

## Limitaciones: "El criterio contable es tuyo"

- Copilot NO conoce las NIF, LISR ni CFF en detalle
- Puede generar formulas incorrectas que parecen correctas
- No tiene acceso a tu contexto fiscal especifico (regimen, obligaciones)
- Los calculos de ISR pueden tener errores en rangos o porcentajes
- No puede firmar declaraciones ni sustituir al contador
- Regla de oro: usa Copilot para acelerar, pero SIEMPRE valida
- Tu criterio profesional es insustituible

## IA Externa: ChatGPT, Gemini, Claude

- No todo es Microsoft Copilot: hay alternativas poderosas
- ChatGPT (OpenAI): excelente para generar macros VBA y Power Query M
- Google Gemini: integrado en Google Sheets, buena alternativa gratuita
- Claude (Anthropic): muy preciso en analisis de texto y documentos
- Caso: copia tu formula o tabla y pegala en ChatGPT para que la explique
- Caso: pide a Claude que te genere una macro para automatizar reportes
- Cualquier IA complementa tu trabajo, pero ninguna lo reemplaza

## Caso: Macro sencilla generada por IA

- Prompt a ChatGPT: 'Genera macro VBA que formatee mi reporte mensual'
- La IA genera codigo VBA funcional en segundos
- Tu lo pegas en el Editor de VBA (Alt+F11) y lo ejecutas
- Ejemplo: macro que aplica formato profesional a todas las hojas
- Ejemplo: macro que exporta cada hoja como PDF individual
- Ejemplo: macro que consolida datos de multiples archivos
- Consejo: pide a la IA que agregue comentarios al codigo para entenderlo

## Resumen del curso completo: M1 a M5

- M1: Logica Contable y Funciones (BUSCARV, SI, ISR, Factor de Actualizacion)
- M2: Tablas Dinamicas y Procesamiento Masivo (analisis de datos reales)
- M3: Visualizacion Profesional (graficos ejecutivos, reportes de impacto)
- M4: Dashboard Inteligente (KPIs, interactividad, entrega profesional)
- M5: Copilot e IA (automatizacion con lenguaje natural, herramientas externas)
- Juntos: del dato crudo al insight accionable, con herramientas modernas

## Comunidad y siguiente paso

- Unete a la comunidad de contadores que usan Excel de forma avanzada
- Practica con los archivos del Pack Excel Pro incluidos en cada modulo
- Comparte tus logros: sube un antes/despues de tu flujo de trabajo
- Mantente actualizado: Microsoft actualiza Copilot constantemente
- Siguiente nivel: Power Query, Power BI, automatizacion con Power Automate
- Tu inversion en aprendizaje hoy se traduce en eficiencia manana

## Recursos y Siguiente Paso

- Pack Excel Pro: archivos de practica para los 5 modulos
- Guia de Prompts para Copilot: 20 prompts listos para usar
- Referencia Modulo 5: checklist, prompts y resumen del curso
- Microsoft Learn: learn.microsoft.com/copilot
- OpenAI ChatGPT: chat.openai.com
- Anthropic Claude: claude.ai
- todoconta.com: recursos adicionales y actualizaciones

*Excel para Contadores y Administrativos — Israel Castro*
