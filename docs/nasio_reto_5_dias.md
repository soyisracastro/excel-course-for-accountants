# Reto 5 Dias — Nas.io
# Excel para Contadores y Administrativos

Textos para cada checkpoint (copiar y pegar en el campo
"Tell your participants what they need to do for this checkpoint").

---

## Welcome (antes de iniciar el reto)

Bienvenido al Reto 5 Dias: Excel para Contadores y Administrativos.

Durante los proximos 5 dias vas a transformar tu forma de trabajar en Excel. No vamos a ver teoria ni tutoriales genericos — vamos a construir herramientas reales que puedes usar en tu trabajo desde el dia 1.

**Que necesitas tener listo antes de empezar:**
- Microsoft Excel instalado (2019, 2021, o Microsoft 365). Funciona en Windows o Mac.
- Descarga el Pack de Archivos del curso (lo encuentras en los recursos de este reto). Son 12 archivos .xlsx + guias de referencia en .md que vamos a usar en cada sesion.
- Guarda los archivos en una carpeta facil de encontrar, por ejemplo: Escritorio > Reto_Excel_Contadores.
- Si tienes Microsoft 365 con Copilot, asegurate de tener los archivos en OneDrive para el Dia 5. Si no tienes Copilot, no te preocupes — el 90% del curso funciona con cualquier version de Excel.

**Como funciona el reto:**
- Cada dia tiene una clase en video donde construimos juntos paso a paso.
- Cada dia incluye archivos de practica con datos reales (nomina, compras, ventas de combustible, declaraciones).
- No necesitas experiencia previa en formulas avanzadas. Si sabes abrir Excel y escribir en una celda, estas listo.
- Al final de los 5 dias vas a tener un dashboard profesional, interactivo y protegido, listo para presentar a un cliente o a tu jefe.

**Estructura del reto:**
- **Dia 1:** Funciones inteligentes (ISR, vencimientos, RFC)
- **Dia 2:** Tablas dinamicas y procesamiento masivo de datos
- **Dia 3:** Graficos profesionales y reportes ejecutivos
- **Dia 4:** Dashboard integrado + proteccion y entrega profesional
- **Dia 5:** Automatizacion con inteligencia artificial (Copilot y mas)

**Tip:** Descarga los archivos y abrelos antes de la primera sesion para que estes listo desde el minuto uno.

Nos vemos en el Dia 1.

---

## Dia 1: Logica Contable y Funciones de Control

Hoy Excel deja de ser una calculadora gigante y se convierte en tu asistente fiscal inteligente.

**Que vamos a hacer:**
- Construir una calculadora de ISR 2026 que se actualiza sola: cambias el ingreso y todo se recalcula al instante (limite inferior, cuota fija, tasa marginal, impuesto total).
- Entender la diferencia entre TRUNCAR y REDONDEAR — y por que el SAT te la rechaza si usas la funcion equivocada (Art. 17-A del CFF, factor de actualizacion al diezmilesimo).
- Crear un semaforo de vencimientos para tu e.firma y sellos digitales: rojo = vencido, amarillo = urgente, verde = al corriente. Nunca mas se te pasa una fecha.
- Extraer la fecha de nacimiento escondida dentro de cualquier RFC con una sola formula.

**Archivos que vas a usar:**
- Calculadora ISR V2026 (ya tiene la tarifa oficial Art. 96 LISR precargada)
- Control de Vencimientos e.firma
- Extraccion RFC Master

**Funciones clave:** SUMA, PROMEDIO, SI, BUSCARV, TRUNCAR, REDONDEAR, EXTRAE, HOY, FECHA.

**Al terminar el dia** vas a tener 3 archivos funcionales que puedes usar tal cual en tu trabajo real desde manana.

---

## Dia 2: Procesamiento Masivo con Tablas Dinamicas

Hoy vamos a procesar cientos de registros en segundos — sin escribir una sola formula.

**Que vamos a hacer:**
- Limpiar datos sucios (que es la realidad del 80% de los archivos que te llegan): quitar signos de peso pegados, espacios invisibles, fechas en formato texto, RFCs con errores de captura.
- Convertir rangos tontos en Tablas inteligentes (Ctrl+T) que se expanden solas, tienen nombres y hacen referencias estructuradas.
- Crear 3 tablas dinamicas desde un archivo real de nomina XML con 500+ registros (20 empleados x 12 meses):
  1. Percepciones por empleado (sueldo, prima vacacional, aguinaldo)
  2. Deducciones desglosadas (ISR retenido, IMSS, otras)
  3. Costo total integrado por trabajador
- Conectar segmentadores (slicers) para filtrar las 3 tablas con un solo clic.
- Vincular los resultados a un papel de trabajo referenciado con la tarifa de ISR.

**Archivos que vas a usar:**
- Limpieza Masiva Layout (220 filas con errores intencionales para practicar)
- Analisis Nomina XML Pivot (500+ registros reales de nomina)
- Papel de Trabajo Referenciado

**Dato revelador:** Un director que gana $50,000/mes paga 30% de ISR. Una recepcionista paga 6%. Lo vas a ver en segundos con tu tabla dinamica, sin calcularlo manualmente.

**Al terminar el dia** vas a poder tomar cualquier archivo de datos masivo y convertirlo en resumenes ejecutivos arrastrando campos. Cero formulas.

---

## Dia 3: Visualizacion de Impacto y Reportes Ejecutivos

Hoy tus numeros van a hablar por si solos — con graficos que se explican en 5 segundos.

**Que vamos a hacer:**
- Aprender a elegir el grafico correcto (arbol de decision): barras para comparar, lineas para tendencia, pastel para composicion, columnas apiladas para composicion + tendencia.
- Construir un dashboard de ventas de combustible (Magna, Premium, Diesel) con graficos dinamicos que se actualizan con filtros.
- Crear una comparativa anual de estado de resultados (2024 vs 2025): ingresos y gastos lado a lado para detectar si los gastos crecen mas rapido que las ventas.
- Aplicar las reglas de diseno profesional:
  - Colores con significado (verde = Magna, rojo = Premium, gris = Diesel)
  - Eliminar ruido visual (gridlines, leyendas innecesarias, efectos 3D)
  - Etiquetas que ahorran preguntas
  - "Menos es mas" (principio de Edward Tufte: maximizar la tinta de datos)

**Archivos que vas a usar:**
- Dashboard Ventas Combustible (12 meses x 3 tipos de combustible)
- Comparativa Anual Ventas y Gastos (2024 vs 2025)

**Por que importa:** El cerebro procesa imagenes 60,000 veces mas rapido que texto (MIT). Un grafico bien hecho reemplaza 5 paginas de numeros. Tu director o tu cliente lo va a entender en 5 segundos.

**Al terminar el dia** vas a tener 2 reportes visuales de calidad profesional, listos para presentar a un cliente o a direccion.

---

## Dia 4: El Dashboard Inteligente y Entrega Profesional

Hoy integramos todo: funciones + datos + graficos = un dashboard ejecutivo interactivo.

**Que vamos a hacer:**
- Construir un dashboard completo de nomina con 240 registros reales:
  - 4 KPIs en la parte superior (Total Percepciones, Total Deducciones, ISR del Periodo, Estatus e.firma)
  - Graficos dinamicos que se actualizan con segmentadores
  - Filtros por periodo, puesto y departamento — un clic cambia todo
- Dominar SUBTOTAL(109,...) para que tus KPIs respeten los filtros activos (no como SUMA que siempre suma todo).
- Conectar un solo segmentador a multiples tablas dinamicas para que todo se sincronice.
- Proteger tu trabajo como profesional:
  - Bloquear celdas con formulas para que nadie las rompa
  - Desbloquear solo las celdas de captura (marcadas en amarillo)
  - Proteger la hoja con contrasena pero permitir usar filtros y segmentadores
- Estrategia de distribucion: cuando enviar PDF (reportes finales, archivo) vs Excel protegido (cuando el usuario necesita interactuar).

**Archivos que vas a usar:**
- Layout Dashboard Contable (plantilla con estructura lista para llenar)
- Dashboard Final Integrado (version terminada con 4 hojas: Datos, Tarifa ISR, Calculadora, Dashboard)

**Este es el modulo de integracion:** todo lo que aprendiste en los dias 1, 2 y 3 se conecta aqui en un solo producto terminado.

**Al terminar el dia** vas a tener un dashboard profesional, protegido, interactivo y listo para entregar a tu jefe o a tu cliente. Lo exportas como PDF o lo compartes como Excel protegido.

---

## Dia 5: Automatizacion con IA (Microsoft 365 Copilot y mas)

Hoy descubrimos como la inteligencia artificial acelera tu trabajo — sin reemplazar tu criterio profesional.

**Que vamos a hacer:**
- Configurar Copilot en Excel: requisitos reales (licencia M365 + archivo en OneDrive + datos en formato Tabla).
- Probar 5 casos de uso practicos:
  1. Analisis con lenguaje natural: "Analiza las ventas por sucursal y dime cual tiene mejor desempeno" — respuesta en 15 segundos.
  2. Generacion de formulas: "Calcula el ISR marginal basado en la percepcion total" — Copilot genera la formula completa.
  3. Columnas inteligentes: "Clasifica cada venta como Alta, Media o Baja" — se agrega la columna con logica SI/IFS automatica.
  4. Graficos instantaneos: "Crea un grafico de barras de ventas por mes" — visualizacion en un clic.
  5. Deteccion de anomalias: "Identifica datos inusuales en la nomina" — encuentra empleados con saltos de sueldo sospechosos.
- Caso practico completo en 3 minutos: identificar al vendedor con peor desempeno, generar grafico comparativo, y obtener resumen ejecutivo con recomendaciones. Lo que antes tomaba 30 minutos.
- Conocer alternativas: ChatGPT para generar macros VBA, Claude para analisis de documentos, Gemini en Google Sheets.
- Entender los limites: Copilot NO conoce el codigo fiscal mexicano en detalle, NO firma declaraciones, y puede generar formulas con errores sutiles. Siempre valida.

**Archivos que vas a usar:**
- Dataset Master Copilot (1,250 transacciones de gasolinera + nomina con anomalias intencionales)
- Guia de Prompts para Copilot (20 prompts listos para usar en 5 categorias)

**Regla de oro:** Copilot es una herramienta para acelerar el trabajo mecanico. El criterio contable es tuyo y solo tuyo.

**Al terminar el dia** vas a saber exactamente cuando SI y cuando NO usar IA en tu trabajo contable, y vas a tener 20 prompts probados que puedes usar desde manana con tus propios datos.

---

## Notas para la configuracion en Nas.io

- **Nombre del reto:** Reto 5 Dias: Excel para Contadores
- **Formato:** Un checkpoint por dia, cada uno con su clase en video
- **Material descargable:** Subir el Pack Excel Pro completo (12 archivos .xlsx + guias .md) como recurso del reto
- **Bonus adicional:** Las guias de VBA con IA, Claude en Excel, y el Cheat Sheet de 80+ atajos se pueden ofrecer como material extra al completar los 5 dias
