# MODULO 4: El Dashboard Inteligente y Entrega Profesional

*ISR + Nomina + Graficos = Dashboard Profesional*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## Que es un Dashboard?

- Un panel de control visual que resume informacion clave en una sola pantalla
- Permite tomar decisiones rapidas sin revisar datos crudos
- Combina: KPIs (numeros clave) + Graficos + Filtros interactivos
- En contabilidad: nomina, ISR, flujo de efectivo, vencimientos
- NO es un reporte completo -- es un resumen ejecutivo visual

## Principios de Diseno: Menos es Mas

- Regla de los 5 segundos: el mensaje principal debe entenderse en 5 segundos
- Maximo 4-6 KPIs visibles (no satures la pantalla)
- Paleta de colores consistente (azul = principal, rojo = alerta, verde = ok)
- Graficos simples: barras para comparar, lineas para tendencias
- Elimina ruido visual: sin bordes de cuadricula, sin decoraciones innecesarias
- Jerarquia visual: KPIs arriba, graficos al centro, detalle abajo

## Segmentadores de Datos (Slicers)

- Filtros visuales e interactivos para Tablas Dinamicas
- El usuario hace clic en botones para filtrar -- sin menus complicados
- Como insertarlos: clic en TD > Insertar > Segmentacion de datos
- Selecciona los campos: Periodo, Puesto, Departamento, etc.
- Se pueden vincular a MULTIPLES Tablas Dinamicas simultaneamente
- El resultado: un dashboard interactivo sin necesidad de macros

## Caso Practico: Slicers Vinculados a TDs

- Paso 1: Crear TD1 con Periodo en filas, Sueldo/ISR en valores
- Paso 2: Crear TD2 con Puesto en filas, conteo de Empleados
- Paso 3: Insertar slicer de 'Periodo' desde TD1
- Paso 4: Clic derecho en slicer > Conexiones de informe > marcar TD2
- Paso 5: Ahora al filtrar por mes, AMBAS tablas se actualizan
- Paso 6: Los graficos basados en las TDs tambien se actualizan automaticamente

## Proyecto Final: La Gran Integracion

- Modulo 1: Calculadora ISR con BUSCARV y tarifa oficial 2026
- Modulo 2: Datos de nomina 20 empleados x 12 meses + Tablas Dinamicas
- Modulo 3: Graficos de barras (comparacion) y lineas (tendencia)
- Modulo 4: TODO junto en un Dashboard con slicers y KPIs
- Archivo de trabajo: 10_Dashboard_Final_Integrado.xlsx
- Resultado: un panel de control de nomina completo y profesional

## KPIs: Los Numeros que Importan

- KPI 1 -- Total Percepciones: =SUBTOTAL(109, Nomina[Sueldo])
- KPI 2 -- Total Deducciones: =SUBTOTAL(109, Nomina[ISR]) + SUBTOTAL(109, Nomina[IMSS])
- KPI 3 -- ISR del Periodo: =SUBTOTAL(109, Nomina[ISR])
- KPI 4 -- e.firma Status: formula SI con dias restantes
- SUBTOTAL(109,...) respeta filtros de Tablas Dinamicas y Slicers
- Alternativa: =GETPIVOTDATA() para extraer valores directamente de TDs

## Construyendo el Layout del Dashboard

- Fila 1: Titulo del dashboard (fuente grande, color azul)
- Filas 2-5: Cuadros de KPI (4 cajas con colores de semaforo)
- Columna A-B: Zona de segmentadores (slicers)
- Filas 7-16: Graficos principales (2 graficos lado a lado)
- Filas 18-25: Tabla de detalle o grafico adicional
- Archivo template: 09_Layout_Dashboard_Contable.xlsx

## Proteccion de Celdas y Hojas

- Todas las celdas vienen bloqueadas por defecto (pero inactivo hasta proteger)
- Paso 1: DESBLOQUEA las celdas de entrada (Formato > Proteccion > desmarcar Bloqueada)
- Paso 2: Resalta las celdas editables con fondo amarillo
- Paso 3: Revisar > Proteger hoja > establece contrasena
- Permite: seleccionar celdas, usar filtros y Tablas Dinamicas
- Resultado: el usuario puede interactuar pero no romper las formulas

## Distribucion: PDF vs Excel Protegido

- PDF: para reportes finales, firma digital, envio a clientes
-    -- No editable, no requiere Excel, archivo ligero
- Excel protegido: para plantillas, calculadoras interactivas
-    -- Mantiene formulas, slicers y TDs activas
- Contrasena de apertura (AES): seguridad real del archivo
- Contrasena de escritura: permite lectura sin edicion
- Marcar como final: solo una sugerencia, no proteccion real

## Demo: Dashboard Final en Accion

- Abre el archivo: 10_Dashboard_Final_Integrado.xlsx
- 1. Revisa la hoja Datos_Nomina (240 registros reales)
- 2. Ve a Calculadora y cambia el sueldo en B4 -- observa el ISR
- 3. Crea una Tabla Dinamica desde Datos_Nomina
- 4. Inserta Segmentadores de Periodo y Puesto
- 5. Crea graficos de barras y lineas
- 6. Mueve todo a la hoja Dashboard y ajusta el layout
- 7. Protege las hojas y guarda como PDF + Excel protegido

## Resumen del Modulo 4

- Un dashboard es un resumen ejecutivo visual -- no un reporte completo
- Los segmentadores (slicers) hacen tu dashboard interactivo sin macros
- KPIs con SUBTOTAL(109,...) respetan filtros automaticamente
- El layout sigue la jerarquia: KPIs > Graficos > Detalle
- Proteccion: desbloquea inputs, bloquea formulas, protege hoja
- Distribuye como PDF (reportes) o Excel protegido (interactivos)
- Este modulo integra TODO: ISR (M1) + Nomina (M2) + Graficos (M3)

## Recursos y Siguiente Paso

- Archivo: 09_Layout_Dashboard_Contable.xlsx (template de layout)
- Archivo: 10_Dashboard_Final_Integrado.xlsx (datos + calculadora + dashboard)
- PDF: 11_Guia_Proteccion_y_Seguridad.pdf (checklist de proteccion)
- PDF: Referencia_Modulo_4.pdf (guia rapida de diseno y slicers)
- Practica: Construye tu propio dashboard con datos reales de tu empresa
- Siguiente: Modulo 5 -- Automatizacion Nativa con Microsoft 365 Copilot

*Excel para Contadores y Administrativos — Israel Castro*
