# Automatizacion con Macros Asistidas por IA

*Bonus 1 — VBA + ChatGPT / Claude para contadores*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## Que es VBA? Cuando Macros vs Formulas

- VBA = Visual Basic for Applications (lenguaje de macros de Office)
- Las formulas resuelven calculos; las macros automatizan procesos
- Ejemplo formula: =BUSCARV(A2, Tabla, 3, 0) -> busca un dato
- Ejemplo macro: FormatearNomina() -> aplica formato a 500 filas en 1 clic
- Regla practica: si lo haces mas de 3 veces, automatizalo con macro
- Las macros ahorran horas en tareas repetitivas de cierre mensual

## El Editor VBA (Alt + F11)

- Alt + F11  ->  abre el editor de Visual Basic
- Panel izquierdo: Explorador de proyecto (VBAProject)
- Insertar -> Modulo  ->  aqui se escribe el codigo
- F5 para ejecutar la macro seleccionada
- F8 para ejecutar linea por linea (debug)
- Ventana Inmediato (Ctrl+G): pruebas rapidas
- No se asusten por la interfaz; solo necesitamos Modulo + F5

## Prompting IA para Generar VBA

- La IA (ChatGPT, Claude) genera macros VBA si describes bien la tarea
- Estructura del prompt:
-    1. Contexto: 'Tengo una hoja de nomina en Excel...'
-    2. Tarea: 'Necesito una macro que formatee las columnas A:H...'
-    3. Detalles: 'Separador de miles, bordes, encabezado azul...'
-    4. Restriccion: 'Compatible con Excel 2019 o superior'
- Pedir que incluya comentarios en espaniol
- Siempre probar en copia del archivo, nunca en el original

## Ejemplo 1: Macro FormatearNomina()

- Objetivo: formato profesional en 1 clic
- Aplica separador de miles a columnas monetarias
- Bordes delgados en todo el rango de datos
- Encabezado: fondo azul, texto blanco, negrita
- Autoajuste de ancho de columnas
- Resultado: de datos crudos a tabla presentable en 2 segundos

## Ejemplo 2: Macro LimpiarVacias()

- Problema comun: datos exportados con filas vacias intercaladas
- Solucion manual: seleccionar, eliminar, repetir... (tedioso)
- Macro: recorre de abajo hacia arriba eliminando filas vacias
- Importante: recorrer de abajo hacia arriba para no saltar filas
- Se aplica sobre la seleccion activa del usuario
- Resultado: datos limpios en milisegundos

## Ejemplo 3: Macro ReporteMensual()

- Crea nueva hoja con nombre automatico: 'Reporte_Feb_2026'
- Copia estructura base desde hoja plantilla
- Inserta fecha de generacion automaticamente
- Ideal para reportes mensuales de cierre
- Prompt IA: 'Macro que cree hoja nueva con nombre mes-anio actual'

## Ejemplo 4: Boton Actualizar Tablas Dinamicas

- Problema: 5 tablas dinamicas en el libro, hay que actualizar una por una
- Macro: actualiza TODAS las tablas dinamicas del libro
- Se puede asignar a un boton con forma en la hoja
- Insertar -> Formas -> Rectangulo -> clic derecho -> Asignar macro
- Resultado: un boton 'Actualizar Todo' que ahorra 2 minutos cada vez

## Claude Code como Asistente de Debugging

- Error comun: 'Run-time error 1004' -> referencia invalida
- Copiar el error + el codigo y pegarlo en Claude/ChatGPT
- Claude Code (terminal): analiza archivos .xlsm directamente
- Prompt: 'Este codigo VBA da error 1004 en la linea X. Explicame por que y corrigelo'
- La IA explica el error Y da la correccion
- Ciclo: Generar -> Probar -> Si falla, pedir correccion -> Probar otra vez

## Seguridad: .xlsm y Trust Center

- .xlsx NO soporta macros -> .xlsm SI soporta macros
- Al guardar: Guardar como -> Libro de Excel habilitado para macros (.xlsm)
- Trust Center: Archivo -> Opciones -> Centro de confianza
- Recomendacion: 'Deshabilitar macros con notificacion'
- Nunca habilitar macros de archivos desconocidos
- Firmar macros digitalmente en ambientes corporativos
- Respaldo: siempre mantener copia .xlsx sin macros

## Recursos y Siguiente Paso

- Editor VBA: Alt + F11 en cualquier Excel
- Documentacion VBA: docs.microsoft.com/en-us/office/vba/api/overview/excel
- ChatGPT / Claude para generar y depurar macros
- Practica: crear las 4 macros de esta sesion en un archivo .xlsm
- Tip: guardar macros utiles en un Personal.xlsb para tenerlas siempre disponibles
- Siguiente: Bonus 2 - Claude en Excel: Tu Segundo Cerebro Contable

*Excel para Contadores y Administrativos — Israel Castro*
