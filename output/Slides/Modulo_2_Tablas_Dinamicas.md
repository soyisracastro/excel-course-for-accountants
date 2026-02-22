# MÓDULO 2: Procesamiento Masivo y Análisis con Tablas Dinámicas

*De datos crudos a analisis profesional en minutos*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## Agenda del Modulo 2

- Limpieza masiva de datos: tecnicas y herramientas
- Tabla vs Rango: por que SIEMPRE usar Tablas
- Creacion de Tablas Dinamicas paso a paso
- Zonas de campo: Filas, Columnas, Valores, Filtros
- Caso practico: Analisis de nomina desde XML
- Papel de trabajo referenciado con BUSCARV
- Segmentadores (Slicers) para analisis interactivo

## El Problema: Datos Sucios

- El 80% del tiempo de analisis se gasta en LIMPIAR datos
- Problemas comunes: celdas vacias, espacios extra, formatos mixtos
- Texto almacenado como numero (subtotales con '$' y comas)
- Fechas en 5 formatos diferentes en la misma columna
- RFCs incompletos o con espacios al inicio/final
- Filas duplicadas que distorsionan los totales
- Sin limpieza, CUALQUIER analisis estara MAL

## Herramientas de Limpieza en Excel

- Buscar y Reemplazar (Ctrl+H): quitar '$', comas, espacios
- ESPACIOS (TRIM): elimina espacios al inicio, final y dobles
- LIMPIAR (CLEAN): elimina caracteres no imprimibles
- SUSTITUIR (SUBSTITUTE): reemplaza texto especifico
- VALOR (VALUE): convierte texto que parece numero a numero real
- Texto en Columnas: separa datos concatenados
- Quitar Duplicados: elimina filas repetidas (Datos > Quitar duplicados)
- Pegado Especial > Valores: 'aplasta' formulas a valores fijos

## Tabla vs Rango: La Diferencia Clave

- RANGO: celdas sueltas sin estructura (A1:H200)
- TABLA: estructura inteligente con nombre (Ctrl+T)
- Tabla agrega encabezados fijos al hacer scroll
- Tabla expande formulas automaticamente al agregar filas
- Tabla permite referencias estructuradas: =SUMA(Tabla[Ventas])
- Tabla es REQUISITO para Tablas Dinamicas eficientes
- Regla de oro: Si tiene encabezados, CONVIERTELO en Tabla

## Crear una Tabla en 3 Pasos

- 1. Selecciona cualquier celda dentro de tus datos
- 2. Presiona Ctrl+T (o Insertar > Tabla)
- 3. Verifica que 'La tabla tiene encabezados' este marcado
- 
- Excel detecta automaticamente el rango de datos
- Asigna un nombre por defecto (Tabla1, Tabla2...)
- CONSEJO: Renombra la tabla inmediatamente (pestana Diseno de Tabla)
- Nombres descriptivos: nomina_2025, gastos_deducibles, catalogo_cuentas

## Tablas Dinamicas: El Superpoder de Excel

- Resumen miles de filas en segundos
- Reorganiza datos arrastrando campos con el mouse
- Calcula sumas, promedios, conteos sin formulas
- Permite analisis multidimensional (por empleado, por mes, por concepto)
- Se actualizan con un clic derecho > Actualizar
- Son LA herramienta de analisis #1 para contadores

## Crear una Tabla Dinamica

- 1. Haz clic en cualquier celda de tu tabla de datos
- 2. Ve a Insertar > Tabla Dinamica
- 3. Elige ubicacion: Nueva hoja (recomendado) o hoja existente
- 4. Aparece el panel 'Campos de tabla dinamica' a la derecha
- 5. Arrastra campos a las 4 zonas: Filtros, Columnas, Filas, Valores
- 6. Excel calcula automaticamente los totales

## Las 4 Zonas de Campos

- FILTROS: campos para filtrar toda la tabla (ej. Anio, Sucursal)
- COLUMNAS: campos que se muestran como encabezados horizontales
- FILAS: campos que se muestran como etiquetas verticales
- VALORES: campos numericos que se calculan (Suma, Promedio, Cuenta)
- 
- Ejemplo nomina:
-   Filtros: Periodo | Filas: Empleado | Columnas: Concepto | Valores: Suma de Importe

## Caso Practico: Analisis de Nomina XML

- Archivo: 05_Analisis_Nomina_XML_Pivot.xlsx
- 500+ registros de 20 empleados, 12 meses
- Datos simulan extraccion de XML CFDI de nomina del SAT
- Columnas: UUID, Empleado, Puesto, Periodo, Clase, Concepto, Importes
- Clase: Percepcion, Deduccion, OtroPago
- 
- Objetivo: Crear 3 tablas dinamicas para analizar la nomina completa

## Pivot 1: Percepciones por Empleado

- Filtro de Clase: 'Percepcion'
- Filas: NombreEmpleado
- Columnas: Concepto (Sueldo, Prima Vacacional)
- Valores: Suma de ImporteGravado
- 
- Preguntas que responde:
-   Quien gana mas? Cuanto se pago en prima vacacional?
-   Cual es el costo total de percepciones por puesto?

## Pivot 2: Deducciones (ISR + IMSS)

- Filtro de Clase: 'Deduccion'
- Filas: NombreEmpleado + Puesto (jerarquia)
- Columnas: Concepto (ISR, IMSS)
- Valores: Suma de ImporteGravado
- 
- Insight clave: El ISR es proporcional al sueldo
- Director paga ~30% de ISR vs Recepcionista ~6%
- IMSS es relativamente parejo (~2.77% para todos)

## Pivot 3: Costo Total por Empleado

- Sin filtro de Clase (todas las clases)
- Filas: NombreEmpleado
- Columnas: Clase (Percepcion, Deduccion, OtroPago)
- Valores: Suma de ImporteGravado + ImporteExento
- 
- Costo neto = Percepciones - Deducciones + OtrosPagos
- Ideal para presupuestos y analisis de costo de nomina

## Segmentadores: Filtros Visuales

- Insertar > Segmentacion de datos (Slicer)
- Botones visuales para filtrar la tabla dinamica
- Pueden conectar UN slicer a MULTIPLES tablas dinamicas
- Ideal para dashboards interactivos
- 
- Recomendacion para nomina:
-   Slicer de Periodo (filtrar por mes)
-   Slicer de Puesto (filtrar por cargo)
-   Slicer de Clase (cambiar entre percepciones/deducciones)

## Papel de Trabajo Referenciado con Tarifa ISR

- Archivo: 06_Papel_Trabajo_Referenciado.xlsx
- Formato de declaracion anual ISR con formulas automaticas
- Usa IFERROR(BUSCARV(...)) para buscar en tarifa ISR 2026
- Secciones: Ingresos, Deducciones, Base Gravable, Calculo ISR
- Celdas amarillas = captura manual, Verdes = calculo automatico
- 
- VINCULACION: Los totales de tus Pivots alimentan este papel
- Percepciones -> Ingresos | ISR retenido -> Retenciones

## Tips Profesionales para Tablas Dinamicas

- SIEMPRE nombra tus tablas fuente (no dejes 'Tabla1')
- Actualiza pivots despues de cambiar datos: clic derecho > Actualizar
- Usa 'Mostrar valores como' para porcentajes y diferencias
- Agrupa fechas por mes/trimestre/anio automaticamente
- Crea campos calculados para metricas personalizadas
- Formato: ajusta numeros a moneda/porcentaje dentro del pivot
- Un pivot por pregunta: no sobrecargues una sola tabla

## Recursos y Siguiente Paso

- Archivo 04: Practica de limpieza masiva (220 filas con errores)
- Archivo 05: Datos de nomina para crear 3 tablas dinamicas
- Archivo 06: Papel de trabajo ISR con BUSCARV y tarifa 2026
- PDF de Referencia: Modulo 2 (checklist, diagramas, ejercicios)
- Siguiente: Modulo 3 - Visualizacion de Impacto y Reportes Ejecutivos

*Excel para Contadores y Administrativos — Israel Castro*
