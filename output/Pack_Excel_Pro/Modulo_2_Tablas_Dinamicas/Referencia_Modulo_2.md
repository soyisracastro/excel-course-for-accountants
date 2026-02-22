# Guia de Referencia

**Modulo 2: Procesamiento Masivo y Análisis con Tablas Dinámicas**

*Tablas, Limpieza de Datos y Tablas Dinamicas*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Tabla vs Rango: Comparacion

Comprender la diferencia entre un rango de celdas y una Tabla de Excel es fundamental para trabajar de forma eficiente. A continuacion se comparan las caracteristicas principales.


| Caracteristica | Rango (celdas sueltas) | Tabla (Ctrl+T) |
| --- | --- | --- |
| Nombre | Solo referencia A1:H200 | Nombre descriptivo: nomina_2025 |
| Encabezados al scroll | Desaparecen al bajar | Se fijan automaticamente |
| Formulas nuevas filas | Hay que copiar manualmente | Se expanden solas |
| Referencias en formulas | =SUMA(C2:C500) | =SUMA(Tabla[Ventas]) |
| Filtros | Configurar manualmente | Incluidos por defecto |
| Formato alterno | Aplicar a mano | Automatico (filas alternadas) |
| Fila de totales | Crear manualmente | Un clic en Diseno > Fila de totales |
| Fuente para Pivots | Puede perder filas nuevas | Siempre incluye todo |
| Ordenar | Riesgo de desalinear columnas | Seguro: ordena filas completas |


**Regla de oro:** Si tus datos tienen encabezados y mas de una fila, conviertelos en Tabla con Ctrl+T. No hay razon para no hacerlo.

---

## 2. Checklist de Limpieza de Datos

Antes de crear cualquier tabla dinamica o reporte, verifica que tus datos cumplan con estos criterios. Usa esta lista como verificacion rapida.


- Sin filas completamente vacias entre los datos
- Encabezados en la primera fila (sin celdas combinadas en encabezados)
- Una sola fila de encabezados (no dos ni tres)
- Cada columna contiene un solo tipo de dato (no mezclar texto y numeros)
- Fechas en formato fecha real (no texto que parece fecha)
- Montos como numeros reales (sin signos '$' ni comas como texto)
- Sin espacios al inicio o final de textos (usar ESPACIOS/TRIM)
- Sin caracteres no imprimibles (usar LIMPIAR/CLEAN)
- RFCs y claves con longitud correcta (usar LARGO/LEN para verificar)
- Sin filas duplicadas (usar Datos > Quitar duplicados)
- Formatos de fecha consistentes (todos DD/MM/AAAA o todos AAAA-MM-DD)
- Nombres estandarizados (no 'SA de CV' y 'S.A. de C.V.' mezclados)
- Subtotal + IVA = Total (verificar con columna auxiliar)
- Datos convertidos a Tabla (Ctrl+T) antes de crear pivots

| Herramienta | Atajo / Ubicacion | Uso principal |
| --- | --- | --- |
| Buscar y Reemplazar | Ctrl+H | Quitar $, comas, espacios masivos |
| ESPACIOS (TRIM) | =ESPACIOS(celda) | Eliminar espacios extra |
| LIMPIAR (CLEAN) | =LIMPIAR(celda) | Quitar caracteres no imprimibles |
| VALOR (VALUE) | =VALOR(celda) | Convertir texto a numero |
| Texto en Columnas | Datos > Texto en columnas | Separar datos concatenados |
| Quitar Duplicados | Datos > Quitar duplicados | Eliminar filas repetidas |
| Pegado Especial | Ctrl+Alt+V > Valores | Convertir formulas a valores |
| SUSTITUIR | =SUSTITUIR(texto,viejo,nuevo) | Reemplazar texto especifico |

---

## 3. Zonas de Campos de Tabla Dinamica

Al crear una Tabla Dinamica, el panel 'Campos de tabla dinamica' muestra cuatro zonas donde se arrastran los campos. Cada zona tiene una funcion especifica:


| Zona | Ubicacion en el Pivot | Que poner aqui | Ejemplo Nomina |
| --- | --- | --- | --- |
| FILTROS | Arriba del pivot (dropdown) | Campos para restringir el analisis | Periodo, Anio, Sucursal |
| COLUMNAS | Encabezados horizontales | Categorias para desglosar (pocas) | Concepto, Clase |
| FILAS | Etiquetas verticales | El eje principal de analisis | NombreEmpleado, Puesto |
| VALORES | Cuerpo de la tabla (numeros) | Campos numericos a calcular | Suma de ImporteGravado |


### Diagrama visual del panel de campos

El panel de campos tiene dos secciones: arriba la lista de todos los campos disponibles, abajo las cuatro zonas. Simplemente arrastra campos de la lista a la zona deseada.


| CAMPOS DISPONIBLES |
| --- |
| UUID | NumEmpleado | NombreEmpleado | Puesto | FechaPago | Periodo | Clase | Concepto | ImporteGravado | ImporteExento |


| FILTROS | COLUMNAS |
| --- | --- |
| Periodo | Concepto |
| FILAS | VALORES |
| NombreEmpleado | Suma de ImporteGravado |


### Tipos de calculo en Valores

Por defecto, Excel usa SUMA para numeros y CUENTA para texto. Puedes cambiar el tipo de calculo haciendo clic en el campo dentro de la zona Valores > Configuracion de campo de valor.

| Funcion | Cuando usarla | Ejemplo |
| --- | --- | --- |
| Suma | Totalizar montos | Total de percepciones por empleado |
| Cuenta | Contar registros | Numero de conceptos por empleado |
| Promedio | Obtener media | Sueldo promedio por puesto |
| Max / Min | Encontrar extremos | Sueldo mas alto / mas bajo |
| % del total general | Ver proporciones | Que % del ISR paga cada empleado |

---

## 4. Ejercicios Practicos

### Ejercicio 1: Limpieza de datos de compras

**Archivo:** 04_Limpieza_Masiva_Layout.xlsx

**Objetivo:** Limpiar 220 filas de datos de compras con errores intencionales.


**Pasos:**

- Abre la hoja 'Datos_Sucios' y revisa los tipos de errores
- Usa Ctrl+H para eliminar signos '$' en la columna Subtotal
- Aplica ESPACIOS(TRIM) en una columna auxiliar para limpiar RFC
- Convierte las fechas texto a formato fecha con DATEVALUE o Texto en Columnas
- Verifica Subtotal + IVA = Total con una columna de validacion
- Usa Datos > Quitar duplicados en la columna Folio
- Convierte el resultado limpio a Tabla con Ctrl+T
- Compara con la hoja 'Datos_Limpios' para verificar tu trabajo

**Tiempo estimado:** 15-20 minutos. **Reto:** Hazlo en menos de 10 minutos usando atajos.


### Ejercicio 2: Tres Tablas Dinamicas de Nomina

**Archivo:** 05_Analisis_Nomina_XML_Pivot.xlsx

**Objetivo:** Crear 3 tablas dinamicas para analizar percepciones, deducciones y costo total de nomina.


**Pivot 1 - Percepciones:**

- Filtrar Clase = Percepcion
- Filas: NombreEmpleado | Columnas: Concepto | Valores: Suma ImporteGravado
- Agregar Slicer por Periodo

**Pivot 2 - Deducciones:**

- Filtrar Clase = Deduccion
- Filas: NombreEmpleado + Puesto | Columnas: Concepto | Valores: Suma ImporteGravado
- Pregunta: Quien paga mas ISR y por que?

**Pivot 3 - Costo Total:**

- Sin filtro de Clase
- Filas: NombreEmpleado | Columnas: Clase | Valores: Suma ImporteGravado + ImporteExento
- Ordenar de mayor a menor costo total

**Tiempo estimado:** 20-25 minutos.


### Ejercicio 3: Papel de Trabajo ISR con BUSCARV

**Archivo:** 06_Papel_Trabajo_Referenciado.xlsx

**Objetivo:** Usar los resultados de las tablas dinamicas para alimentar el papel de trabajo de declaracion anual ISR.


**Pasos:**

- Del Pivot 1, obtener el total de percepciones gravadas de un empleado
- Capturar ese monto en la celda 'Sueldos y salarios gravados' del papel
- Capturar deducciones personales ficticias (gastos medicos, colegiaturas)
- Observar como las formulas BUSCARV calculan automaticamente el ISR
- Del Pivot 2, obtener el total de ISR retenido del mismo empleado
- Capturarlo en 'Retenciones de ISR en el ejercicio'
- Verificar: si ISR causado < retenciones = saldo a favor (devolucion)

**Tiempo estimado:** 10-15 minutos. **Valor:** Este es el flujo real de trabajo en un despacho contable.

---

## 5. Atajos Clave del Modulo 2

| Atajo | Accion | Cuando usarlo |
| --- | --- | --- |
| Ctrl+T | Convertir rango a Tabla | Siempre que tengas datos con encabezados |
| Ctrl+H | Buscar y Reemplazar | Limpieza masiva de caracteres |
| Ctrl+Shift+L | Activar/desactivar filtros | Filtrar datos rapidamente |
| Alt+N+V | Insertar Tabla Dinamica | Crear un nuevo pivot |
| Alt+F5 | Actualizar Tabla Dinamica | Despues de cambiar datos fuente |
| Ctrl+A | Seleccionar todo | Antes de aplicar formato |
| Ctrl+1 | Formato de celdas | Cambiar formato de numeros/fechas |
| Ctrl+Z | Deshacer | Si algo sale mal (hasta 100 veces) |
| F2 | Editar celda | Ver/modificar formulas |
| Ctrl+` | Mostrar formulas | Ver todas las formulas de la hoja |


## Notas Finales

Las tablas dinamicas son la herramienta de analisis mas poderosa de Excel. Combinadas con datos limpios y bien estructurados en tablas, permiten responder cualquier pregunta de negocio en segundos. Practica con los tres archivos de este modulo hasta que el proceso sea automatico: limpiar, tabular, pivotear, analizar.


En el **Modulo 3** convertiremos estos analisis en graficos profesionales, formato condicional avanzado y reportes ejecutivos listos para presentar.
