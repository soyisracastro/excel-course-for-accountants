# Referencia - Modulo 1

**Modulo 1 de 5**

*Logica Contable y Funciones de Control*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Tarjetas de Funciones Esenciales

Cada tarjeta resume la sintaxis, descripcion y un ejemplo contable de las funciones cubiertas en el Modulo 1.


### SUMA

| Sintaxis | =SUMA(numero1, [numero2], ...) |
| --- | --- |
| Descripcion | Suma todos los numeros en un rango o lista de argumentos. Ignora celdas con texto o vacias. |
| Ejemplo contable | =SUMA(B2:B13) — Suma las ventas mensuales de enero a diciembre para obtener el total anual. |


### PROMEDIO

| Sintaxis | =PROMEDIO(numero1, [numero2], ...) |
| --- | --- |
| Descripcion | Calcula la media aritmetica de los valores. Ignora celdas vacias pero NO ignora ceros. |
| Ejemplo contable | =PROMEDIO(C2:C13) — Gasto promedio mensual de un departamento para presupuesto del siguiente ejercicio. |


### TRUNCAR

| Sintaxis | =TRUNCAR(numero, num_decimales) |
| --- | --- |
| Descripcion | Corta un numero al numero de decimales indicado SIN redondear. Obligatorio para Factor de Actualizacion (Art. 17-A CFF, 4 decimales). |
| Ejemplo contable | =TRUNCAR(141.200/136.163, 4) — Calcula el Factor de Actualizacion truncado al diezmilesimo: 1.0369. |


### SI

| Sintaxis | =SI(prueba_logica, valor_verdadero, valor_falso) |
| --- | --- |
| Descripcion | Evalua una condicion y devuelve un valor si es verdadera y otro si es falsa. Se puede anidar hasta 7 niveles (recomendado: 3). |
| Ejemplo contable | =SI(C2<0, "VENCIDO", SI(C2<=15, "URGENTE", "OK")) — Semaforo de vencimientos fiscales. |


### BUSCARV

| Sintaxis | =BUSCARV(valor_buscado, tabla, num_columna, [ordenado]) |
| --- | --- |
| Descripcion | Busca un valor en la primera columna de una tabla y devuelve el valor de otra columna en la misma fila. Usa VERDADERO para rangos (tarifas) y FALSO para coincidencia exacta (catalogos). |
| Ejemplo contable | =BUSCARV(B4, Tarifa_Anual_2026, 3, VERDADERO) — Obtiene la cuota fija del rango ISR correspondiente al ingreso. |


### HOY

| Sintaxis | =HOY() |
| --- | --- |
| Descripcion | Devuelve la fecha actual del sistema. No requiere argumentos. Se actualiza cada vez que se recalcula la hoja. |
| Ejemplo contable | =DIAS("17/04/2026", HOY()) — Calcula los dias restantes para la fecha limite de la declaracion anual. |


### FECHA

| Sintaxis | =FECHA(anio, mes, dia) |
| --- | --- |
| Descripcion | Construye una fecha a partir de tres componentes separados. Util para armar fechas desde datos dispersos o extraidos con EXTRAE. |
| Ejemplo contable | =FECHA(1985, 3, 15) — Construye la fecha 15/03/1985 a partir de los datos extraidos del RFC CAST850315HN7. |


### EXTRAE

| Sintaxis | =EXTRAE(texto, posicion_inicial, num_caracteres) |
| --- | --- |
| Descripcion | Extrae un numero determinado de caracteres de una cadena de texto, empezando en la posicion indicada. Devuelve texto (multiplicar por 1 para convertir a numero). |
| Ejemplo contable | =EXTRAE("CAST850315HN7", 5, 2) — Extrae "85" (anio de nacimiento) del RFC de una persona fisica. |


---

## Tarifa ISR Anual 2026 — Art. 152 LISR (Anexo 8 RMF)

Tarifa actualizada por inflacion acumulada >10% desde noviembre 2022. Publicada en el DOF el 28 de diciembre de 2025.


| Limite Inferior | Limite Superior | Cuota Fija | % Excedente |
| --- | --- | --- | --- |
| $0.01 | $10,135.11 | $0.00 | 1.92% |
| $10,135.12 | $86,022.11 | $194.59 | 6.40% |
| $86,022.12 | $151,176.19 | $5,051.37 | 10.88% |
| $151,176.20 | $175,735.66 | $12,140.13 | 16.00% |
| $175,735.67 | $210,403.69 | $16,069.64 | 17.92% |
| $210,403.70 | $424,353.97 | $22,282.14 | 21.36% |
| $424,353.98 | $668,840.14 | $67,981.92 | 23.52% |
| $668,840.15 | $1,276,925.98 | $125,485.07 | 30.00% |
| $1,276,925.99 | $1,702,567.97 | $307,910.81 | 32.00% |
| $1,702,567.98 | $5,107,703.92 | $444,116.23 | 34.00% |
| $5,107,703.93 | En adelante | $1,601,862.46 | 35.00% |


**Procedimiento de calculo:** (1) Ubicar la base gravable en la tarifa. (2) Restar el limite inferior. (3) Multiplicar el excedente por el porcentaje. (4) Sumar la cuota fija. Resultado = ISR del ejercicio.


---

## Guia: Factor de Actualizacion — Art. 17-A CFF

El Factor de Actualizacion ajusta montos fiscales por el efecto de la inflacion. Se calcula con el Indice Nacional de Precios al Consumidor (INPC) publicado por el INEGI.


### Paso 1: Identificar los periodos

Determina el mes mas reciente del periodo de actualizacion (INPC reciente) y el mes mas antiguo (INPC anterior). Ejemplo: actualizacion de diciembre 2024 a diciembre 2025.

### Paso 2: Obtener los valores del INPC

Consulta los valores en la pagina del INEGI o en el DOF. INPC Dic 2025 = 141.200 (estimado). INPC Dic 2024 = 136.163.

### Paso 3: Dividir INPC reciente entre INPC anterior

Factor = 141.200 / 136.163 = 1.036996...

### Paso 4: Truncar a 4 decimales (diezmilesimo)

El Art. 17-A CFF establece que el factor se trunca, NO se redondea. En Excel: =TRUNCAR(141.200/136.163, 4) = **1.0369**

### Paso 5: Aplicar el factor al monto original

Monto actualizado = Monto original x Factor. Ejemplo: $100,000.00 x 1.0369 = **$103,690.00**


| Concepto | Valor | Formula Excel |
| --- | --- | --- |
| INPC reciente (Dic 2025) | 141.200 | Celda B4 |
| INPC anterior (Dic 2024) | 136.163 | Celda B5 |
| Factor sin truncar | 1.036996... | =B4/B5 |
| Factor truncado (CFF) | 1.0369 | =TRUNCAR(B4/B5, 4) |
| Monto original | $100,000.00 | Celda B9 |
| Monto actualizado | $103,690.00 | =B9*B7 |

---

## Ejercicios de Calculo ISR 2026 con Respuestas

Para cada escenario, calcula el ISR anual 2026 usando la tarifa del Art. 152 LISR. Las respuestas incluyen el desglose paso a paso.


### Ejercicio 1: Empleado con sueldo fijo

| Concepto | Monto |
| --- | --- |
| Ingreso anual | $280,000.00 |
| Deducciones autorizadas | $45,000.00 |
| Base gravable | $235,000.00 |
| Limite inferior | $210,403.70 |
| Excedente | $24,596.30 |
| % sobre excedente | 21.36% |
| ISR marginal | $5,253.77 |
| Cuota fija | $22,282.14 |
| ISR del ejercicio | $27,535.91 |


### Ejercicio 2: Freelancer con ingresos variables

| Concepto | Monto |
| --- | --- |
| Ingreso anual | $520,000.00 |
| Deducciones autorizadas | $120,000.00 |
| Base gravable | $400,000.00 |
| Limite inferior | $210,403.70 |
| Excedente | $189,596.30 |
| % sobre excedente | 21.36% |
| ISR marginal | $40,497.77 |
| Cuota fija | $22,282.14 |
| ISR del ejercicio | $62,779.91 |


### Ejercicio 3: Socio de empresa (dividendos)

| Concepto | Monto |
| --- | --- |
| Ingreso anual | $1,500,000.00 |
| Deducciones autorizadas | $350,000.00 |
| Base gravable | $1,150,000.00 |
| Limite inferior | $668,840.15 |
| Excedente | $481,159.85 |
| % sobre excedente | 30.00% |
| ISR marginal | $144,347.95 |
| Cuota fija | $125,485.07 |
| ISR del ejercicio | $269,833.03 |


### Ejercicio 4: Trabajador zona fronteriza

| Concepto | Monto |
| --- | --- |
| Ingreso anual | $180,000.00 |
| Deducciones autorizadas | $30,000.00 |
| Base gravable | $150,000.00 |
| Limite inferior | $86,022.12 |
| Excedente | $63,977.88 |
| % sobre excedente | 10.88% |
| ISR marginal | $6,960.79 |
| Cuota fija | $5,051.37 |
| ISR del ejercicio | $12,012.16 |


### Ejercicio 5: Director general (ingreso alto)

| Concepto | Monto |
| --- | --- |
| Ingreso anual | $3,200,000.00 |
| Deducciones autorizadas | $600,000.00 |
| Base gravable | $2,600,000.00 |
| Limite inferior | $1,702,567.98 |
| Excedente | $897,432.02 |
| % sobre excedente | 34.00% |
| ISR marginal | $305,126.89 |
| Cuota fija | $444,116.23 |
| ISR del ejercicio | $749,243.12 |


---

## Atajos de Teclado del Modulo 1

Dominar estos atajos te ahorrara minutos cada dia y horas cada mes.


| Atajo | Accion | Cuando usarlo |
| --- | --- | --- |
| F2 | Entrar a modo edicion de celda | Para ver y editar formulas (debuggear) |
| Ctrl + * | Seleccionar toda la region de datos | Para seleccionar una tabla completa rapido |
| Tab | Autocompletar funcion sugerida | Cuando escribes =SU... y aparece SUMA |
| Ctrl + Z | Deshacer la ultima accion | Cuando cometes un error (funciona multiples veces) |
| Ctrl + Y | Rehacer (repetir ultima accion) | Para re-aplicar algo que deshiciste |
| Alt + = | Insertar SUMA automaticamente | Debajo de una columna de numeros |
| Ctrl + F3 | Administrador de nombres de rango | Para nombrar rangos (VentasEnero, Tarifa, etc.) |
| F4 | Alternar referencia absoluta ($) | Para fijar celdas en formulas antes de arrastrar |
| Ctrl + ` | Mostrar/ocultar formulas en la hoja | Para revisar todas las formulas de un vistazo |
| Ctrl + Shift + L | Activar/desactivar filtros | Para filtrar datos en una tabla rapidamente |


**Tip:** Practica un atajo nuevo cada dia durante una semana. Al final del curso tendras mas de 40 atajos en tu memoria muscular.
