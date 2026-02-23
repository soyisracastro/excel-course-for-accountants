# Guia de Prompts para Copilot

**Modulo 5 - Automatizacion Nativa con Microsoft 365 Copilot**

*20 prompts listos para usar con datos contables en Excel*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Como usar esta guia

Esta guia contiene 20 prompts disenados especificamente para contadores y administrativos que usan Microsoft 365 Copilot en Excel. Cada prompt esta pensado para trabajar con el archivo **12_Dataset_Master_Copilot.xlsx** incluido en este modulo.


**Requisitos previos:**

- Licencia Microsoft 365 con Copilot habilitado.
- Archivo guardado en OneDrive o SharePoint (obligatorio).
- Datos en formato Tabla con nombre (ya configurado en el archivo).

**Estructura de cada prompt:**

- **Prompt exacto:** Lo que escribiras en el panel de Copilot.
- **Que esperar:** La respuesta esperada de la IA.
- **Como validar:** Como verificar que la respuesta sea correcta con funciones de Excel.

**Recuerda:** Copilot es una herramienta de apoyo, no un sustituto de tu criterio profesional. Siempre valida los resultados.


---

## 1. Analisis de Datos


### Prompt 1


**Prompt exacto:**

```
"Analiza las ventas por sucursal y dime cual tiene mejor desempeno en los ultimos 3 meses."
```


**Que esperar:**

Copilot generara una tabla resumen con totales por sucursal (Centro, Norte, Sur) filtrando Oct-Dic 2025, indicando cual tiene mayor volumen de ventas.


**Como validar:**

Crea una tabla dinamica manual con filtro de fecha Oct-Dic y agrupa por Sucursal. Compara los totales con la respuesta de Copilot.


### Prompt 2


**Prompt exacto:**

```
"Cual vendedor tiene el mejor desempeno en ventas totales? Muestra un ranking de los 5 mejores."
```


**Que esperar:**

Un ranking ordenado de vendedores por Venta_Total acumulada. Vendedor_3 deberia aparecer en los ultimos lugares.


**Como validar:**

Usa SUMAR.SI para acumular Venta_Total por vendedor y ordena de mayor a menor. Verifica que Vendedor_3 este abajo.


### Prompt 3


**Prompt exacto:**

```
"Compara el volumen de litros vendidos por tipo de combustible entre las sucursales."
```


**Que esperar:**

Una tabla cruzada Sucursal vs TipoCombustible mostrando suma de litros. Norte deberia mostrar mayor proporcion de Premium.


**Como validar:**

Crea tabla dinamica con Sucursal en filas, TipoCombustible en columnas y Suma de Litros en valores.


### Prompt 4


**Prompt exacto:**

```
"Cual es la tendencia de ventas mes a mes durante 2025? Hay alguna estacionalidad?"
```


**Que esperar:**

Un analisis temporal con ventas mensuales mostrando si hay meses altos o bajos. Copilot puede identificar tendencias.


**Como validar:**

Agrupa las fechas por mes con una tabla dinamica y grafica la serie temporal. Observa si coincide con el analisis de Copilot.


### Prompt 5


**Prompt exacto:**

```
"Que porcentaje de las ventas se pagan con cada metodo de pago? Desglosalo por sucursal."
```


**Que esperar:**

Porcentajes de Efectivo, Tarjeta y Transferencia por sucursal en formato tabla o grafico.


**Como validar:**

Usa CONTAR.SI.CONJUNTO para contar transacciones por MetodoPago y Sucursal. Calcula los porcentajes manualmente.


---

## 2. Deteccion de Errores y Anomalias


### Prompt 6


**Prompt exacto:**

```
"Identifica anomalias en la tabla de nomina. Hay empleados con cambios inusuales de sueldo?"
```


**Que esperar:**

Copilot deberia detectar los 2 empleados con incrementos subitos de sueldo (empleado 3 en julio, empleado 11 en octubre).


**Como validar:**

Filtra por cada empleado y grafica su SueldoBase por periodo. Los saltos seran visibles como picos en la linea.


### Prompt 7


**Prompt exacto:**

```
"Hay datos faltantes en la nomina? Que empleados tienen meses sin registro?"
```


**Que esperar:**

Deberia identificar al empleado con 3 meses faltantes (abril, mayo, junio 2025).


**Como validar:**

Usa CONTAR.SI para contar registros por empleado. El que tenga menos de 12 registros mensuales tiene meses faltantes.


### Prompt 8


**Prompt exacto:**

```
"Detecta si hay empleados con horas extra inusualmente altas. En que periodos ocurre?"
```


**Que esperar:**

Copilot deberia senalar que diciembre tiene picos de horas extra en todos los empleados.


**Como validar:**

Calcula el promedio de HorasExtra por periodo. Diciembre deberia tener un promedio significativamente mayor.


### Prompt 9


**Prompt exacto:**

```
"Revisa si algun vendedor tiene un rendimiento consistentemente bajo comparado con el promedio."
```


**Que esperar:**

Identificacion de Vendedor_3 como el de menor rendimiento sistematico (transacciones pequenas).


**Como validar:**

Calcula promedio de Litros y Venta_Total por vendedor. Vendedor_3 tendra promedios notablemente menores.


---

## 3. Calculos y Formulas


### Prompt 10


**Prompt exacto:**

```
"Crea una columna que calcule el ISR marginal para cada empleado basado en su TotalPercepcion mensual."
```


**Que esperar:**

Copilot agregara una columna con formula que aplique la tarifa Art. 96 LISR, ubicando el rango y aplicando el porcentaje correspondiente.


**Como validar:**

Compara los valores de la nueva columna con la columna ISR existente. Deben ser iguales o muy cercanos.


### Prompt 11


**Prompt exacto:**

```
"Calcula una comision del 2% sobre Venta_Total para cada vendedor y agregala como nueva columna."
```


**Que esperar:**

Una nueva columna 'Comision' con la formula =Venta_Total*0.02 aplicada a todas las filas.


**Como validar:**

Verifica manualmente: multiplica Venta_Total por 0.02 en algunas filas y compara.


### Prompt 12


**Prompt exacto:**

```
"Agrega una columna que clasifique cada venta como 'Alta' (>$5,000), 'Media' ($1,000-$5,000) o 'Baja' (<$1,000)."
```


**Que esperar:**

Copilot creara una columna con funcion SI anidada o IFS que clasifique por rango de Venta_Total.


**Como validar:**

Filtra por cada categoria y verifica que los montos correspondan a los rangos definidos.


### Prompt 13


**Prompt exacto:**

```
"Calcula el sueldo neto promedio por puesto y ordena de mayor a menor."
```


**Que esperar:**

Un resumen con el promedio de NetoPagar agrupado por Puesto, ordenado descendentemente.


**Como validar:**

Usa PROMEDIO.SI para calcular el promedio de NetoPagar por cada puesto unico.


---

## 4. Graficos y Visualizacion


### Prompt 14


**Prompt exacto:**

```
"Crea un grafico de barras que muestre las ventas totales por mes durante 2025."
```


**Que esperar:**

Un grafico de barras verticales con 12 barras (Ene-Dic) mostrando la suma de Venta_Total por mes.


**Como validar:**

Crea tu propio grafico con tabla dinamica de Fecha (agrupada por mes) vs Suma de Venta_Total.


### Prompt 15


**Prompt exacto:**

```
"Muestra la distribucion de ventas por tipo de combustible con un grafico de pastel."
```


**Que esperar:**

Un grafico circular con 3 segmentos (Magna, Premium, Diesel) mostrando proporcion de ventas.


**Como validar:**

Suma Venta_Total por TipoCombustible y crea un grafico circular manual para comparar.


### Prompt 16


**Prompt exacto:**

```
"Genera un grafico de lineas que muestre la evolucion del sueldo base de los 5 empleados con mayor sueldo."
```


**Que esperar:**

Un grafico de lineas con 5 series temporales mostrando SueldoBase por periodo.


**Como validar:**

Identifica los 5 empleados con mayor SueldoBase y graficalos manualmente con tabla dinamica.


### Prompt 17


**Prompt exacto:**

```
"Crea un grafico comparativo de ventas por turno (Matutino, Vespertino, Nocturno) para cada sucursal."
```


**Que esperar:**

Un grafico de barras agrupadas con 3 grupos (sucursales) y 3 barras cada uno (turnos).


**Como validar:**

Tabla dinamica con Sucursal en filas, Turno en columnas y Suma de Venta_Total en valores.


---

## 5. Automatizacion y Resumen


### Prompt 18


**Prompt exacto:**

```
"Genera un resumen ejecutivo de la tabla de ventas: totales, promedios, mejor sucursal, mejor vendedor y tendencia."
```


**Que esperar:**

Un parrafo o tabla con KPIs principales: venta total, promedio por transaccion, sucursal lider, vendedor estrella.


**Como validar:**

Calcula cada KPI manualmente con funciones SUMA, PROMEDIO, MAX, y verifica que coincidan.


### Prompt 19


**Prompt exacto:**

```
"Crea una tabla de frecuencia que muestre cuantas transacciones hay por rango de litros (0-50, 50-100, 100-200, 200-500)."
```


**Que esperar:**

Una tabla con 4 filas mostrando el conteo de transacciones en cada rango de litros.


**Como validar:**

Usa CONTAR.SI.CONJUNTO con criterios de rango para contar transacciones en cada intervalo.


### Prompt 20


**Prompt exacto:**

```
"Resume las deducciones totales por tipo (ISR, IMSS, Otras) para toda la nomina y calcula el porcentaje que representa cada una."
```


**Que esperar:**

Una tabla resumen con 3 filas: ISR total, IMSS total, OtrasDeducciones total, y su porcentaje del total de deducciones.


**Como validar:**

Suma cada columna de deducciones y calcula el porcentaje de cada una sobre TotalDeduccion.


---

## Consejos para mejores resultados con Copilot

- Se especifico: en lugar de 'analiza los datos', di exactamente que columnas y que operacion.
- Menciona el nombre de la tabla: 'En la tabla Ventas_Gasolinera, calcula...'
- Pide un paso a la vez: no combines multiples solicitudes en un solo prompt.
- Si la respuesta no es correcta, reformula el prompt con mas detalle.
- Usa Copilot para explorar, pero siempre valida con formulas tradicionales.
- Guarda las formulas utiles que Copilot genere para reutilizarlas.

## Limitaciones actuales de Copilot en Excel

- No puede acceder a archivos locales; el archivo debe estar en la nube (OneDrive/SharePoint).
- Solo trabaja con datos en formato Tabla (Ctrl+T).
- Puede generar formulas incorrectas; siempre verifica los resultados.
- No reemplaza el criterio contable profesional (NIF, LISR, CFF).
- Disponibilidad limitada a ciertos planes de Microsoft 365.
- Las respuestas pueden variar si repites el mismo prompt.