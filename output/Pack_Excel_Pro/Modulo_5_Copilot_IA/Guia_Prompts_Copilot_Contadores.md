# Guía de Prompts para Copilot

**Módulo 5 — Automatización Nativa con Microsoft 365 Copilot**

*20 prompts listos para usar con datos contables en Excel*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Cómo usar esta guía

Esta guía contiene 20 prompts diseñados específicamente para contadores y administrativos que usan Microsoft 365 Copilot en Excel. Cada prompt está pensado para trabajar con el archivo **12_Dataset_Master_Copilot.xlsx** incluido en este módulo.


**Requisitos previos:**

- Licencia Microsoft 365 con Copilot habilitado.
- Archivo guardado en OneDrive o SharePoint (obligatorio).
- Datos en formato Tabla con nombre (ya configurado en el archivo).

**Estructura de cada prompt:**

- **Prompt exacto:** Lo que escribirás en el panel de Copilot.
- **Qué esperar:** La respuesta esperada de la IA.
- **Cómo validar:** Cómo verificar que la respuesta sea correcta con funciones de Excel.

**Recuerda:** Copilot es una herramienta de apoyo, no un sustituto de tu criterio profesional. Siempre valida los resultados.


---

## 1. Análisis de Datos


### Prompt 1


**Prompt exacto:**

```
"Analiza las ventas por sucursal y dime cuál tiene mejor desempeño en los últimos 3 meses."
```


**Qué esperar:**

Copilot generará una tabla resumen con totales por sucursal (Centro, Norte, Sur) filtrando Oct-Dic 2025, indicando cuál tiene mayor volumen de ventas.


**Cómo validar:**

Crea una tabla dinámica manual con filtro de fecha Oct-Dic y agrupa por Sucursal. Compara los totales con la respuesta de Copilot.


### Prompt 2


**Prompt exacto:**

```
"¿Cuál vendedor tiene el mejor desempeño en ventas totales? Muestra un ranking de los 5 mejores."
```


**Qué esperar:**

Un ranking ordenado de vendedores por Venta_Total acumulada. Vendedor_3 debería aparecer en los últimos lugares.


**Cómo validar:**

Usa SUMAR.SI para acumular Venta_Total por vendedor y ordena de mayor a menor. Verifica que Vendedor_3 esté abajo.


### Prompt 3


**Prompt exacto:**

```
"Compara el volumen de litros vendidos por tipo de combustible entre las sucursales."
```


**Qué esperar:**

Una tabla cruzada Sucursal vs TipoCombustible mostrando suma de litros. Norte debería mostrar mayor proporción de Premium.


**Cómo validar:**

Crea tabla dinámica con Sucursal en filas, TipoCombustible en columnas y Suma de Litros en valores.


### Prompt 4


**Prompt exacto:**

```
"¿Cuál es la tendencia de ventas mes a mes durante 2025? ¿Hay alguna estacionalidad?"
```


**Qué esperar:**

Un análisis temporal con ventas mensuales mostrando si hay meses altos o bajos. Copilot puede identificar tendencias.


**Cómo validar:**

Agrupa las fechas por mes con una tabla dinámica y grafica la serie temporal. Observa si coincide con el análisis de Copilot.


### Prompt 5


**Prompt exacto:**

```
"¿Qué porcentaje de las ventas se pagan con cada método de pago? Desglósalo por sucursal."
```


**Qué esperar:**

Porcentajes de Efectivo, Tarjeta y Transferencia por sucursal en formato tabla o gráfico.


**Cómo validar:**

Usa CONTAR.SI.CONJUNTO para contar transacciones por MetodoPago y Sucursal. Calcula los porcentajes manualmente.


---

## 2. Detección de Errores y Anomalías


### Prompt 6


**Prompt exacto:**

```
"Identifica anomalías en la tabla de nómina. ¿Hay empleados con cambios inusuales de sueldo?"
```


**Qué esperar:**

Copilot debería detectar los 2 empleados con incrementos súbitos de sueldo (empleado 3 en julio, empleado 11 en octubre).


**Cómo validar:**

Filtra por cada empleado y grafica su SueldoBase por periodo. Los saltos serán visibles como picos en la línea.


### Prompt 7


**Prompt exacto:**

```
"¿Hay datos faltantes en la nómina? ¿Qué empleados tienen meses sin registro?"
```


**Qué esperar:**

Debería identificar al empleado con 3 meses faltantes (abril, mayo, junio 2025).


**Cómo validar:**

Usa CONTAR.SI para contar registros por empleado. El que tenga menos de 12 registros mensuales tiene meses faltantes.


### Prompt 8


**Prompt exacto:**

```
"Detecta si hay empleados con horas extra inusualmente altas. ¿En qué periodos ocurre?"
```


**Qué esperar:**

Copilot debería señalar que diciembre tiene picos de horas extra en todos los empleados.


**Cómo validar:**

Calcula el promedio de HorasExtra por periodo. Diciembre debería tener un promedio significativamente mayor.


### Prompt 9


**Prompt exacto:**

```
"Revisa si algún vendedor tiene un rendimiento consistentemente bajo comparado con el promedio."
```


**Qué esperar:**

Identificación de Vendedor_3 como el de menor rendimiento sistemático (transacciones pequeñas).


**Cómo validar:**

Calcula promedio de Litros y Venta_Total por vendedor. Vendedor_3 tendrá promedios notablemente menores.


---

## 3. Cálculos y Fórmulas


### Prompt 10


**Prompt exacto:**

```
"Crea una columna que calcule el ISR marginal para cada empleado basado en su TotalPercepcion mensual."
```


**Qué esperar:**

Copilot agregará una columna con fórmula que aplique la tarifa Art. 96 LISR, ubicando el rango y aplicando el porcentaje correspondiente.


**Cómo validar:**

Compara los valores de la nueva columna con la columna ISR existente. Deben ser iguales o muy cercanos.


### Prompt 11


**Prompt exacto:**

```
"Calcula una comisión del 2% sobre Venta_Total para cada vendedor y agrégala como nueva columna."
```


**Qué esperar:**

Una nueva columna 'Comisión' con la fórmula =Venta_Total*0.02 aplicada a todas las filas.


**Cómo validar:**

Verifica manualmente: multiplica Venta_Total por 0.02 en algunas filas y compara.


### Prompt 12


**Prompt exacto:**

```
"Agrega una columna que clasifique cada venta como 'Alta' (>$5,000), 'Media' ($1,000-$5,000) o 'Baja' (<$1,000)."
```


**Qué esperar:**

Copilot creará una columna con función SI anidada o IFS que clasifique por rango de Venta_Total.


**Cómo validar:**

Filtra por cada categoría y verifica que los montos correspondan a los rangos definidos.


### Prompt 13


**Prompt exacto:**

```
"Calcula el sueldo neto promedio por puesto y ordena de mayor a menor."
```


**Qué esperar:**

Un resumen con el promedio de NetoPagar agrupado por Puesto, ordenado descendentemente.


**Cómo validar:**

Usa PROMEDIO.SI para calcular el promedio de NetoPagar por cada puesto único.


---

## 4. Gráficos y Visualización


### Prompt 14


**Prompt exacto:**

```
"Crea un gráfico de barras que muestre las ventas totales por mes durante 2025."
```


**Qué esperar:**

Un gráfico de barras verticales con 12 barras (Ene-Dic) mostrando la suma de Venta_Total por mes.


**Cómo validar:**

Crea tu propio gráfico con tabla dinámica de Fecha (agrupada por mes) vs Suma de Venta_Total.


### Prompt 15


**Prompt exacto:**

```
"Muestra la distribución de ventas por tipo de combustible con un gráfico de pastel."
```


**Qué esperar:**

Un gráfico circular con 3 segmentos (Magna, Premium, Diesel) mostrando proporción de ventas.


**Cómo validar:**

Suma Venta_Total por TipoCombustible y crea un gráfico circular manual para comparar.


### Prompt 16


**Prompt exacto:**

```
"Genera un gráfico de líneas que muestre la evolución del sueldo base de los 5 empleados con mayor sueldo."
```


**Qué esperar:**

Un gráfico de líneas con 5 series temporales mostrando SueldoBase por periodo.


**Cómo validar:**

Identifica los 5 empleados con mayor SueldoBase y grafícalos manualmente con tabla dinámica.


### Prompt 17


**Prompt exacto:**

```
"Crea un gráfico comparativo de ventas por turno (Matutino, Vespertino, Nocturno) para cada sucursal."
```


**Qué esperar:**

Un gráfico de barras agrupadas con 3 grupos (sucursales) y 3 barras cada uno (turnos).


**Cómo validar:**

Tabla dinámica con Sucursal en filas, Turno en columnas y Suma de Venta_Total en valores.


---

## 5. Automatización y Resumen


### Prompt 18


**Prompt exacto:**

```
"Genera un resumen ejecutivo de la tabla de ventas: totales, promedios, mejor sucursal, mejor vendedor y tendencia."
```


**Qué esperar:**

Un párrafo o tabla con KPIs principales: venta total, promedio por transacción, sucursal líder, vendedor estrella.


**Cómo validar:**

Calcula cada KPI manualmente con funciones SUMA, PROMEDIO, MAX, y verifica que coincidan.


### Prompt 19


**Prompt exacto:**

```
"Crea una tabla de frecuencia que muestre cuántas transacciones hay por rango de litros (0-50, 50-100, 100-200, 200-500)."
```


**Qué esperar:**

Una tabla con 4 filas mostrando el conteo de transacciones en cada rango de litros.


**Cómo validar:**

Usa CONTAR.SI.CONJUNTO con criterios de rango para contar transacciones en cada intervalo.


### Prompt 20


**Prompt exacto:**

```
"Resume las deducciones totales por tipo (ISR, IMSS, Otras) para toda la nómina y calcula el porcentaje que representa cada una."
```


**Qué esperar:**

Una tabla resumen con 3 filas: ISR total, IMSS total, OtrasDeducciones total, y su porcentaje del total de deducciones.


**Cómo validar:**

Suma cada columna de deducciones y calcula el porcentaje de cada una sobre TotalDeduccion.


---

## Consejos para mejores resultados con Copilot

- Sé específico: en lugar de 'analiza los datos', di exactamente qué columnas y qué operación.
- Menciona el nombre de la tabla: 'En la tabla Ventas_Gasolinera, calcula...'
- Pide un paso a la vez: no combines múltiples solicitudes en un solo prompt.
- Si la respuesta no es correcta, reformula el prompt con más detalle.
- Usa Copilot para explorar, pero siempre valida con fórmulas tradicionales.
- Guarda las fórmulas útiles que Copilot genere para reutilizarlas.

## Limitaciones actuales de Copilot en Excel

- No puede acceder a archivos locales; el archivo debe estar en la nube (OneDrive/SharePoint).
- Solo trabaja con datos en formato Tabla (Ctrl+T).
- Puede generar fórmulas incorrectas; siempre verifica los resultados.
- No reemplaza el criterio contable profesional (NIF, LISR, CFF).
- Disponibilidad limitada a ciertos planes de Microsoft 365.
- Las respuestas pueden variar si repites el mismo prompt.