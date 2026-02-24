# Guía de Referencia

**Módulo 3: Visualización de Impacto y Reportes Ejecutivos**

*Visualización de datos, selección de gráficos y mejores prácticas*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Guía de selección de gráficos

Antes de crear cualquier gráfico, responde esta pregunta: **¿Qué quiero comunicar?** La respuesta determina el tipo de gráfico ideal.


### Árbol de decisión

Sigue este flujo para elegir el gráfico correcto:


| Pregunta | Respuesta | Tipo de gráfico |
| --- | --- | --- |
| ¿Quiero comparar cantidades entre categorías? | Sí | Barras (horizontales o verticales) |
| ¿Quiero mostrar tendencia en el tiempo? | Sí | Líneas |
| ¿Quiero mostrar proporción del total? | Sí, 5-6 categorías máx | Pastel o Dona |
| ¿Quiero composición a lo largo del tiempo? | Sí | Columnas apiladas |
| ¿Quiero relación entre dos variables? | Sí | Dispersión (scatter) |
| ¿Quiero mostrar progreso hacia una meta? | Sí | Indicador o barra de progreso |


### Errores comunes en la selección

- Usar pastel con más de 6 categorías (se vuelve ilegible)
- Usar líneas para datos categóricos sin orden temporal
- Usar efecto 3D (distorsiona las proporciones)
- Usar doble eje Y sin justificación clara
- Usar gráficos complejos cuando uno simple basta

---

## 2. Checklist de limpieza visual

Después de crear tu gráfico, revisa cada punto de esta lista. El objetivo es eliminar todo lo que no comunica información útil.


| # | Verificación | Acción |
| --- | --- | --- |
| 1 | Título descriptivo | Cambiar 'Gráfico 1' por título que responda '¿de qué es este gráfico?' |
| 2 | ¿Leyenda necesaria? | Si solo hay 1 serie, eliminar la leyenda |
| 3 | Botones de campo | En gráficos dinámicos: clic derecho > Ocultar botones de campo |
| 4 | Líneas de cuadrícula | Eliminar si el gráfico es simple y las etiquetas ya muestran valores |
| 5 | Etiquetas de datos | Agregar valores sobre barras/puntos si son pocos datos |
| 6 | Formato de ejes | Usar K (miles), M (millones). Quitar decimales innecesarios |
| 7 | Colores | Máximo 3-4 colores. Consistentes con el significado |
| 8 | Tamaño de fuente | Mínimo 10pts para etiquetas, 14pts para títulos |
| 9 | Orden de datos | Ordenar barras de mayor a menor (o cronológico si aplica) |
| 10 | Fuente de datos | Incluir periodo y unidad (ej: 'Ventas 2025, cifras en MDP') |


**Principio de Edward Tufte:** "Maximiza la tinta de datos, minimiza la tinta de decoración." Cada píxel debe tener un propósito.


---

## 3. Paleta de colores de referencia

Usa estos colores de forma consistente en todos tus reportes. Cada color tiene un significado intuitivo que facilita la lectura.


| Color | Código HEX | Uso recomendado | Ejemplo |
| --- | --- | --- | --- |
| Azul | #2563EB | Color principal / institucional | Títulos, barras principales |
| Verde | #10B981 | Positivo / crecimiento / Magna | Utilidades, cumplimiento, Magna |
| Rojo | #EF4444 | Atención / negativo / Premium | Gastos excesivos, alertas, Premium |
| Amarillo | #F59E0B | Precaución / intermedio | Advertencias, datos pendientes |
| Gris | #64748B | Secundario / referencia / Diesel | Datos de periodo anterior, Diesel |
| Gris claro | #CBD5E1 | Bordes / fondos | Líneas de cuadrícula, bordes de tabla |


### Ejemplo aplicado: Gasolinera

- Magna = Verde (#10B981) — la bomba verde que todos conocen
- Premium = Rojo (#EF4444) — la bomba roja de alto octanaje
- Diesel = Gris (#64748B) — la bomba gris/negra para vehículos pesados

### Consideraciones de accesibilidad

- No depender únicamente de rojo vs verde (daltonismo)
- Usar texturas o patrones además de color cuando sea posible
- Asegurar contraste suficiente entre texto y fondo
- Probar el gráfico en escala de grises para verificar legibilidad

---

## 4. Ejercicios prácticos

Completa estos ejercicios usando los archivos del Módulo 3. Cada ejercicio refuerza un concepto diferente de visualización.


### Ejercicio 1: Gráfico de combustible personalizado

**Archivo:** 07_Dashboard_Ventas_Combustible.xlsx


**Instrucciones:**

- Abre la hoja 'Datos' y selecciona los meses de Enero a Junio
- Crea un gráfico de líneas con marcadores para los litros de cada combustible
- Aplica los colores correctos: Magna=verde, Premium=rojo, Diesel=gris
- Agrega etiquetas de datos solo en los puntos máximo y mínimo
- Coloca un título descriptivo y elimina la cuadrícula

**Qué aprendes:** A crear gráficos de líneas con colores significativos y limpieza visual profesional.


### Ejercicio 2: Estado de resultados con gráfico combinado

**Archivo:** 08_Comparativa_Anual_Ventas_Gastos.xlsx


**Instrucciones:**

- En la hoja 'Estado_Resultados', selecciona Total Ingresos y Total Gastos para ambos años
- Crea un gráfico de barras agrupadas que muestre los 4 valores
- Usa azul para ingresos y rojo para gastos
- Agrega una línea que muestre la Utilidad Bruta de cada año (eje secundario)
- Formatea el eje Y en millones (ej: $80M)

**Qué aprendes:** A crear gráficos combinados (barras + línea) y usar ejes secundarios para mostrar diferentes escalas.


### Ejercicio 3: Dashboard básico con segmentadores

**Archivo:** Crear desde cero usando los datos de combustible


**Instrucciones:**

- Copia la hoja 'Datos' del archivo 07 a un libro nuevo
- Convierte los datos en Tabla (Ctrl+T)
- Crea una Tabla Dinámica en una hoja nueva
- Inserta un Gráfico Dinámico de columnas apiladas
- Agrega un segmentador por Trimestre (agrupa los meses)
- Aplica la paleta de colores del curso y limpia el gráfico

**Qué aprendes:** La combinación Tabla Dinámica + Gráfico Dinámico + Segmentador, que es la base del dashboard del Módulo 4.


---

## 5. Fórmulas útiles para gráficos

Estas fórmulas de Excel te ayudan a preparar datos para gráficos más efectivos.


| Fórmula | Descripción | Ejemplo |
| --- | --- | --- |
| SUMAR.SI | Suma condicional para agrupar categorías | =SUMAR.SI(A:A,"Magna",B:B) |
| CONTAR.SI | Cuenta registros por categoría | =CONTAR.SI(A:A,"Enero") |
| PROMEDIO.SI | Promedio condicional | =PROMEDIO.SI(A:A,"Premium",C:C) |
| MAX / MIN | Identifica picos y valles para etiquetas | =MAX(B2:B13) |
| TEXTO | Formatea números para etiquetas personalizadas | =TEXTO(B2,"$#,##0") |
| REDONDEAR | Simplifica cifras para presentación | =REDONDEAR(B2/1000000,1) |


### Atajos de teclado para gráficos en Excel

| Atajo | Acción |
| --- | --- |
| Alt + F1 | Insertar gráfico rápido en la hoja actual |
| F11 | Insertar gráfico en hoja nueva |
| Ctrl + T | Convertir rango en Tabla (base para gráficos dinámicos) |
| Alt + N + C | Abrir menú de insertar gráfico |
| Ctrl + 1 | Abrir formato de elemento seleccionado |
| Supr | Eliminar elemento seleccionado del gráfico |
