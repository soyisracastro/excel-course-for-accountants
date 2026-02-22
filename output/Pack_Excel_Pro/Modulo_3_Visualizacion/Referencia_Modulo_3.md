# Guia de Referencia

**Modulo 3: Visualización de Impacto y Reportes Ejecutivos**

*Visualizacion de datos, seleccion de graficos y mejores practicas*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Guia de seleccion de graficos

Antes de crear cualquier grafico, responde esta pregunta: **Que quiero comunicar?** La respuesta determina el tipo de grafico ideal.


### Arbol de decision

Sigue este flujo para elegir el grafico correcto:


| Pregunta | Respuesta | Tipo de grafico |
| --- | --- | --- |
| Quiero comparar cantidades
entre categorias? | Si | Barras (horizontales
o verticales) |
| Quiero mostrar tendencia
en el tiempo? | Si | Lineas |
| Quiero mostrar proporcion
del total? | Si, 5-6 categorias max | Pastel o Dona |
| Quiero composicion
a lo largo del tiempo? | Si | Columnas apiladas |
| Quiero relacion entre
dos variables? | Si | Dispersion (scatter) |
| Quiero mostrar progreso
hacia una meta? | Si | Indicador o barra
de progreso |


### Errores comunes en la seleccion

- Usar pastel con mas de 6 categorias (se vuelve ilegible)
- Usar lineas para datos categoricos sin orden temporal
- Usar efecto 3D (distorsiona las proporciones)
- Usar doble eje Y sin justificacion clara
- Usar graficos complejos cuando uno simple basta
---

## 2. Checklist de limpieza visual

Despues de crear tu grafico, revisa cada punto de esta lista. El objetivo es eliminar todo lo que no comunica informacion util.


| # | Verificacion | Accion |
| --- | --- | --- |
| 1 | Titulo descriptivo | Cambiar 'Grafico 1' por titulo que responda
'de que es este grafico?' |
| 2 | Leyenda necesaria? | Si solo hay 1 serie, eliminar la leyenda |
| 3 | Botones de campo | En graficos dinamicos: click derecho >
Ocultar botones de campo |
| 4 | Lineas de cuadricula | Eliminar si el grafico es simple
y las etiquetas ya muestran valores |
| 5 | Etiquetas de datos | Agregar valores sobre barras/puntos
si son pocos datos |
| 6 | Formato de ejes | Usar K (miles), M (millones).
Quitar decimales innecesarios |
| 7 | Colores | Maximo 3-4 colores.
Consistentes con el significado |
| 8 | Tamano de fuente | Minimo 10pts para etiquetas,
14pts para titulos |
| 9 | Orden de datos | Ordenar barras de mayor a menor
(o cronologico si aplica) |
| 10 | Fuente de datos | Incluir periodo y unidad
(ej: 'Ventas 2025, cifras en MDP') |


**Principio de Edward Tufte:** "Maximiza la tinta de datos, minimiza la tinta de decoracion." Cada pixel debe tener un proposito.

---

## 3. Paleta de colores de referencia

Usa estos colores de forma consistente en todos tus reportes. Cada color tiene un significado intuitivo que facilita la lectura.


| Color | Codigo HEX | Uso recomendado | Ejemplo |
| --- | --- | --- | --- |
| Azul | #2563EB | Color principal / institucional | Titulos, barras principales |
| Verde | #10B981 | Positivo / crecimiento / Magna | Utilidades, cumplimiento, Magna |
| Rojo | #EF4444 | Atencion / negativo / Premium | Gastos excesivos, alertas, Premium |
| Amarillo | #F59E0B | Precaucion / intermedio | Advertencias, datos pendientes |
| Gris | #64748B | Secundario / referencia / Diesel | Datos de periodo anterior, Diesel |
| Gris claro | #CBD5E1 | Bordes / fondos | Lineas de cuadricula, bordes de tabla |


### Ejemplo aplicado: Gasolinera

- Magna = Verde (#10B981) -- la bomba verde que todos conocen
- Premium = Rojo (#EF4444) -- la bomba roja de alto octanaje
- Diesel = Gris (#64748B) -- la bomba gris/negra para vehiculos pesados

### Consideraciones de accesibilidad

- No depender unicamente de rojo vs verde (daltonismo)
- Usar texturas o patrones ademas de color cuando sea posible
- Asegurar contraste suficiente entre texto y fondo
- Probar el grafico en escala de grises para verificar legibilidad
---

## 4. Ejercicios practicos

Completa estos ejercicios usando los archivos del Modulo 3. Cada ejercicio refuerza un concepto diferente de visualizacion.


### Ejercicio 1: Grafico de combustible personalizado

**Archivo:** 07_Dashboard_Ventas_Combustible.xlsx


**Instrucciones:**

- Abre la hoja 'Datos' y selecciona los meses de Enero a Junio
- Crea un grafico de lineas con marcadores para los litros de cada combustible
- Aplica los colores correctos: Magna=verde, Premium=rojo, Diesel=gris
- Agrega etiquetas de datos solo en los puntos maximo y minimo
- Coloca un titulo descriptivo y elimina la cuadricula

**Que aprendes:** A crear graficos de lineas con colores significativos y limpieza visual profesional.


### Ejercicio 2: Estado de resultados con grafico combinado

**Archivo:** 08_Comparativa_Anual_Ventas_Gastos.xlsx


**Instrucciones:**

- En la hoja 'Estado_Resultados', selecciona Total Ingresos y Total Gastos para ambos anios
- Crea un grafico de barras agrupadas que muestre los 4 valores
- Usa azul para ingresos y rojo para gastos
- Agrega una linea que muestre la Utilidad Bruta de cada anio (eje secundario)
- Formatea el eje Y en millones (ej: $80M)

**Que aprendes:** A crear graficos combinados (barras + linea) y usar ejes secundarios para mostrar diferentes escalas.


### Ejercicio 3: Dashboard basico con segmentadores

**Archivo:** Crear desde cero usando los datos de combustible


**Instrucciones:**

- Copia la hoja 'Datos' del archivo 07 a un libro nuevo
- Convierte los datos en Tabla (Ctrl+T)
- Crea una Tabla Dinamica en una hoja nueva
- Inserta un Grafico Dinamico de columnas apiladas
- Agrega un segmentador por Trimestre (agrupa los meses)
- Aplica la paleta de colores del curso y limpia el grafico

**Que aprendes:** La combinacion Tabla Dinamica + Grafico Dinamico + Segmentador, que es la base del dashboard del Modulo 4.

---

## 5. Formulas utiles para graficos

Estas formulas de Excel te ayudan a preparar datos para graficos mas efectivos.


| Formula | Descripcion | Ejemplo |
| --- | --- | --- |
| SUMAR.SI | Suma condicional para
agrupar categorias | =SUMAR.SI(A:A,"Magna",B:B) |
| CONTAR.SI | Cuenta registros por
categoria | =CONTAR.SI(A:A,"Enero") |
| PROMEDIO.SI | Promedio condicional | =PROMEDIO.SI(A:A,"Premium",C:C) |
| MAX / MIN | Identifica picos y valles
para etiquetas | =MAX(B2:B13) |
| TEXTO | Formatea numeros para
etiquetas personalizadas | =TEXTO(B2,"$#,##0") |
| REDONDEAR | Simplifica cifras para
presentacion | =REDONDEAR(B2/1000000,1) |


### Atajos de teclado para graficos en Excel

| Atajo | Accion |
| --- | --- |
| Alt + F1 | Insertar grafico rapido en la hoja actual |
| F11 | Insertar grafico en hoja nueva |
| Ctrl + T | Convertir rango en Tabla (base para graficos dinamicos) |
| Alt + N + C | Abrir menu de insertar grafico |
| Ctrl + 1 | Abrir formato de elemento seleccionado |
| Supr | Eliminar elemento seleccionado del grafico |
