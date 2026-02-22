# MÓDULO 3: Visualización de Impacto y Reportes Ejecutivos

*Graficos que cuentan historias, no solo muestran numeros*

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## "Una imagen vale mas que mil palabras" en contabilidad

- Un director no lee 500 filas de datos -- lee un grafico de 5 segundos
- El cerebro procesa imagenes 60,000x mas rapido que texto (MIT)
- En juntas directivas, los reportes visuales generan decisiones mas rapidas
- Error comun: mostrar TODOS los datos en vez del mensaje clave
- Tu trabajo: convertir datos en una historia clara y accionable

## Psicologia del grafico: arbol de decision

- Comparar cantidades entre categorias --> Barras (horizontal o vertical)
- Mostrar tendencia en el tiempo --> Lineas
- Mostrar composicion / proporcion --> Pastel o dona (maximo 5-6 categorias)
- Comparar partes de un todo a lo largo del tiempo --> Columnas apiladas
- Mostrar relacion entre dos variables --> Dispersion (scatter)
- Regla de oro: si el grafico necesita explicacion, escogiste mal el tipo

## Graficos dinamicos desde Tablas Dinamicas

- Selecciona tu Tabla Dinamica -> Insertar -> Grafico Dinamico
- El grafico se actualiza automaticamente al filtrar la tabla
- Puedes usar segmentadores (slicers) para filtrar visualmente
- Ventaja: no necesitas reconstruir el grafico cuando cambian los datos
- Combinacion poderosa: Tabla Dinamica + Grafico Dinamico + Segmentador

## Caso practico: Ventas por vendedor (grafico de barras)

- Escenario: 5 vendedores, ventas trimestrales
- Tabla Dinamica: Filas = Vendedor, Valores = Suma de Ventas
- Grafico de barras horizontales --> facil comparar nombres largos
- Ordenar de mayor a menor para impacto visual inmediato
- Agregar etiquetas de datos para mostrar cifras exactas
- Color unico para todas las barras (evitar arcoiris innecesario)

## Caso practico: Productos mas vendidos (grafico de pastel)

- Escenario: participacion porcentual de 5 productos
- Tabla Dinamica: Filas = Producto, Valores = Suma de Ventas
- Grafico de pastel / dona con porcentajes en etiquetas
- Maximo 5-6 categorias (agrupar el resto como 'Otros')
- Evitar efecto 3D -- distorsiona las proporciones visualmente
- Alternativa: barras apiladas al 100% si hay muchas categorias

## Limpieza visual: menos es mas

- Ocultar botones de campo en graficos dinamicos
- Quitar leyendas redundantes (si solo hay una serie)
- Eliminar lineas de cuadricula si no aportan informacion
- Usar formato de eje: quitar decimales innecesarios
- Titulo claro y descriptivo (no 'Grafico 1')
- Alinear el grafico con los datos de la hoja

## Colores con sentido: no al arcoiris

- Cada color debe tener un significado o proposito
- Ejemplo gasolinera: Magna = verde, Premium = rojo, Diesel = gris
- Verde suele significar positivo / crecimiento / aprobado
- Rojo suele significar atencion / negativo / rechazo
- Gris para datos secundarios o de referencia
- Maximo 3-4 colores por grafico -- consistencia entre reportes
- Considerar daltonismo: no depender solo de rojo vs verde

## Caso: Ventas de combustible (columnas apiladas mensuales)

- Archivo: 07_Dashboard_Ventas_Combustible.xlsx
- Datos: 12 meses x 3 tipos de combustible (litros y monto)
- Grafico de columnas apiladas: cada columna = mes, cada color = tipo
- Se ve la composicion (que tipo vende mas) Y la tendencia (meses altos/bajos)
- Colores: Magna (#10B981), Premium (#EF4444), Diesel (#64748B)
- La altura total de la columna = total de ventas del mes

## Lectura de estacionalidad en el grafico

- Enero bajo: 'cuesta de enero' -- la gente gasta menos despues de diciembre
- Febrero-marzo: recuperacion gradual
- Abril-mayo: estabilizacion cercana al promedio
- Junio-agosto: pico de verano (vacaciones, viajes, mas consumo)
- Septiembre-noviembre: descenso gradual, regreso a clases, ahorro
- Diciembre: baja ligera (gastos navidenos en otras cosas)
- Estas tendencias ayudan a PLANEAR compras y flujo de efectivo

## Comparativa anual: 2024 vs 2025

- Archivo: 08_Comparativa_Anual_Ventas_Gastos.xlsx
- Comparar el mismo rubro en dos periodos es basico en contabilidad
- Barras lado a lado: gris para 2024 (pasado), color fuerte para 2025 (actual)
- Se responde: crecimos o nos encogimos?
- La variacion porcentual da contexto: 5 millones mas suena diferente si es 5% o 50%
- Siempre incluir ambos ejes: absoluto y porcentual

## Estado de Resultados comparativo visual

- Ingresos: Ventas + Otros Ingresos = Total Ingresos
- Gastos: Compras + Gastos Generales + Nomina = Total Gastos
- Utilidad Bruta = Total Ingresos - Total Gastos
- Un grafico para ingresos (azul), otro para gastos (rojo)
- La separacion en dos graficos evita confusion visual
- Insight clave: si gastos crecen mas rapido que ingresos, hay problema

## Datos y etiquetas profesionales

- Etiquetas de datos: mostrar el valor exacto sobre la barra
- Formato del eje: usar miles (K) o millones (M) para cifras grandes
- No repetir informacion: si esta en la etiqueta, no necesita estar en el eje
- Titulo del grafico: debe responder 'de que es este grafico?'
- Fuente de datos: incluir siempre el periodo y la unidad
- Tamano legible: minimo 10 pts para etiquetas, 14 pts para titulos

## Resumen del Modulo 3

- Elegir el tipo de grafico correcto es el 80% del trabajo
- Graficos dinamicos + segmentadores = analisis interactivo
- Limpieza visual: quitar todo lo que no comunica
- Colores con significado, no con decoracion
- Columnas apiladas para composicion + tendencia temporal
- Comparativas anuales: siempre mostrar variacion porcentual
- Cada grafico debe contar UNA historia clara

## Recursos y Siguiente Paso

- Archivo: 07_Dashboard_Ventas_Combustible.xlsx (columnas apiladas)
- Archivo: 08_Comparativa_Anual_Ventas_Gastos.xlsx (barras comparativas)
- PDF: Referencia_Modulo_3.pdf (guia de seleccion de graficos y ejercicios)
- Referencia: Edward Tufte - The Visual Display of Quantitative Information
- Siguiente: Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

*Excel para Contadores y Administrativos — Israel Castro*
