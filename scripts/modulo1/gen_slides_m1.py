"""
Generador: Modulo_1_Logica_Contable.pptx + Script_Modulo_1.md
Modulo 1 — Logica Contable y Funciones de Control

Slides:  ~20 diapositivas cubriendo funciones basicas, ISR 2026, BUSCARV,
         funciones de fecha, RFC, y casos practicos.
Script:  Teleprompter (~35-40 min lectura, ~4500 palabras).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import SLIDES_DIR, TELEPROMPTER_DIR
from scripts.generators.pptx_gen import SlidesGenerator


def build():
    gen = SlidesGenerator(
        filename="Modulo_1_Logica_Contable.md",
        output_dir=SLIDES_DIR,
        script_filename="Script_Modulo_1.md",
        script_dir=TELEPROMPTER_DIR,
    )

    # ── Slide 1: Portada ───────────────────────────────────────────
    gen.add_title_slide(1, "Logica Contable y Funciones de Control")
    gen.script_lines.append(
        "Bienvenidos al Modulo 1 del curso Excel para Contadores y Administrativos. "
        "Soy Israel Castro, contador publico y desarrollador de software, y voy a ser tu guia "
        "durante todo este programa. En este primer modulo vamos a sentar las bases de todo "
        "lo que viene despues. Basicamente, lo que quiero lograr hoy es que dejes de ver a "
        "Excel como una calculadora gigante y empieces a verlo como un asistente inteligente "
        "que trabaja para ti. Vamos a cubrir funciones esenciales, logica contable, calculo de "
        "ISR con la tarifa 2026 oficial, y varias herramientas que vas a usar todos los dias "
        "en tu practica profesional. Asi que preparate, abre tu Excel, y vamonos.\n"
    )

    # ── Slide 2: Objetivo del modulo ───────────────────────────────
    gen.add_content_slide(
        title="Objetivo del Modulo",
        bullets=[
            "Excel deja de ser calculadora, se convierte en asistente",
            "Dominar funciones de control: SI, BUSCARV, TRUNCAR",
            "Calcular ISR 2026 de forma automatica y dinamica",
            "Extraer informacion del RFC con funciones de texto",
            "Crear un calendario de vencimientos con semaforo",
        ],
        script_text=(
            "El objetivo central de este modulo es muy claro: que Excel deje de ser una calculadora "
            "donde escribes numeros y le das enter, y se convierta en un asistente que piensa por ti. "
            "Yo se que muchos de ustedes ya saben sumar en Excel, ya saben hacer una resta, una "
            "multiplicacion. Pero eso es como tener un Ferrari y usarlo solo para ir al super. "
            "Hoy vamos a aprender a manejar ese Ferrari. Vamos a dominar funciones de control como "
            "SI, BUSCARV y TRUNCAR. Vamos a calcular el ISR 2026 con la tarifa oficial publicada en "
            "el Diario Oficial de la Federacion. Vamos a extraer datos del RFC usando funciones de "
            "texto. Y vamos a construir un calendario de vencimientos con un semaforo automatico. "
            "Todo esto en un solo modulo. De hecho, lo que vamos a ver hoy es la base para todo lo "
            "que viene en los modulos 2 al 5. Asi que pongan atencion porque esto es fundamental."
        ),
    )

    # ── Slide 3: Que es una funcion ────────────────────────────────
    gen.add_content_slide(
        title="Que es una Funcion? La Analogia de la Maquina de Cafe",
        bullets=[
            "INPUT: los datos que le das (argumentos)",
            "PROCESO: lo que hace internamente (la magia)",
            "OUTPUT: el resultado que te entrega",
            "Analogia: maquina de cafe",
            "  - INPUT: agua + cafe molido + cantidad de tazas",
            "  - PROCESO: calienta, filtra, mezcla",
            "  - OUTPUT: tu cafe listo para tomar",
        ],
        script_text=(
            "Antes de meternos a formulas, necesito que entendamos algo basico: que es una funcion. "
            "Y la mejor forma de explicarlo es con una analogia. Piensa en una maquina de cafe. "
            "Tu le pones agua, le pones cafe molido, le dices cuantas tazas quieres. Eso es el input, "
            "los datos de entrada. La maquina calienta el agua, la filtra por el cafe, la mezcla. "
            "Ese es el proceso. Y al final te sale tu cafe calientito, listo para tomar. Ese es el "
            "output, el resultado. Las funciones de Excel trabajan exactamente igual. Tu les das datos "
            "de entrada, que se llaman argumentos, la funcion hace algo con esos datos internamente, "
            "y te devuelve un resultado. Lo bonito es que tu no necesitas saber como funciona por "
            "dentro la maquina. Solo necesitas saber que meterle y que esperar. Eso es el poder de las "
            "funciones. Basicamente, cada funcion es una maquinita especializada: una suma, una busca "
            "datos, otra compara valores, otra extrae texto. Tu trabajo es saber cual maquina usar "
            "para cada tarea. Y eso es exactamente lo que vamos a aprender hoy."
        ),
    )

    # ── Slide 4: Anatomia de una funcion ───────────────────────────
    gen.add_content_slide(
        title="Anatomia de una Funcion en Excel",
        bullets=[
            "Estructura: =NOMBRE(argumento1, argumento2, ...)",
            "Siempre empieza con el signo igual (=)",
            "El NOMBRE indica que hace la funcion",
            "Los ARGUMENTOS van entre parentesis, separados por coma",
            "Ejemplo: =SUMA(A1, A2, A3)",
            "Ejemplo: =SI(A1>100, \"Alto\", \"Bajo\")",
            "Consejo: presiona Tab para autocompletar funciones",
        ],
        script_text=(
            "Ahora veamos la anatomia. Toda funcion en Excel tiene la misma estructura. Primero el "
            "signo de igual, que es como decirle a Excel: oye, lo que sigue es una instruccion, no "
            "texto. Despues viene el nombre de la funcion, por ejemplo SUMA, SI, BUSCARV. Y luego "
            "entre parentesis van los argumentos, que son los datos que le pasas a la funcion, "
            "separados por coma. Por ejemplo, SUMA abre parentesis A1 coma A2 coma A3 cierra "
            "parentesis. Le estas diciendo: sumame estas tres celdas. Otro ejemplo: SI abre parentesis "
            "A1 mayor que 100 coma la palabra Alto entre comillas coma la palabra Bajo entre comillas "
            "cierra parentesis. Le estas diciendo: si A1 es mayor que 100, ponme Alto, si no, ponme "
            "Bajo. Un tip muy util: cuando empiezas a escribir el nombre de una funcion, Excel te "
            "muestra sugerencias. Si presionas Tab, se autocompleta. Esto te ahorra mucho tiempo y "
            "te evita errores de escritura. De hecho, les recomiendo que siempre usen Tab en lugar de "
            "escribir el nombre completo. Listo, esa es la anatomia basica. Simple, verdad? Ahora "
            "vamos a poner las manos en la masa."
        ),
    )

    # ── Slide 5: SUMA y PROMEDIO ───────────────────────────────────
    gen.add_content_slide(
        title="SUMA y PROMEDIO: Tus Primeras Funciones Contables",
        bullets=[
            "=SUMA(rango) — Suma todos los valores del rango",
            "=SUMA(B2:B50) — Suma ventas de enero",
            "=PROMEDIO(rango) — Calcula la media aritmetica",
            "=PROMEDIO(C2:C13) — Gasto promedio mensual",
            "Rangos: B2:B50 (columna), B2:F2 (fila), B2:F50 (bloque)",
            "Atajo: Alt + = inserta SUMA automaticamente",
            "Tip: Nombra tus rangos (Ctrl+F3) para mayor claridad",
        ],
        script_text=(
            "Empecemos con las dos funciones mas basicas pero mas usadas en contabilidad: SUMA y "
            "PROMEDIO. SUMA hace exactamente lo que imaginas: le pasas un rango de celdas y te da el "
            "total. Por ejemplo, si tienes ventas diarias en las celdas B2 a B50, escribes "
            "=SUMA(B2:B50) y obtienes el total del mes. Facil. PROMEDIO es igual de simple: te da la "
            "media aritmetica. Si tienes gastos mensuales de enero a diciembre en C2 a C13, con "
            "=PROMEDIO(C2:C13) obtienes el gasto promedio mensual. Ahora, algo importante: los rangos. "
            "B2:B50 es un rango vertical, una columna. B2:F2 es un rango horizontal, una fila. Y "
            "B2:F50 es un bloque completo, un rectangulo de datos. Necesitas entender esto porque "
            "todas las funciones trabajan con rangos. Un atajo que les va a encantar: si seleccionas "
            "una celda debajo de una columna con numeros y presionas Alt mas el signo de igual, Excel "
            "automaticamente te inserta la funcion SUMA con el rango correcto. Pruebalo. Otro tip "
            "avanzado: pueden nombrar sus rangos con Ctrl+F3. En lugar de escribir B2:B50, le pones "
            "un nombre como VentasEnero, y despues escribes =SUMA(VentasEnero). Mucho mas claro y "
            "profesional."
        ),
    )

    # ── Slide 6: Ejercicio Caja Negra ──────────────────────────────
    gen.add_content_slide(
        title="Ejercicio: La Caja Negra — Debuggeando Formulas",
        bullets=[
            "Formula 1:  =SUMA(A1:A10)  pero A5 tiene texto \"N/A\"",
            "  -> Resultado: #VALOR!  |  Solucion: usar SUMA.SI o limpiar datos",
            "Formula 2:  =PROMEDIO(B1:B5)  pero B3 esta vacia",
            "  -> Resultado: promedia solo 4 valores  |  Cuidado con celdas vacias",
            "Formula 3:  =A1+B1*C1  sin parentesis",
            "  -> Resultado: multiplicacion va primero  |  Solucion: =(A1+B1)*C1",
            "Leccion: F2 para entrar a la celda y ver que esta pasando",
        ],
        script_text=(
            "Ahora quiero que hagamos un ejercicio que yo le llamo La Caja Negra. La idea es simple: "
            "les voy a mostrar tres formulas que dan un resultado inesperado, y ustedes tienen que "
            "descubrir por que. Esto es debuggear, y es una habilidad critica. Formula uno: tienes "
            "=SUMA(A1:A10) pero en la celda A5 alguien escribio el texto N/A en lugar de un numero. "
            "Que pasa? Te da error #VALOR! porque SUMA no puede sumar texto. La solucion es usar "
            "SUMA.SI para ignorar errores, o mejor aun, limpiar tus datos desde el principio. "
            "Formula dos: tienes =PROMEDIO(B1:B5) pero B3 esta vacia. El resultado parece correcto "
            "pero esta mal, porque PROMEDIO ignora las celdas vacias. Esta promediando entre 4 "
            "valores en lugar de 5. Si B3 deberia ser cero, tienes que escribir cero explicitamente. "
            "Formula tres: tienes =A1+B1*C1 y esperabas que primero sumara A1 mas B1 y luego "
            "multiplicara por C1. Pero Excel respeta la jerarquia de operaciones: multiplicacion va "
            "primero. La solucion es usar parentesis: =(A1+B1)*C1. Moraleja: cuando algo no cuadra, "
            "presiona F2 para entrar a la celda, ve la formula, revisa los rangos, revisa los datos. "
            "F2 es tu mejor amigo para debuggear."
        ),
    )

    # ── Slide 7: TRUNCAR vs REDONDEAR ──────────────────────────────
    gen.add_content_slide(
        title="Precision Fiscal: TRUNCAR vs REDONDEAR",
        bullets=[
            "=REDONDEAR(3.14159, 2) -> 3.14  (redondea al mas cercano)",
            "=TRUNCAR(3.14159, 2)   -> 3.14  (corta sin redondear)",
            "=REDONDEAR(3.1489, 2)  -> 3.15  (redondeo sube)",
            "=TRUNCAR(3.1489, 2)    -> 3.14  (truncar siempre corta)",
            "Art. 17-A CFF: Factor de Actualizacion se trunca al diezmilésimo (4 decimales)",
            "=TRUNCAR(INPC_reciente / INPC_anterior, 4)",
            "En fiscalidad: TRUNCAR es la regla, REDONDEAR es la excepcion",
        ],
        script_text=(
            "Esta diapositiva es crucial para cualquier contador. Hay una diferencia enorme entre "
            "TRUNCAR y REDONDEAR, y en el ambito fiscal te puede costar dinero si las confundes. "
            "REDONDEAR hace lo que aprendiste en la escuela: si el siguiente digito es 5 o mayor, "
            "sube; si es menor que 5, baja. TRUNCAR simplemente corta. No le importa el siguiente "
            "digito, simplemente elimina todo lo que esta despues del decimal que le indiques. "
            "Veamos un ejemplo: 3.1489 redondeado a 2 decimales es 3.15, porque el tercer decimal "
            "es 8, que es mayor a 5, entonces sube. Pero 3.1489 truncado a 2 decimales es 3.14, "
            "simplemente corta en el segundo decimal y listo. Ahora, por que importa esto? Porque "
            "el articulo 17-A del Codigo Fiscal de la Federacion establece que el Factor de "
            "Actualizacion se debe truncar al diezmilésimo, es decir, a 4 decimales. No redondear, "
            "truncar. Si tu redondeas en lugar de truncar, tu factor puede ser ligeramente mayor, y "
            "eso afecta el monto actualizado. En una auditoria, eso se detecta. Asi que la regla de "
            "oro en fiscalidad mexicana es: TRUNCAR es la norma, REDONDEAR es la excepcion. Grabenlo."
        ),
    )

    # ── Slide 8: Factor de Actualizacion ───────────────────────────
    gen.add_content_slide(
        title="Factor de Actualizacion con INPC 2026",
        bullets=[
            "Formula: Factor = INPC reciente / INPC anterior",
            "INPC Dic 2025 (reciente): 141.200",
            "INPC Dic 2024 (anterior): 136.163",
            "Factor sin truncar: 141.200 / 136.163 = 1.036996...",
            "Factor truncado (CFF): =TRUNCAR(141.200/136.163, 4) = 1.0369",
            "Aplicacion: Monto original x Factor = Monto actualizado",
            "Ejemplo: $100,000 x 1.0369 = $103,690.00",
        ],
        script_text=(
            "Pongamos en practica lo que acabamos de aprender. El Factor de Actualizacion se usa "
            "para ajustar montos por inflacion, y se calcula dividiendo el INPC del periodo reciente "
            "entre el INPC del periodo anterior. Usando los datos del INEGI, el INPC de diciembre "
            "2025 es 141.200 y el de diciembre 2024 es 136.163. Si dividimos, nos da 1.036996 y "
            "muchos decimales mas. Pero como ya vimos, el CFF nos dice que debemos truncar a 4 "
            "decimales. Entonces en Excel escribimos =TRUNCAR(141.200/136.163, 4) y nos da 1.0369. "
            "Ese es nuestro factor oficial. Ahora, para que sirve? Para actualizar cualquier monto "
            "por inflacion. Si tienes una contribucion de 100 mil pesos, la multiplicas por 1.0369 "
            "y obtienes 103,690 pesos. Eso es lo que vale esa contribucion en pesos actuales. En el "
            "archivo de Excel que les prepare, la hoja Actualizacion_CFF ya tiene todo esto armado. "
            "Solo necesitan cambiar los valores del INPC y el monto original, y automaticamente les "
            "calcula el factor truncado y el monto actualizado. Asi de poderoso es Excel cuando lo "
            "usas con funciones."
        ),
    )

    # ── Slide 9: Coeficiente de Utilidad ───────────────────────────
    gen.add_content_slide(
        title="Coeficiente de Utilidad — Art. 14 CFF",
        bullets=[
            "CU = Utilidad Fiscal / Ingresos Nominales del ejercicio anterior",
            "Se usa para calcular pagos provisionales de ISR",
            "Ejemplo: Utilidad Fiscal $450,000 / Ingresos $3,200,000",
            "CU = 0.140625 -> TRUNCAR a 4 dec = 0.1406",
            "Pago provisional = (Ingresos del periodo x CU) - Pagos anteriores",
            "Excel formula: =TRUNCAR(B2/B3, 4)",
            "BUSCARV puede traer los datos desde tu balance de comprobacion",
        ],
        script_text=(
            "Otro concepto fiscal donde Excel brilla es el Coeficiente de Utilidad del articulo 14 "
            "del CFF. Este coeficiente se usa para calcular los pagos provisionales de ISR de las "
            "personas morales. La formula es: utilidad fiscal del ejercicio anterior dividida entre "
            "los ingresos nominales del mismo ejercicio. Por ejemplo, si tu empresa tuvo una utilidad "
            "fiscal de 450 mil pesos y los ingresos fueron 3 millones 200 mil, el coeficiente es "
            "0.140625. Pero atencion, otra vez aplica el TRUNCAR a 4 decimales. Entonces tu CU "
            "oficial es 0.1406. Con ese coeficiente, cada mes multiplicas tus ingresos acumulados "
            "del periodo por 0.1406, le restas los pagos provisionales que ya hiciste, y eso te da "
            "tu pago provisional del mes. En Excel todo esto se puede automatizar. De hecho, les "
            "recomiendo que en su hoja de pagos provisionales tengan el CU calculado con la formula "
            "=TRUNCAR(B2/B3, 4), donde B2 es la utilidad fiscal y B3 son los ingresos. Y mas "
            "adelante, cuando veamos BUSCARV, van a poder jalar esos datos directamente desde su "
            "balance de comprobacion o su estado de resultados. Todo conectado, todo automatico."
        ),
    )

    # ── Slide 10: Funcion SI ───────────────────────────────────────
    gen.add_content_slide(
        title="Funcion SI: La Toma de Decisiones en Excel",
        bullets=[
            "Sintaxis: =SI(prueba_logica, valor_si_verdadero, valor_si_falso)",
            "Ejemplo basico: =SI(B2>10000, \"Requiere autorizacion\", \"Aprobado\")",
            "Con numeros: =SI(C5>=0, C5, 0)  (no permitir negativos)",
            "Comparadores: =, >, <, >=, <=, <>",
            "SI anidado: =SI(A1>100, \"A\", SI(A1>50, \"B\", \"C\"))",
            "Maximo recomendado: 3 niveles de anidacion",
            "Alternativa moderna: SI.CONJUNTO (Excel 365)",
        ],
        script_text=(
            "Llegamos a una de las funciones mas poderosas de Excel: la funcion SI. Esta funcion le "
            "da a Excel la capacidad de tomar decisiones. Tiene tres argumentos: primero la prueba "
            "logica, que es una pregunta de si o no. Segundo, que hacer si la respuesta es verdadero. "
            "Y tercero, que hacer si la respuesta es falso. Por ejemplo: =SI(B2 es mayor que 10000, "
            "pon Requiere autorizacion, si no pon Aprobado). Asi de simple. Puedes usar cualquier "
            "comparador: igual, mayor que, menor que, mayor o igual, menor o igual, y diferente que, "
            "que se escribe con los signos menor y mayor juntos. Ahora, lo interesante es que puedes "
            "anidar funciones SI, es decir, meter un SI dentro de otro SI. Por ejemplo: SI A1 es "
            "mayor que 100, pon A; si no, revisa: SI A1 es mayor que 50, pon B; si no, pon C. Eso "
            "te da tres categorias con dos niveles de SI. Mi recomendacion es no pasar de 3 niveles "
            "de anidacion, porque se vuelve muy dificil de leer y de debuggear. Si necesitas mas "
            "condiciones, en Excel 365 existe SI.CONJUNTO, que es mucho mas limpio. Pero por ahora, "
            "con SI anidado hasta 3 niveles van a cubrir el 90 por ciento de sus necesidades."
        ),
    )

    # ── Slide 11: Caso practico - Validacion bancaria ──────────────
    gen.add_content_slide(
        title="Caso Practico: Validacion Bancaria con SI",
        bullets=[
            "Escenario: Archivo de pagos a proveedores",
            "Columna A: Nombre del proveedor",
            "Columna B: Banco (Santander, BBVA, Banorte, Otro)",
            "Columna C: Monto del pago",
            "Columna D: Comision por transferencia SPEI",
            "Formula: =SI(B2=\"Santander\", 0, SI(B2=\"BBVA\", 4.50, 7.80))",
            "Resultado: Santander $0 (mismo banco), BBVA $4.50, otros $7.80",
            "Columna E: =C2+D2  (Total con comision)",
        ],
        script_text=(
            "Veamos un caso practico que muchos de ustedes van a reconocer. Imagina que tienes un "
            "archivo de pagos a proveedores. En la columna A tienes el nombre del proveedor, en la B "
            "el banco del proveedor, en la C el monto del pago. Y necesitas calcular la comision por "
            "transferencia SPEI en la columna D. Si el proveedor tiene cuenta en Santander y tu "
            "empresa tambien esta en Santander, la transferencia entre cuentas del mismo banco no "
            "tiene costo. Si es a BBVA, la comision es 4.50 pesos. Y si es a cualquier otro banco, "
            "son 7.80 pesos. La formula seria: =SI(B2 igual a Santander entre comillas, 0, SI(B2 "
            "igual a BBVA entre comillas, 4.50, 7.80)). Fijense como usamos un SI anidado: primero "
            "pregunta si es Santander, si si, cero. Si no, pregunta si es BBVA, si si, 4.50. Si "
            "tampoco, entonces 7.80. Y en la columna E simplemente sumas el monto mas la comision: "
            "=C2+D2. Esto parece simple, pero imagina que tienes 500 pagos al mes. Calcular esto "
            "manualmente te tomaria horas. Con la formula, lo haces una vez, la arrastras, y listo. "
            "Eso es trabajar inteligentemente."
        ),
    )

    # ── Slide 12: BUSCARV ──────────────────────────────────────────
    gen.add_content_slide(
        title="BUSCARV: Los 4 Argumentos en Contexto ISR",
        bullets=[
            "=BUSCARV(valor_buscado, tabla, columna, [ordenado])",
            "Argumento 1 - valor_buscado: el ingreso gravable",
            "Argumento 2 - tabla: la tarifa ISR 2026 (rango o nombre)",
            "Argumento 3 - columna: que dato quieres (1=lim_inf, 3=cuota, 4=%)",
            "Argumento 4 - ordenado: VERDADERO (busqueda aproximada para rangos)",
            "VERDADERO: busca el valor mas cercano MENOR O IGUAL (para tarifas)",
            "FALSO: busca coincidencia exacta (para catalogos)",
            "Ejemplo ISR: =BUSCARV(450000, Tarifa_Anual_2026, 3, VERDADERO)",
        ],
        script_text=(
            "BUSCARV. Esta es probablemente la funcion mas importante para un contador en Excel. "
            "BUSCARV busca un valor en la primera columna de una tabla y te devuelve un valor de otra "
            "columna de la misma fila. Tiene 4 argumentos. El primero es el valor que quieres buscar, "
            "por ejemplo tu ingreso gravable de 450 mil pesos. El segundo es la tabla donde va a "
            "buscar, que en nuestro caso es la tarifa ISR 2026. El tercero es el numero de columna "
            "del dato que quieres: 1 para el limite inferior, 3 para la cuota fija, 4 para el "
            "porcentaje. Y el cuarto argumento es crucial: VERDADERO o FALSO. VERDADERO significa "
            "busqueda aproximada, busca el valor mas cercano que sea menor o igual a lo que buscas. "
            "Esto es perfecto para tarifas de rangos como el ISR, porque tu ingreso cae dentro de un "
            "rango. FALSO significa busqueda exacta, que se usa para catalogos donde necesitas una "
            "coincidencia precisa. Para ISR siempre usamos VERDADERO. Entonces, =BUSCARV(450000, "
            "Tarifa_Anual_2026, 3, VERDADERO) te devuelve la cuota fija del rango donde cae un "
            "ingreso de 450 mil pesos. Asi de facil. Ahora veamos como armar el calculo completo."
        ),
    )

    # ── Slide 13: Caso practico ISR 2026 ───────────────────────────
    gen.add_content_slide(
        title="Caso Practico: Calculo ISR 2026 Paso a Paso",
        bullets=[
            "Paso 1: Base gravable = Ingresos - Deducciones = $450,000",
            "Paso 2: Limite Inferior  =BUSCARV(450000, Tarifa, 1, VERDADERO) = $424,353.98",
            "Paso 3: Excedente = $450,000 - $424,353.98 = $25,646.02",
            "Paso 4: % Marginal =BUSCARV(450000, Tarifa, 4, VERDADERO) = 23.52%",
            "Paso 5: ISR Marginal = $25,646.02 x 23.52% = $6,031.94",
            "Paso 6: Cuota Fija =BUSCARV(450000, Tarifa, 3, VERDADERO) = $67,981.92",
            "Paso 7: ISR Total = $6,031.94 + $67,981.92 = $74,013.86",
            "Todo con 3 BUSCARV y operaciones basicas",
        ],
        script_text=(
            "Vamos a calcular el ISR 2026 paso a paso. Supongamos que un contribuyente tiene ingresos "
            "anuales de 600 mil pesos y deducciones autorizadas de 150 mil. Su base gravable es 450 "
            "mil pesos. Paso uno, ya lo tenemos. Paso dos: necesitamos el limite inferior del rango "
            "donde cae 450 mil pesos. Con BUSCARV buscamos 450,000 en la tarifa, columna 1, con "
            "VERDADERO, y nos da 424,353.98. Paso tres: el excedente es 450,000 menos 424,353.98 "
            "igual a 25,646.02 pesos. Paso cuatro: buscamos el porcentaje marginal, BUSCARV 450,000 "
            "en la tarifa, columna 4, VERDADERO, y nos da 23.52 por ciento. Paso cinco: el ISR "
            "marginal es 25,646.02 multiplicado por 23.52 por ciento, igual a 6,031.94 pesos. Paso "
            "seis: la cuota fija, BUSCARV 450,000 en la tarifa, columna 3, VERDADERO, nos da "
            "67,981.92. Y paso siete: el ISR total es el ISR marginal mas la cuota fija: 6,031.94 "
            "mas 67,981.92 igual a 74,013.86 pesos. Siete pasos, 3 formulas de BUSCARV y operaciones "
            "basicas. Lo mas importante: si cambias el ingreso en la celda de base gravable, TODO se "
            "recalcula automaticamente. Eso es el poder de las funciones."
        ),
    )

    # ── Slide 14: Dinamismo ────────────────────────────────────────
    gen.add_content_slide(
        title="Dinamismo: Cambia el Ingreso, Cambia Todo",
        bullets=[
            "Celda de entrada (amarilla): Base Gravable = $450,000",
            "Cambia a $280,000 -> ISR se recalcula automaticamente",
            "Cambia a $1,500,000 -> ISR cambia de rango y porcentaje",
            "No necesitas rehacer nada: Excel recalcula en tiempo real",
            "Aplicacion: Planeacion fiscal con escenarios",
            "\"Que pasa si mi cliente gana 200 mil mas?\"",
            "\"Cuanto se ahorra con 50 mil mas de deducciones?\"",
            "Esto es lo que separa a un contador que usa Excel de uno que sabe Excel",
        ],
        script_text=(
            "Aqui es donde Excel se pone verdaderamente interesante. Fijense: toda la calculadora de "
            "ISR depende de una sola celda, la base gravable. Si cambias esa celda de 450 mil a 280 "
            "mil, todos los calculos se actualizan al instante: nuevo limite inferior, nuevo "
            "excedente, nuevo porcentaje, nuevo ISR. Si la cambias a un millon 500 mil, cambia a otro "
            "rango completamente diferente con otro porcentaje y otra cuota fija, y el ISR se "
            "recalcula automaticamente. No necesitas borrar nada, no necesitas rehacer formulas. Todo "
            "esta conectado. Y aqui es donde esto se vuelve una herramienta de planeacion fiscal. "
            "Imaginate que un cliente te pregunta: oye, que pasa si gano 200 mil pesos mas este "
            "anio? Tu cambias la celda, y en un segundo le dices exactamente cuanto mas va a pagar "
            "de ISR. O te pregunta: cuanto me ahorro si meto 50 mil pesos mas de deducciones? "
            "Cambias la celda y listo. De hecho, esto es lo que separa a un contador que usa Excel de "
            "uno que sabe Excel. El que sabe Excel construye herramientas dinamicas que le permiten "
            "responder preguntas al instante. El que solo usa Excel tiene que recalcular todo "
            "manualmente cada vez. Y eso, basicamente, es tiempo y dinero."
        ),
    )

    # ── Slide 15: Funciones de fecha ───────────────────────────────
    gen.add_content_slide(
        title="Funciones de Fecha: HOY, DIA, MES, ANIO",
        bullets=[
            "=HOY() — Fecha actual (se actualiza sola cada dia)",
            "=DIA(fecha) — Extrae el dia (1-31)",
            "=MES(fecha) — Extrae el mes (1-12)",
            "=ANIO(fecha) — Extrae el anio (2026)",
            "=FECHA(anio, mes, dia) — Construye una fecha a partir de partes",
            "=DIAS(fecha_final, fecha_inicial) — Dias entre dos fechas",
            "Ejemplo: =DIAS(\"17/04/2026\", HOY()) -> dias para la declaracion anual",
            "Las fechas en Excel son numeros seriales (1 = 1 enero 1900)",
        ],
        script_text=(
            "Cambiemos de tema y hablemos de las funciones de fecha, que son super utiles para "
            "contadores. La funcion HOY, sin argumentos, te da la fecha de hoy. Lo interesante es que "
            "cada dia que abras el archivo, se actualiza automaticamente. DIA, MES y ANIO extraen las "
            "partes de una fecha. Si tienes la fecha 15 de marzo de 2026 en una celda, DIA te da 15, "
            "MES te da 3, ANIO te da 2026. La funcion FECHA hace lo contrario: le das anio, mes y "
            "dia por separado y te construye una fecha. Y DIAS te calcula la diferencia en dias entre "
            "dos fechas. Por ejemplo, si quieres saber cuantos dias faltan para la declaracion anual "
            "del 17 de abril, escribes =DIAS(17 de abril de 2026, HOY()) y te da los dias restantes. "
            "Un dato que pocos saben: las fechas en Excel son numeros seriales. El 1 de enero de 1900 "
            "es el numero 1, el 2 de enero es el 2, y asi sucesivamente. Hoy es un numero de cinco "
            "digitos. Esto es importante porque puedes sumar y restar fechas. Si a HOY() le sumas 30, "
            "obtienes la fecha de dentro de 30 dias. Esto lo vamos a usar para crear nuestro "
            "calendario de vencimientos."
        ),
    )

    # ── Slide 16: EXTRAE y RFC ─────────────────────────────────────
    gen.add_content_slide(
        title="EXTRAE: Desarmando el RFC Caracter por Caracter",
        bullets=[
            "RFC persona fisica: CAST850315HN7 (13 caracteres)",
            "Posiciones 1-4: Apellido(s) y nombre -> CAST",
            "Posiciones 5-6: Anio de nacimiento -> 85 (1985)",
            "Posiciones 7-8: Mes de nacimiento -> 03 (marzo)",
            "Posiciones 9-10: Dia de nacimiento -> 15",
            "=EXTRAE(RFC, 5, 2) -> \"85\" (anio)",
            "=EXTRAE(RFC, 7, 2) -> \"03\" (mes)",
            "=EXTRAE(RFC, 9, 2) -> \"15\" (dia)",
            "Combinando: =FECHA(1900+EXTRAE(A1,5,2), EXTRAE(A1,7,2), EXTRAE(A1,9,2))",
        ],
        script_text=(
            "Ahora una funcion de texto que es oro para los contadores: EXTRAE. Esta funcion saca una "
            "porcion de texto de una celda. Le dices en que posicion empezar y cuantos caracteres "
            "tomar. Y con esto podemos desarmar el RFC. El RFC de una persona fisica tiene 13 "
            "caracteres. Los primeros 4 son las letras del nombre y apellidos. Las posiciones 5 y 6 "
            "son el anio de nacimiento. Las posiciones 7 y 8 son el mes. Y las posiciones 9 y 10 son "
            "el dia. Entonces, si tenemos el RFC CAST850315HN7 en la celda A1, con =EXTRAE(A1, 5, 2) "
            "obtenemos 85, que es el anio. Con =EXTRAE(A1, 7, 2) obtenemos 03, que es el mes. Y con "
            "=EXTRAE(A1, 9, 2) obtenemos 15, que es el dia. La pregunta es: como sabe Excel si 85 es "
            "1985 o 1885? Pues no lo sabe, nosotros tenemos que definir la logica. Pero para efectos "
            "practicos, si el anio es mayor a 30, asumimos que es 1900 y algo; si es menor o igual a "
            "30, asumimos que es 2000 y algo. Todo esto lo puedes construir con SI y EXTRAE "
            "combinados. De hecho, en el siguiente ejercicio vamos a extraer la fecha de cumpleanos "
            "completa desde el RFC. Es un ejercicio que a los alumnos les encanta."
        ),
    )

    # ── Slide 17: Calendario semaforo ──────────────────────────────
    gen.add_content_slide(
        title="Calendario de Vencimientos con Semaforo (SI Anidado)",
        bullets=[
            "Columna A: Obligacion fiscal (e.firma, DIOT, declaracion, etc.)",
            "Columna B: Fecha de vencimiento",
            "Columna C: Dias restantes  =B2-HOY()",
            "Columna D: Semaforo con SI anidado:",
            "  =SI(C2<0, \"VENCIDO\", SI(C2<=15, \"URGENTE\", SI(C2<=30, \"PROXIMO\", \"OK\")))",
            "Rojo: VENCIDO (dias < 0)  |  Amarillo: URGENTE (0-15 dias)",
            "Naranja: PROXIMO (16-30 dias)  |  Verde: OK (> 30 dias)",
            "Formato condicional para colorear automaticamente",
        ],
        script_text=(
            "Aqui viene algo que van a usar todos los dias. Vamos a crear un calendario de "
            "vencimientos con un sistema de semaforo automatico. Imaginen una tabla donde la columna A "
            "tiene las obligaciones fiscales: renovacion de e.firma, DIOT, declaracion mensual, "
            "declaracion anual, etcetera. La columna B tiene la fecha de vencimiento de cada una. La "
            "columna C calcula automaticamente cuantos dias faltan con la formula =B2-HOY(). Y la "
            "columna D es el semaforo, que usa un SI anidado de 3 niveles. La logica es: si los dias "
            "restantes son menores que cero, pon VENCIDO, eso es rojo, ya se paso. Si son entre 0 y "
            "15, pon URGENTE, eso es amarillo, tienes que actuar ya. Si son entre 16 y 30, pon "
            "PROXIMO, naranja, ve preparandolo. Y si son mas de 30, pon OK, verde, todavia tienes "
            "tiempo. Lo mejor es que esto se combina con formato condicional para que las celdas se "
            "pinten automaticamente del color del semaforo. En el Modulo 3 vamos a profundizar en "
            "formato condicional, pero les adelanto que pueden pintar la celda de rojo si dice "
            "VENCIDO, amarillo si dice URGENTE, y asi. El resultado es un tablero visual que te dice "
            "de un vistazo que obligaciones requieren atencion inmediata. Es de las herramientas mas "
            "utiles que van a construir en este curso."
        ),
    )

    # ── Slide 18: Cumpleanos desde RFC ─────────────────────────────
    gen.add_content_slide(
        title="Caso Practico: Cumpleanos desde el RFC",
        bullets=[
            "RFC en celda A2: CAST850315HN7",
            "Extraer anio: =EXTRAE(A2, 5, 2)  -> \"85\"",
            "Definir siglo: =SI(EXTRAE(A2,5,2)*1>30, 1900, 2000)",
            "Anio completo: =SI(EXTRAE(A2,5,2)*1>30, 1900, 2000) + EXTRAE(A2,5,2)",
            "Mes: =EXTRAE(A2, 7, 2) * 1  (multiplicar por 1 convierte texto a numero)",
            "Dia: =EXTRAE(A2, 9, 2) * 1",
            "Fecha completa: =FECHA(anio_completo, mes, dia)",
            "Edad: =ENTERO((HOY()-fecha_nacimiento)/365.25)",
        ],
        script_text=(
            "Pongamos todo junto en un caso practico que es divertido y util a la vez. Vamos a "
            "extraer la fecha de cumpleanos de una persona a partir de su RFC. Tenemos el RFC "
            "CAST850315HN7 en la celda A2. Ya sabemos que las posiciones 5 y 6 son el anio, 7 y 8 "
            "el mes, 9 y 10 el dia. Pero EXTRAE nos devuelve texto, no numeros. Entonces hay un "
            "truco: si multiplicas texto que parece numero por 1, Excel lo convierte a numero. "
            "EXTRAE(A2, 5, 2) te da el texto 85. Si lo multiplicas por 1, te da el numero 85. "
            "Ahora, para definir el siglo usamos un SI: si el numero es mayor que 30, le sumamos "
            "1900; si no, le sumamos 2000. Asi, 85 se convierte en 1985, y 03 se convertiria en "
            "2003. El mes lo sacamos con EXTRAE posicion 7, 2 caracteres, multiplicado por 1. El dia "
            "igual, posicion 9, 2 caracteres, por 1. Y con la funcion FECHA juntamos todo: "
            "=FECHA(anio_completo, mes, dia). Nos da 15 de marzo de 1985. Y de pilón, podemos "
            "calcular la edad: la fecha de hoy menos la fecha de nacimiento, dividido entre 365.25, "
            "y le aplicamos ENTERO para quitar los decimales. En este ejemplo da 40 anios. Esto es "
            "un excelente ejercicio para demostrar como combinar multiples funciones en una sola "
            "solucion. Funciones de texto, funciones logicas, funciones de fecha, todo trabajando "
            "junto."
        ),
    )

    # ── Slide 19: Resumen y preview ────────────────────────────────
    gen.add_content_slide(
        title="Resumen del Modulo 1 y Preview del Modulo 2",
        bullets=[
            "SUMA, PROMEDIO: tus funciones base para todo calculo",
            "TRUNCAR vs REDONDEAR: precision fiscal obligatoria",
            "SI y SI anidado: Excel toma decisiones por ti",
            "BUSCARV: busca datos en tablas (tarifas, catalogos)",
            "Funciones de fecha: control de vencimientos y plazos",
            "EXTRAE: manipulacion de texto (RFC, CURP, claves)",
            "En el Modulo 2: Tablas Dinamicas para procesar miles de registros",
            "Preview: de 10,000 facturas a un resumen ejecutivo en 3 clics",
        ],
        script_text=(
            "Hagamos un resumen rapido de todo lo que vimos. Aprendimos que las funciones son "
            "maquinas que reciben datos, los procesan y te dan un resultado. Dominamos SUMA y "
            "PROMEDIO para calculos basicos con rangos. Entendimos la diferencia critica entre "
            "TRUNCAR y REDONDEAR, y por que en fiscalidad mexicana TRUNCAR es la regla. Aprendimos "
            "la funcion SI para que Excel tome decisiones, y vimos como anidarla para crear "
            "semaforos. Dominamos BUSCARV con sus 4 argumentos y la usamos para calcular ISR 2026 "
            "con la tarifa oficial. Vimos funciones de fecha para controlar vencimientos. Y usamos "
            "EXTRAE para desarmar el RFC y extraer informacion util. Eso es una base solida. Ahora, "
            "en el Modulo 2 vamos a subir de nivel. Vamos a hablar de Tablas Dinamicas, que son la "
            "herramienta mas poderosa de Excel para analisis de datos masivos. Imagina que tienes "
            "10 mil facturas del anio y necesitas un resumen por cliente, por mes, por tipo de "
            "producto. Con tablas dinamicas lo haces en literalmente 3 clics. Sin formulas, sin "
            "copiar y pegar, sin perder horas. Asi que nos vemos en el Modulo 2. Van a quedar "
            "impresionados."
        ),
    )

    # ── Slide 20: Cierre ───────────────────────────────────────────
    gen.add_closing_slide(
        next_module="Modulo 2 — Procesamiento Masivo y Tablas Dinamicas",
        resources=[
            "Archivo: 01_Calculadora_ISR_V2026.xlsx",
            "Archivo: 02_Control_Vencimientos_EFirma.xlsx",
            "Archivo: 03_Extraccion_RFC_Master.xlsx",
            "PDF: Referencia_Modulo_1.pdf (tarjetas de funciones y ejercicios)",
            "Grupo de WhatsApp para dudas del modulo",
        ],
    )
    gen.script_lines.append(
        "Estos son los recursos descargables del Modulo 1. Tienen tres archivos de Excel: la "
        "calculadora de ISR con la tarifa 2026, el control de vencimientos de e.firma con semaforo, "
        "y el archivo de extraccion de RFC. Ademas tienen el PDF de referencia con las tarjetas de "
        "funciones, la tarifa ISR condensada, la guia del Factor de Actualizacion y ejercicios con "
        "respuestas. Todo listo para que practiquen. Les recomiendo que descarguen todo, abran los "
        "archivos, presionen F2 en las celdas con formula y entiendan como funcionan. La practica "
        "es la clave. Nos vemos en el Modulo 2, donde las cosas se van a poner todavia mas "
        "interesantes con las Tablas Dinamicas. Hasta la proxima.\n"
    )

    gen.save()
    print(f"\n  Slides: {gen.slide_count}")


if __name__ == "__main__":
    build()
