# MÓDULO 1: Logica Contable y Funciones de Control

Israel Castro — CPA & Software Engineer — Excel para Contadores y Administrativos 2026

## Objetivo del Modulo

- Excel deja de ser calculadora, se convierte en asistente
- Dominar funciones de control: SI, BUSCARV, TRUNCAR
- Calcular ISR 2026 de forma automatica y dinamica
- Extraer informacion del RFC con funciones de texto
- Crear un calendario de vencimientos con semaforo

## Que es una Funcion? La Analogia de la Maquina de Cafe

- INPUT: los datos que le das (argumentos)
- PROCESO: lo que hace internamente (la magia)
- OUTPUT: el resultado que te entrega
- Analogia: maquina de cafe
-   - INPUT: agua + cafe molido + cantidad de tazas
-   - PROCESO: calienta, filtra, mezcla
-   - OUTPUT: tu cafe listo para tomar

## Anatomia de una Funcion en Excel

- Estructura: =NOMBRE(argumento1, argumento2, ...)
- Siempre empieza con el signo igual (=)
- El NOMBRE indica que hace la funcion
- Los ARGUMENTOS van entre parentesis, separados por coma
- Ejemplo: =SUMA(A1, A2, A3)
- Ejemplo: =SI(A1>100, "Alto", "Bajo")
- Consejo: presiona Tab para autocompletar funciones

## SUMA y PROMEDIO: Tus Primeras Funciones Contables

- =SUMA(rango) — Suma todos los valores del rango
- =SUMA(B2:B50) — Suma ventas de enero
- =PROMEDIO(rango) — Calcula la media aritmetica
- =PROMEDIO(C2:C13) — Gasto promedio mensual
- Rangos: B2:B50 (columna), B2:F2 (fila), B2:F50 (bloque)
- Atajo: Alt + = inserta SUMA automaticamente
- Tip: Nombra tus rangos (Ctrl+F3) para mayor claridad

## Ejercicio: La Caja Negra — Debuggeando Formulas

- Formula 1:  =SUMA(A1:A10)  pero A5 tiene texto "N/A"
-   -> Resultado: #VALOR!  |  Solucion: usar SUMA.SI o limpiar datos
- Formula 2:  =PROMEDIO(B1:B5)  pero B3 esta vacia
-   -> Resultado: promedia solo 4 valores  |  Cuidado con celdas vacias
- Formula 3:  =A1+B1*C1  sin parentesis
-   -> Resultado: multiplicacion va primero  |  Solucion: =(A1+B1)*C1
- Leccion: F2 para entrar a la celda y ver que esta pasando

## Precision Fiscal: TRUNCAR vs REDONDEAR

- =REDONDEAR(3.14159, 2) -> 3.14  (redondea al mas cercano)
- =TRUNCAR(3.14159, 2)   -> 3.14  (corta sin redondear)
- =REDONDEAR(3.1489, 2)  -> 3.15  (redondeo sube)
- =TRUNCAR(3.1489, 2)    -> 3.14  (truncar siempre corta)
- Art. 17-A CFF: Factor de Actualizacion se trunca al diezmilésimo (4 decimales)
- =TRUNCAR(INPC_reciente / INPC_anterior, 4)
- En fiscalidad: TRUNCAR es la regla, REDONDEAR es la excepcion

## Factor de Actualizacion con INPC 2026

- Formula: Factor = INPC reciente / INPC anterior
- INPC Dic 2025 (reciente): 141.200
- INPC Dic 2024 (anterior): 136.163
- Factor sin truncar: 141.200 / 136.163 = 1.036996...
- Factor truncado (CFF): =TRUNCAR(141.200/136.163, 4) = 1.0369
- Aplicacion: Monto original x Factor = Monto actualizado
- Ejemplo: $100,000 x 1.0369 = $103,690.00

## Coeficiente de Utilidad — Art. 14 CFF

- CU = Utilidad Fiscal / Ingresos Nominales del ejercicio anterior
- Se usa para calcular pagos provisionales de ISR
- Ejemplo: Utilidad Fiscal $450,000 / Ingresos $3,200,000
- CU = 0.140625 -> TRUNCAR a 4 dec = 0.1406
- Pago provisional = (Ingresos del periodo x CU) - Pagos anteriores
- Excel formula: =TRUNCAR(B2/B3, 4)
- BUSCARV puede traer los datos desde tu balance de comprobacion

## Funcion SI: La Toma de Decisiones en Excel

- Sintaxis: =SI(prueba_logica, valor_si_verdadero, valor_si_falso)
- Ejemplo basico: =SI(B2>10000, "Requiere autorizacion", "Aprobado")
- Con numeros: =SI(C5>=0, C5, 0)  (no permitir negativos)
- Comparadores: =, >, <, >=, <=, <>
- SI anidado: =SI(A1>100, "A", SI(A1>50, "B", "C"))
- Maximo recomendado: 3 niveles de anidacion
- Alternativa moderna: SI.CONJUNTO (Excel 365)

## Caso Practico: Validacion Bancaria con SI

- Escenario: Archivo de pagos a proveedores
- Columna A: Nombre del proveedor
- Columna B: Banco (Santander, BBVA, Banorte, Otro)
- Columna C: Monto del pago
- Columna D: Comision por transferencia SPEI
- Formula: =SI(B2="Santander", 0, SI(B2="BBVA", 4.50, 7.80))
- Resultado: Santander $0 (mismo banco), BBVA $4.50, otros $7.80
- Columna E: =C2+D2  (Total con comision)

## BUSCARV: Los 4 Argumentos en Contexto ISR

- =BUSCARV(valor_buscado, tabla, columna, [ordenado])
- Argumento 1 - valor_buscado: el ingreso gravable
- Argumento 2 - tabla: la tarifa ISR 2026 (rango o nombre)
- Argumento 3 - columna: que dato quieres (1=lim_inf, 3=cuota, 4=%)
- Argumento 4 - ordenado: VERDADERO (busqueda aproximada para rangos)
- VERDADERO: busca el valor mas cercano MENOR O IGUAL (para tarifas)
- FALSO: busca coincidencia exacta (para catalogos)
- Ejemplo ISR: =BUSCARV(450000, Tarifa_Anual_2026, 3, VERDADERO)

## Caso Practico: Calculo ISR 2026 Paso a Paso

- Paso 1: Base gravable = Ingresos - Deducciones = $450,000
- Paso 2: Limite Inferior  =BUSCARV(450000, Tarifa, 1, VERDADERO) = $424,353.98
- Paso 3: Excedente = $450,000 - $424,353.98 = $25,646.02
- Paso 4: % Marginal =BUSCARV(450000, Tarifa, 4, VERDADERO) = 23.52%
- Paso 5: ISR Marginal = $25,646.02 x 23.52% = $6,031.94
- Paso 6: Cuota Fija =BUSCARV(450000, Tarifa, 3, VERDADERO) = $67,981.92
- Paso 7: ISR Total = $6,031.94 + $67,981.92 = $74,013.86
- Todo con 3 BUSCARV y operaciones basicas

## Dinamismo: Cambia el Ingreso, Cambia Todo

- Celda de entrada (amarilla): Base Gravable = $450,000
- Cambia a $280,000 -> ISR se recalcula automaticamente
- Cambia a $1,500,000 -> ISR cambia de rango y porcentaje
- No necesitas rehacer nada: Excel recalcula en tiempo real
- Aplicacion: Planeacion fiscal con escenarios
- "Que pasa si mi cliente gana 200 mil mas?"
- "Cuanto se ahorra con 50 mil mas de deducciones?"
- Esto es lo que separa a un contador que usa Excel de uno que sabe Excel

## Funciones de Fecha: HOY, DIA, MES, ANIO

- =HOY() — Fecha actual (se actualiza sola cada dia)
- =DIA(fecha) — Extrae el dia (1-31)
- =MES(fecha) — Extrae el mes (1-12)
- =ANIO(fecha) — Extrae el anio (2026)
- =FECHA(anio, mes, dia) — Construye una fecha a partir de partes
- =DIAS(fecha_final, fecha_inicial) — Dias entre dos fechas
- Ejemplo: =DIAS("17/04/2026", HOY()) -> dias para la declaracion anual
- Las fechas en Excel son numeros seriales (1 = 1 enero 1900)

## EXTRAE: Desarmando el RFC Caracter por Caracter

- RFC persona fisica: CAST850315HN7 (13 caracteres)
- Posiciones 1-4: Apellido(s) y nombre -> CAST
- Posiciones 5-6: Anio de nacimiento -> 85 (1985)
- Posiciones 7-8: Mes de nacimiento -> 03 (marzo)
- Posiciones 9-10: Dia de nacimiento -> 15
- =EXTRAE(RFC, 5, 2) -> "85" (anio)
- =EXTRAE(RFC, 7, 2) -> "03" (mes)
- =EXTRAE(RFC, 9, 2) -> "15" (dia)
- Combinando: =FECHA(1900+EXTRAE(A1,5,2), EXTRAE(A1,7,2), EXTRAE(A1,9,2))

## Calendario de Vencimientos con Semaforo (SI Anidado)

- Columna A: Obligacion fiscal (e.firma, DIOT, declaracion, etc.)
- Columna B: Fecha de vencimiento
- Columna C: Dias restantes  =B2-HOY()
- Columna D: Semaforo con SI anidado:
-   =SI(C2<0, "VENCIDO", SI(C2<=15, "URGENTE", SI(C2<=30, "PROXIMO", "OK")))
- Rojo: VENCIDO (dias < 0)  |  Amarillo: URGENTE (0-15 dias)
- Naranja: PROXIMO (16-30 dias)  |  Verde: OK (> 30 dias)
- Formato condicional para colorear automaticamente

## Caso Practico: Cumpleanos desde el RFC

- RFC en celda A2: CAST850315HN7
- Extraer anio: =EXTRAE(A2, 5, 2)  -> "85"
- Definir siglo: =SI(EXTRAE(A2,5,2)*1>30, 1900, 2000)
- Anio completo: =SI(EXTRAE(A2,5,2)*1>30, 1900, 2000) + EXTRAE(A2,5,2)
- Mes: =EXTRAE(A2, 7, 2) * 1  (multiplicar por 1 convierte texto a numero)
- Dia: =EXTRAE(A2, 9, 2) * 1
- Fecha completa: =FECHA(anio_completo, mes, dia)
- Edad: =ENTERO((HOY()-fecha_nacimiento)/365.25)

## Resumen del Modulo 1 y Preview del Modulo 2

- SUMA, PROMEDIO: tus funciones base para todo calculo
- TRUNCAR vs REDONDEAR: precision fiscal obligatoria
- SI y SI anidado: Excel toma decisiones por ti
- BUSCARV: busca datos en tablas (tarifas, catalogos)
- Funciones de fecha: control de vencimientos y plazos
- EXTRAE: manipulacion de texto (RFC, CURP, claves)
- En el Modulo 2: Tablas Dinamicas para procesar miles de registros
- Preview: de 10,000 facturas a un resumen ejecutivo en 3 clics

## Recursos y Siguiente Paso

- Archivo: 01_Calculadora_ISR_V2026.xlsx
- Archivo: 02_Control_Vencimientos_EFirma.xlsx
- Archivo: 03_Extraccion_RFC_Master.xlsx
- PDF: Referencia_Modulo_1.pdf (tarjetas de funciones y ejercicios)
- Grupo de WhatsApp para dudas del modulo
- Siguiente: Modulo 2 — Procesamiento Masivo y Tablas Dinamicas

*Excel para Contadores y Administrativos — Israel Castro*
