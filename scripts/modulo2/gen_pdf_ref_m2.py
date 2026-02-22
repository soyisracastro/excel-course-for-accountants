"""
Generador: Referencia_Modulo_2.pdf (~5 paginas)
Modulo 2 -- Procesamiento Masivo y Analisis con Tablas Dinamicas

Contenido:
  - Tabla vs Rango: comparacion
  - Checklist de limpieza de datos
  - Diagrama de zonas de campos de Tabla Dinamica
  - 3 ejercicios practicos
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK, MODULOS
from scripts.generators.md_gen import MarkdownGenerator

inch = 72  # compat: MarkdownGenerator ignores col_widths

OUTPUT_DIR = PACK / "Modulo_2_Tablas_Dinamicas"
M2 = MODULOS[2]


def build():
    pdf = MarkdownGenerator(
        filename="Referencia_Modulo_2.md",
        output_dir=OUTPUT_DIR,
        title="Referencia Modulo 2 - Tablas Dinamicas"
    )

    # ── Portada ─────────────────────────────────────────────────
    pdf.add_cover(
        title="Guia de Referencia",
        subtitle="Tablas, Limpieza de Datos y Tablas Dinamicas",
        modulo="Modulo 2: " + M2["nombre"]
    )

    # ================================================================
    # SECCION 1: TABLA vs RANGO
    # ================================================================
    pdf.add_section("1. Tabla vs Rango: Comparacion")
    pdf.add_text(
        "Comprender la diferencia entre un rango de celdas y una Tabla de Excel "
        "es fundamental para trabajar de forma eficiente. A continuacion se "
        "comparan las caracteristicas principales."
    )
    pdf.add_spacer(0.15)

    tabla_comparacion = [
        ["Caracteristica", "Rango (celdas sueltas)", "Tabla (Ctrl+T)"],
        ["Nombre", "Solo referencia A1:H200", "Nombre descriptivo: nomina_2025"],
        ["Encabezados al scroll", "Desaparecen al bajar", "Se fijan automaticamente"],
        ["Formulas nuevas filas", "Hay que copiar manualmente", "Se expanden solas"],
        ["Referencias en formulas", "=SUMA(C2:C500)", "=SUMA(Tabla[Ventas])"],
        ["Filtros", "Configurar manualmente", "Incluidos por defecto"],
        ["Formato alterno", "Aplicar a mano", "Automatico (filas alternadas)"],
        ["Fila de totales", "Crear manualmente", "Un clic en Diseno > Fila de totales"],
        ["Fuente para Pivots", "Puede perder filas nuevas", "Siempre incluye todo"],
        ["Ordenar", "Riesgo de desalinear columnas", "Seguro: ordena filas completas"],
    ]
    pdf.add_table(tabla_comparacion, col_widths=[120, 170, 190])

    pdf.add_spacer(0.15)
    pdf.add_text(
        "<b>Regla de oro:</b> Si tus datos tienen encabezados y mas de una fila, "
        "conviertelos en Tabla con Ctrl+T. No hay razon para no hacerlo."
    )

    # ================================================================
    # SECCION 2: CHECKLIST DE LIMPIEZA
    # ================================================================
    pdf.add_page_break()
    pdf.add_section("2. Checklist de Limpieza de Datos")
    pdf.add_text(
        "Antes de crear cualquier tabla dinamica o reporte, verifica que tus datos "
        "cumplan con estos criterios. Usa esta lista como verificacion rapida."
    )
    pdf.add_spacer(0.1)

    checklist_items = [
        "Sin filas completamente vacias entre los datos",
        "Encabezados en la primera fila (sin celdas combinadas en encabezados)",
        "Una sola fila de encabezados (no dos ni tres)",
        "Cada columna contiene un solo tipo de dato (no mezclar texto y numeros)",
        "Fechas en formato fecha real (no texto que parece fecha)",
        "Montos como numeros reales (sin signos '$' ni comas como texto)",
        "Sin espacios al inicio o final de textos (usar ESPACIOS/TRIM)",
        "Sin caracteres no imprimibles (usar LIMPIAR/CLEAN)",
        "RFCs y claves con longitud correcta (usar LARGO/LEN para verificar)",
        "Sin filas duplicadas (usar Datos > Quitar duplicados)",
        "Formatos de fecha consistentes (todos DD/MM/AAAA o todos AAAA-MM-DD)",
        "Nombres estandarizados (no 'SA de CV' y 'S.A. de C.V.' mezclados)",
        "Subtotal + IVA = Total (verificar con columna auxiliar)",
        "Datos convertidos a Tabla (Ctrl+T) antes de crear pivots",
    ]

    for item in checklist_items:
        pdf.add_bullet(item)

    pdf.add_spacer(0.15)

    herramientas = [
        ["Herramienta", "Atajo / Ubicacion", "Uso principal"],
        ["Buscar y Reemplazar", "Ctrl+H", "Quitar $, comas, espacios masivos"],
        ["ESPACIOS (TRIM)", "=ESPACIOS(celda)", "Eliminar espacios extra"],
        ["LIMPIAR (CLEAN)", "=LIMPIAR(celda)", "Quitar caracteres no imprimibles"],
        ["VALOR (VALUE)", "=VALOR(celda)", "Convertir texto a numero"],
        ["Texto en Columnas", "Datos > Texto en columnas", "Separar datos concatenados"],
        ["Quitar Duplicados", "Datos > Quitar duplicados", "Eliminar filas repetidas"],
        ["Pegado Especial", "Ctrl+Alt+V > Valores", "Convertir formulas a valores"],
        ["SUSTITUIR", "=SUSTITUIR(texto,viejo,nuevo)", "Reemplazar texto especifico"],
    ]
    pdf.add_table(herramientas, col_widths=[110, 140, 200])

    # ================================================================
    # SECCION 3: ZONAS DE CAMPOS DE TABLA DINAMICA
    # ================================================================
    pdf.add_page_break()
    pdf.add_section("3. Zonas de Campos de Tabla Dinamica")
    pdf.add_text(
        "Al crear una Tabla Dinamica, el panel 'Campos de tabla dinamica' muestra "
        "cuatro zonas donde se arrastran los campos. Cada zona tiene una funcion "
        "especifica:"
    )
    pdf.add_spacer(0.15)

    zonas = [
        ["Zona", "Ubicacion en el Pivot", "Que poner aqui", "Ejemplo Nomina"],
        ["FILTROS", "Arriba del pivot (dropdown)", "Campos para restringir el analisis",
         "Periodo, Anio, Sucursal"],
        ["COLUMNAS", "Encabezados horizontales", "Categorias para desglosar (pocas)",
         "Concepto, Clase"],
        ["FILAS", "Etiquetas verticales", "El eje principal de analisis",
         "NombreEmpleado, Puesto"],
        ["VALORES", "Cuerpo de la tabla (numeros)", "Campos numericos a calcular",
         "Suma de ImporteGravado"],
    ]
    pdf.add_table(zonas, col_widths=[60, 120, 140, 130])

    pdf.add_spacer(0.2)
    pdf.add_subsection("Diagrama visual del panel de campos")
    pdf.add_text(
        "El panel de campos tiene dos secciones: arriba la lista de todos los campos "
        "disponibles, abajo las cuatro zonas. Simplemente arrastra campos de la lista "
        "a la zona deseada."
    )
    pdf.add_spacer(0.1)

    # Diagrama en formato tabla
    diagrama = [
        ["CAMPOS DISPONIBLES"],
        ["UUID | NumEmpleado | NombreEmpleado | Puesto | FechaPago | Periodo | "
         "Clase | Concepto | ImporteGravado | ImporteExento"],
    ]
    pdf.add_table(diagrama, col_widths=[480], header=True)

    pdf.add_spacer(0.1)

    layout = [
        ["FILTROS", "COLUMNAS"],
        ["Periodo", "Concepto"],
        ["FILAS", "VALORES"],
        ["NombreEmpleado", "Suma de ImporteGravado"],
    ]
    pdf.add_table(layout, col_widths=[240, 240], header=False)

    pdf.add_spacer(0.15)
    pdf.add_subsection("Tipos de calculo en Valores")
    pdf.add_text(
        "Por defecto, Excel usa SUMA para numeros y CUENTA para texto. "
        "Puedes cambiar el tipo de calculo haciendo clic en el campo dentro de "
        "la zona Valores > Configuracion de campo de valor."
    )

    calculos = [
        ["Funcion", "Cuando usarla", "Ejemplo"],
        ["Suma", "Totalizar montos", "Total de percepciones por empleado"],
        ["Cuenta", "Contar registros", "Numero de conceptos por empleado"],
        ["Promedio", "Obtener media", "Sueldo promedio por puesto"],
        ["Max / Min", "Encontrar extremos", "Sueldo mas alto / mas bajo"],
        ["% del total general", "Ver proporciones", "Que % del ISR paga cada empleado"],
    ]
    pdf.add_table(calculos, col_widths=[100, 170, 200])

    # ================================================================
    # SECCION 4: EJERCICIOS PRACTICOS
    # ================================================================
    pdf.add_page_break()
    pdf.add_section("4. Ejercicios Practicos")

    # Ejercicio 1
    pdf.add_subsection("Ejercicio 1: Limpieza de datos de compras")
    pdf.add_text("<b>Archivo:</b> 04_Limpieza_Masiva_Layout.xlsx")
    pdf.add_text("<b>Objetivo:</b> Limpiar 220 filas de datos de compras con errores intencionales.")
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Pasos:</b>")
    pdf.add_bullet("Abre la hoja 'Datos_Sucios' y revisa los tipos de errores")
    pdf.add_bullet("Usa Ctrl+H para eliminar signos '$' en la columna Subtotal")
    pdf.add_bullet("Aplica ESPACIOS(TRIM) en una columna auxiliar para limpiar RFC")
    pdf.add_bullet("Convierte las fechas texto a formato fecha con DATEVALUE o Texto en Columnas")
    pdf.add_bullet("Verifica Subtotal + IVA = Total con una columna de validacion")
    pdf.add_bullet("Usa Datos > Quitar duplicados en la columna Folio")
    pdf.add_bullet("Convierte el resultado limpio a Tabla con Ctrl+T")
    pdf.add_bullet("Compara con la hoja 'Datos_Limpios' para verificar tu trabajo")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Tiempo estimado:</b> 15-20 minutos. "
        "<b>Reto:</b> Hazlo en menos de 10 minutos usando atajos."
    )

    pdf.add_spacer(0.2)

    # Ejercicio 2
    pdf.add_subsection("Ejercicio 2: Tres Tablas Dinamicas de Nomina")
    pdf.add_text("<b>Archivo:</b> 05_Analisis_Nomina_XML_Pivot.xlsx")
    pdf.add_text(
        "<b>Objetivo:</b> Crear 3 tablas dinamicas para analizar percepciones, "
        "deducciones y costo total de nomina."
    )
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Pivot 1 - Percepciones:</b>")
    pdf.add_bullet("Filtrar Clase = Percepcion")
    pdf.add_bullet("Filas: NombreEmpleado | Columnas: Concepto | Valores: Suma ImporteGravado")
    pdf.add_bullet("Agregar Slicer por Periodo")
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Pivot 2 - Deducciones:</b>")
    pdf.add_bullet("Filtrar Clase = Deduccion")
    pdf.add_bullet("Filas: NombreEmpleado + Puesto | Columnas: Concepto | Valores: Suma ImporteGravado")
    pdf.add_bullet("Pregunta: Quien paga mas ISR y por que?")
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Pivot 3 - Costo Total:</b>")
    pdf.add_bullet("Sin filtro de Clase")
    pdf.add_bullet("Filas: NombreEmpleado | Columnas: Clase | Valores: Suma ImporteGravado + ImporteExento")
    pdf.add_bullet("Ordenar de mayor a menor costo total")
    pdf.add_spacer(0.1)
    pdf.add_text("<b>Tiempo estimado:</b> 20-25 minutos.")

    pdf.add_spacer(0.2)

    # Ejercicio 3
    pdf.add_subsection("Ejercicio 3: Papel de Trabajo ISR con BUSCARV")
    pdf.add_text("<b>Archivo:</b> 06_Papel_Trabajo_Referenciado.xlsx")
    pdf.add_text(
        "<b>Objetivo:</b> Usar los resultados de las tablas dinamicas para alimentar "
        "el papel de trabajo de declaracion anual ISR."
    )
    pdf.add_spacer(0.05)
    pdf.add_text("<b>Pasos:</b>")
    pdf.add_bullet("Del Pivot 1, obtener el total de percepciones gravadas de un empleado")
    pdf.add_bullet("Capturar ese monto en la celda 'Sueldos y salarios gravados' del papel")
    pdf.add_bullet("Capturar deducciones personales ficticias (gastos medicos, colegiaturas)")
    pdf.add_bullet("Observar como las formulas BUSCARV calculan automaticamente el ISR")
    pdf.add_bullet("Del Pivot 2, obtener el total de ISR retenido del mismo empleado")
    pdf.add_bullet("Capturarlo en 'Retenciones de ISR en el ejercicio'")
    pdf.add_bullet("Verificar: si ISR causado < retenciones = saldo a favor (devolucion)")
    pdf.add_spacer(0.1)
    pdf.add_text(
        "<b>Tiempo estimado:</b> 10-15 minutos. "
        "<b>Valor:</b> Este es el flujo real de trabajo en un despacho contable."
    )

    # ================================================================
    # SECCION 5: RESUMEN Y ATAJOS
    # ================================================================
    pdf.add_page_break()
    pdf.add_section("5. Atajos Clave del Modulo 2")

    atajos = [
        ["Atajo", "Accion", "Cuando usarlo"],
        ["Ctrl+T", "Convertir rango a Tabla", "Siempre que tengas datos con encabezados"],
        ["Ctrl+H", "Buscar y Reemplazar", "Limpieza masiva de caracteres"],
        ["Ctrl+Shift+L", "Activar/desactivar filtros", "Filtrar datos rapidamente"],
        ["Alt+N+V", "Insertar Tabla Dinamica", "Crear un nuevo pivot"],
        ["Alt+F5", "Actualizar Tabla Dinamica", "Despues de cambiar datos fuente"],
        ["Ctrl+A", "Seleccionar todo", "Antes de aplicar formato"],
        ["Ctrl+1", "Formato de celdas", "Cambiar formato de numeros/fechas"],
        ["Ctrl+Z", "Deshacer", "Si algo sale mal (hasta 100 veces)"],
        ["F2", "Editar celda", "Ver/modificar formulas"],
        ["Ctrl+`", "Mostrar formulas", "Ver todas las formulas de la hoja"],
    ]
    pdf.add_table(atajos, col_widths=[90, 170, 210])

    pdf.add_spacer(0.3)
    pdf.add_section("Notas Finales")
    pdf.add_text(
        "Las tablas dinamicas son la herramienta de analisis mas poderosa de Excel. "
        "Combinadas con datos limpios y bien estructurados en tablas, permiten "
        "responder cualquier pregunta de negocio en segundos. "
        "Practica con los tres archivos de este modulo hasta que el proceso sea "
        "automatico: limpiar, tabular, pivotear, analizar."
    )
    pdf.add_spacer(0.1)
    pdf.add_text(
        "En el <b>Modulo 3</b> convertiremos estos analisis en graficos profesionales, "
        "formato condicional avanzado y reportes ejecutivos listos para presentar."
    )

    pdf.save()


if __name__ == "__main__":
    build()
