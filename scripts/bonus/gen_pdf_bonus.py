"""
Generador de Markdowns Bonus:
  1. Guia_VBA_con_IA.md — Plantillas de prompts y 5 macros listas para copiar
  2. Guia_Claude_en_Excel.md — Guia de instalacion, prompts y comparativa

Salida: output/Pack_Excel_Pro/Bonus/
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.generators.md_gen import MarkdownGenerator

OUTPUT_DIR = PACK / "Bonus"


# =====================================================================
# PDF 1 — Guia VBA con IA
# =====================================================================

def _build_vba_guide():
    pdf = MarkdownGenerator(
        filename="Guia_VBA_con_IA.md",
        output_dir=OUTPUT_DIR,
        title="Guia VBA con IA - Macros Listas para Copiar",
    )

    # ── Portada ───────────────────────────────────────────────────
    pdf.add_cover(
        title="Guia VBA con IA",
        subtitle="5 macros listas para copiar + plantillas de prompts",
        modulo="Bonus - Material Complementario",
    )

    # ── Seccion 1: Introduccion ───────────────────────────────────
    pdf.add_section("Introduccion")
    pdf.add_text(
        "Esta guia contiene 5 macros VBA listas para copiar y pegar en Excel. "
        "Cada macro incluye el codigo completo, instrucciones de uso, y el prompt "
        "exacto que puedes usar en ChatGPT o Claude para generar o modificar la macro."
    )
    pdf.add_spacer(0.1)
    pdf.add_text(
        "**Requisitos:** Excel 2019, 2021, o Microsoft 365 en Windows. "
        "Guardar el archivo como .xlsm (habilitado para macros). "
        "Abrir el editor VBA con Alt + F11, insertar un Modulo, y pegar el codigo."
    )

    # ── Seccion 2: Plantillas de Prompts ──────────────────────────
    pdf.add_section("Plantillas de Prompts para Generar VBA")
    pdf.add_text(
        "Usa estas plantillas como punto de partida para pedirle a la IA "
        "que genere macros personalizadas."
    )
    pdf.add_spacer(0.1)

    prompts = [
        (
            "Prompt basico",
            "Genera una macro VBA para Excel que [TAREA]. "
            "Los datos estan en la hoja [NOMBRE] en el rango [RANGO]. "
            "Incluye comentarios en espaniol. Compatible con Excel [VERSION]."
        ),
        (
            "Prompt con formato",
            "Crea una macro VBA que formatee la hoja activa: separador de miles "
            "en columnas [X:Y], bordes delgados en todo el rango, encabezado con "
            "fondo azul y texto blanco. Autoajustar ancho de columnas."
        ),
        (
            "Prompt de limpieza",
            "Genera una macro VBA que limpie datos: elimine filas vacias en la "
            "seleccion actual, quite espacios extra con Trim, y convierta texto "
            "a mayusculas en la columna [X]."
        ),
        (
            "Prompt de reporte",
            "Crea una macro VBA que genere un reporte mensual: nueva hoja con "
            "nombre 'Reporte_[Mes]_[Anio]', copie la estructura de la hoja "
            "Plantilla, e inserte la fecha de generacion en A1."
        ),
        (
            "Prompt de debugging",
            "Este codigo VBA da el error '[ERROR]' en la linea [N]. "
            "Explicame que causa el error y dame el codigo corregido. "
            "El codigo es: [PEGAR CODIGO]"
        ),
    ]

    for titulo, texto in prompts:
        pdf.add_subsection(titulo)
        pdf.add_code(texto)
        pdf.add_spacer(0.1)

    pdf.add_page_break()

    # ── Macros ────────────────────────────────────────────────────

    # -- Macro 1: FormatearNomina --
    pdf.add_section("Macro 1: FormatearNomina()")
    pdf.add_subsection("Proposito")
    pdf.add_text(
        "Aplica formato profesional a una tabla de nomina: separador de miles "
        "con 2 decimales en columnas monetarias (E:H), bordes delgados en todo "
        "el rango, encabezado con fondo azul y texto blanco, y autoajuste de columnas."
    )
    pdf.add_subsection("Codigo VBA")
    pdf.add_code(
        'Sub FormatearNomina()\n'
        '    Dim ws As Worksheet\n'
        '    Dim rng As Range\n'
        '    Dim lastRow As Long, lastCol As Long\n'
        '    \n'
        '    Set ws = ActiveSheet\n'
        '    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row\n'
        '    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column\n'
        '    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))\n'
        '    \n'
        "    ' Formato numerico a columnas monetarias (E:H)\n"
        '    ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 8)).NumberFormat = "#,##0.00"\n'
        '    \n'
        "    ' Bordes delgados\n"
        '    rng.Borders.LineStyle = xlContinuous\n'
        '    rng.Borders.Weight = xlThin\n'
        '    \n'
        "    ' Encabezado azul\n"
        '    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))\n'
        '        .Interior.Color = RGB(37, 99, 235)\n'
        '        .Font.Color = RGB(255, 255, 255)\n'
        '        .Font.Bold = True\n'
        '    End With\n'
        '    \n'
        "    ' Autoajuste\n"
        '    rng.Columns.AutoFit\n'
        '    \n'
        '    MsgBox "Nomina formateada correctamente.", vbInformation\n'
        'End Sub'
    )
    pdf.add_subsection("Como usar")
    pdf.add_bullet("Abrir el archivo de nomina en Excel")
    pdf.add_bullet("Alt + F11 -> Insertar -> Modulo -> Pegar el codigo")
    pdf.add_bullet("Posicionar el cursor dentro de Sub FormatearNomina()")
    pdf.add_bullet("Presionar F5 para ejecutar")
    pdf.add_subsection("Prompt para generarla")
    pdf.add_code(
        "Genera una macro VBA llamada FormatearNomina para Excel. Los datos estan "
        "en la hoja activa. La macro debe: aplicar formato #,##0.00 a las columnas "
        "E:H desde fila 2, bordes delgados a todo el rango con datos, pintar la "
        "fila 1 de azul RGB(37,99,235) con texto blanco y negrita, y autoajustar "
        "columnas. Mostrar MsgBox al terminar. Comentarios en espaniol."
    )

    pdf.add_page_break()

    # -- Macro 2: LimpiarVacias --
    pdf.add_section("Macro 2: LimpiarVacias()")
    pdf.add_subsection("Proposito")
    pdf.add_text(
        "Elimina filas completamente vacias dentro de la seleccion actual del usuario. "
        "Recorre de abajo hacia arriba para evitar saltar filas al eliminar."
    )
    pdf.add_subsection("Codigo VBA")
    pdf.add_code(
        'Sub LimpiarVacias()\n'
        '    Dim rng As Range\n'
        '    Dim i As Long\n'
        '    \n'
        '    Set rng = Selection\n'
        '    \n'
        "    ' Recorrer de abajo hacia arriba\n"
        '    For i = rng.Rows.Count To 1 Step -1\n'
        '        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then\n'
        '            rng.Rows(i).EntireRow.Delete\n'
        '        End If\n'
        '    Next i\n'
        '    \n'
        '    MsgBox "Filas vacias eliminadas.", vbInformation\n'
        'End Sub'
    )
    pdf.add_subsection("Como usar")
    pdf.add_bullet("Seleccionar el rango de datos con posibles filas vacias")
    pdf.add_bullet("Ejecutar la macro con F5 desde el editor VBA")
    pdf.add_bullet("Verificar que las filas vacias fueron eliminadas")
    pdf.add_subsection("Prompt para generarla")
    pdf.add_code(
        "Genera una macro VBA llamada LimpiarVacias. Debe recorrer la seleccion "
        "actual del usuario de abajo hacia arriba. Si una fila esta completamente "
        "vacia (CountA = 0), eliminarla con EntireRow.Delete. Mostrar MsgBox al "
        "terminar. Comentarios en espaniol."
    )

    pdf.add_page_break()

    # -- Macro 3: ReporteMensual --
    pdf.add_section("Macro 3: ReporteMensual()")
    pdf.add_subsection("Proposito")
    pdf.add_text(
        "Crea una nueva hoja con nombre automatico basado en el mes y anio actual "
        "(ej: Reporte_Feb_2026). Verifica que no exista previamente. Inserta "
        "encabezado con fecha de generacion."
    )
    pdf.add_subsection("Codigo VBA")
    pdf.add_code(
        'Sub ReporteMensual()\n'
        '    Dim ws As Worksheet\n'
        '    Dim nombreHoja As String\n'
        '    Dim meses As Variant\n'
        '    \n'
        '    meses = Array("Ene", "Feb", "Mar", "Abr", "May", "Jun", _\n'
        '                   "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")\n'
        '    \n'
        '    nombreHoja = "Reporte_" & meses(Month(Date) - 1) & "_" & Year(Date)\n'
        '    \n'
        "    ' Verificar si ya existe\n"
        '    On Error Resume Next\n'
        '    Set ws = ThisWorkbook.Sheets(nombreHoja)\n'
        '    On Error GoTo 0\n'
        '    \n'
        '    If Not ws Is Nothing Then\n'
        '        MsgBox "La hoja " & nombreHoja & " ya existe.", vbExclamation\n'
        '        Exit Sub\n'
        '    End If\n'
        '    \n'
        "    ' Crear nueva hoja\n"
        '    Set ws = ThisWorkbook.Sheets.Add( _\n'
        '        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))\n'
        '    ws.Name = nombreHoja\n'
        '    \n'
        "    ' Encabezado\n"
        '    ws.Range("A1").Value = "Reporte Mensual - " & Format(Date, "MMMM YYYY")\n'
        '    ws.Range("A2").Value = "Generado: " & Format(Now, "DD/MM/YYYY HH:MM")\n'
        '    ws.Range("A1").Font.Bold = True\n'
        '    ws.Range("A1").Font.Size = 14\n'
        '    \n'
        '    MsgBox "Hoja \'" & nombreHoja & "\' creada.", vbInformation\n'
        'End Sub'
    )
    pdf.add_subsection("Como usar")
    pdf.add_bullet("Abrir el libro donde se necesita el reporte mensual")
    pdf.add_bullet("Ejecutar la macro; crea la hoja automaticamente")
    pdf.add_bullet("Si la hoja ya existe, muestra advertencia y no la sobrescribe")
    pdf.add_subsection("Prompt para generarla")
    pdf.add_code(
        "Genera una macro VBA llamada ReporteMensual. Debe crear una hoja nueva "
        "con nombre Reporte_[Mes]_[Anio] usando la fecha actual. Verificar que la "
        "hoja no exista antes de crearla. En A1 poner titulo con formato MMMM YYYY, "
        "en A2 poner fecha y hora de generacion. A1 en negrita tamano 14. MsgBox "
        "al terminar. Comentarios en espaniol."
    )

    pdf.add_page_break()

    # -- Macro 4: ActualizarPivots --
    pdf.add_section("Macro 4: ActualizarPivots()")
    pdf.add_subsection("Proposito")
    pdf.add_text(
        "Recorre todas las hojas del libro y actualiza cada tabla dinamica encontrada. "
        "Muestra un conteo de cuantas tablas fueron actualizadas."
    )
    pdf.add_subsection("Codigo VBA")
    pdf.add_code(
        'Sub ActualizarPivots()\n'
        '    Dim ws As Worksheet\n'
        '    Dim pt As PivotTable\n'
        '    Dim contador As Long\n'
        '    \n'
        '    contador = 0\n'
        '    \n'
        '    For Each ws In ThisWorkbook.Worksheets\n'
        '        For Each pt In ws.PivotTables\n'
        '            pt.RefreshTable\n'
        '            contador = contador + 1\n'
        '        Next pt\n'
        '    Next ws\n'
        '    \n'
        '    MsgBox contador & " tabla(s) dinamica(s) actualizada(s).", vbInformation\n'
        'End Sub'
    )
    pdf.add_subsection("Como usar")
    pdf.add_bullet("Tener un libro con una o mas tablas dinamicas")
    pdf.add_bullet("Ejecutar la macro; actualiza TODAS las tablas de todas las hojas")
    pdf.add_bullet("Asignar a un boton: Insertar > Formas > clic derecho > Asignar macro")
    pdf.add_subsection("Prompt para generarla")
    pdf.add_code(
        "Genera una macro VBA llamada ActualizarPivots. Debe recorrer todas las "
        "hojas del libro activo, y para cada hoja recorrer todas sus tablas dinamicas "
        "y ejecutar RefreshTable. Contar cuantas tablas se actualizaron y mostrarlo "
        "en un MsgBox. Comentarios en espaniol."
    )

    pdf.add_page_break()

    # -- Macro 5: ExportarPDF --
    pdf.add_section("Macro 5: ExportarPDF()")
    pdf.add_subsection("Proposito")
    pdf.add_text(
        "Exporta la hoja activa como archivo PDF. El nombre del archivo incluye "
        "el nombre de la hoja y la fecha actual. Guarda en la misma carpeta del libro."
    )
    pdf.add_subsection("Codigo VBA")
    pdf.add_code(
        'Sub ExportarPDF()\n'
        '    Dim ws As Worksheet\n'
        '    Dim rutaPDF As String\n'
        '    Dim rutaLibro As String\n'
        '    \n'
        '    Set ws = ActiveSheet\n'
        '    rutaLibro = ThisWorkbook.Path\n'
        '    \n'
        '    If rutaLibro = "" Then\n'
        '        MsgBox "Guarde el libro primero.", vbExclamation\n'
        '        Exit Sub\n'
        '    End If\n'
        '    \n'
        '    rutaPDF = rutaLibro & "\\" & ws.Name & "_" & _\n'
        '             Format(Date, "YYYY-MM-DD") & ".pdf"\n'
        '    \n'
        '    ws.ExportAsFixedFormat _\n'
        '        Type:=xlTypePDF, _\n'
        '        Filename:=rutaPDF, _\n'
        '        Quality:=xlQualityStandard, _\n'
        '        IncludeDocProperties:=True, _\n'
        '        OpenAfterPublish:=True\n'
        '    \n'
        '    MsgBox "PDF exportado: " & rutaPDF, vbInformation\n'
        'End Sub'
    )
    pdf.add_subsection("Como usar")
    pdf.add_bullet("Abrir la hoja que desea exportar como PDF")
    pdf.add_bullet("Configurar area de impresion si es necesario (Vista > Saltos de pagina)")
    pdf.add_bullet("Ejecutar la macro; el PDF se guarda en la misma carpeta del libro")
    pdf.add_bullet("El archivo se abre automaticamente despues de exportar")
    pdf.add_subsection("Prompt para generarla")
    pdf.add_code(
        "Genera una macro VBA llamada ExportarPDF. Debe exportar la hoja activa "
        "como PDF usando ExportAsFixedFormat. El nombre del archivo debe ser "
        "[NombreHoja]_[Fecha].pdf y guardarse en la misma carpeta del libro. "
        "Verificar que el libro este guardado antes de exportar. Abrir el PDF "
        "despues de crearlo. MsgBox con la ruta. Comentarios en espaniol."
    )

    # ── Seccion final: Tips ───────────────────────────────────────
    pdf.add_page_break()
    pdf.add_section("Tips Finales")
    pdf.add_bullet(
        "**Personal.xlsb:** Guarda tus macros mas utiles en el libro Personal "
        "(Archivo > Opciones > Guardar macro en > Libro de macros personal). "
        "Asi estan disponibles en todos tus archivos."
    )
    pdf.add_bullet(
        "**Seguridad:** Solo habilita macros de fuentes confiables. "
        "Configura el Trust Center en: Archivo > Opciones > Centro de confianza."
    )
    pdf.add_bullet(
        "**Respaldo:** Siempre prueba macros en una copia del archivo, nunca en el original."
    )
    pdf.add_bullet(
        "**Errores:** Si una macro da error, copia el mensaje de error y el codigo, "
        "pegalo en Claude o ChatGPT, y pide explicacion y correccion."
    )
    pdf.add_bullet(
        "**Documentacion:** Agrega comentarios (lineas con ') a tus macros para "
        "que tu yo futuro entienda que hacen."
    )

    pdf.save()


# =====================================================================
# PDF 2 — Guia Claude en Excel
# =====================================================================

def _build_claude_guide():
    pdf = MarkdownGenerator(
        filename="Guia_Claude_en_Excel.md",
        output_dir=OUTPUT_DIR,
        title="Guia Claude en Excel - Tu Segundo Cerebro Contable",
    )

    # ── Portada ───────────────────────────────────────────────────
    pdf.add_cover(
        title="Guia Claude en Excel",
        subtitle="Instalacion, prompts contables y mejores practicas",
        modulo="Bonus - Material Complementario",
    )

    # ── Seccion 1: Instalacion ────────────────────────────────────
    pdf.add_section("Guia de Instalacion del Add-in Claude para Excel")
    pdf.add_text(
        "Desde enero de 2026, Claude esta disponible como add-in oficial para "
        "suscriptores de Microsoft 365 Pro. La instalacion es directa desde el "
        "Marketplace de Microsoft."
    )
    pdf.add_spacer(0.1)
    pdf.add_subsection("Requisitos")
    pdf.add_bullet("Microsoft 365 Pro (suscripcion activa)")
    pdf.add_bullet("Excel para Windows (version 2312 o superior) o Excel para Web")
    pdf.add_bullet("Conexion a internet para la instalacion y uso del add-in")
    pdf.add_bullet("Cuenta de Microsoft vinculada a la suscripcion M365")
    pdf.add_spacer(0.1)

    pdf.add_subsection("Pasos de Instalacion")
    pdf.add_bullet("**Paso 1:** Abrir Excel y crear o abrir cualquier libro")
    pdf.add_bullet("**Paso 2:** Ir a la pestana Insertar en la cinta de opciones")
    pdf.add_bullet("**Paso 3:** Clic en 'Obtener complementos' (o 'Get Add-ins')")
    pdf.add_bullet("**Paso 4:** En el cuadro de busqueda, escribir 'Claude for Excel'")
    pdf.add_bullet("**Paso 5:** Seleccionar el add-in oficial de Anthropic y clic en 'Agregar'")
    pdf.add_bullet("**Paso 6:** Aceptar los permisos solicitados")
    pdf.add_bullet("**Paso 7:** El panel de Claude aparecera en la pestana Inicio")
    pdf.add_spacer(0.1)

    pdf.add_subsection("Verificacion")
    pdf.add_text(
        "Para verificar que la instalacion fue exitosa, abrir el panel lateral de "
        "Claude y escribir una pregunta simple como 'Hola, que puedes hacer con esta "
        "hoja?'. Claude debe responder con una descripcion de sus capacidades."
    )

    pdf.add_page_break()

    # ── Seccion 2: 10 Prompts Contables ───────────────────────────
    pdf.add_section("10 Prompts Contables para Claude en Excel")
    pdf.add_text(
        "Estos prompts estan disenados especificamente para tareas contables y "
        "administrativas. Copia y pega directamente en el panel de Claude."
    )
    pdf.add_spacer(0.1)

    prompts = [
        (
            "1. Analisis de nomina",
            "Analiza esta tabla de nomina. Identifica los empleados cuyo sueldo "
            "esta por encima del promedio de su departamento. Muestra el nombre, "
            "departamento, sueldo, y la diferencia con el promedio."
        ),
        (
            "2. Deteccion de anomalias en gastos",
            "Revisa esta tabla de gastos mensuales. Identifica movimientos que se "
            "desvien mas del 30% del promedio historico de su categoria. Senala "
            "cuales podrian requerir revision."
        ),
        (
            "3. Generacion de formula BUSCARV + SIERROR",
            "Necesito una formula que busque el RFC de la columna A en la tabla "
            "Proveedores y traiga el nombre de la columna 3. Si no lo encuentra, "
            "debe mostrar 'No encontrado'. Explicame paso a paso."
        ),
        (
            "4. Explicacion de formula heredada",
            "Explicame paso a paso que hace esta formula, como si yo fuera "
            "contador y no programador. Descompone cada funcion anidada: "
            "[PEGAR FORMULA AQUI]"
        ),
        (
            "5. Conciliacion bancaria",
            "Tengo dos tablas: Estado_Banco y Registro_Contable. Ambas tienen "
            "columnas Fecha, Referencia, y Monto. Identifica las diferencias: "
            "movimientos que estan en banco pero no en contabilidad, y viceversa."
        ),
        (
            "6. Proyeccion de flujo de efectivo",
            "Con base en los ingresos y egresos de los ultimos 6 meses en esta "
            "tabla, proyecta el flujo de efectivo para los proximos 3 meses. "
            "Usa tendencia lineal y senala meses con posible deficit."
        ),
        (
            "7. Clasificacion de cuentas contables",
            "Clasifica estos movimientos en las categorias del catalogo de cuentas: "
            "Activo, Pasivo, Capital, Ingreso, Costo, Gasto. Agrega la clasificacion "
            "en una columna nueva."
        ),
        (
            "8. Calculo de ISR con tarifas",
            "Usando la tarifa del Art. 96 LISR vigente para 2026, calcula la "
            "retencion de ISR mensual para cada empleado de esta tabla. Muestra "
            "el desglose: base gravable, limite inferior, excedente, impuesto "
            "marginal, cuota fija, y total ISR."
        ),
        (
            "9. Resumen ejecutivo de ventas",
            "Genera un resumen ejecutivo de esta tabla de ventas: total por "
            "producto, por region, por vendedor. Identifica el top 5 de productos "
            "y el bottom 5. Sugiere acciones basadas en los datos."
        ),
        (
            "10. Auditoria de formulas",
            "Revisa todas las formulas de esta hoja. Identifica: celdas con "
            "errores (#REF!, #N/A, #VALOR!), referencias circulares, formulas "
            "que podrian simplificarse, y celdas con valores hardcodeados donde "
            "deberia haber formulas."
        ),
    ]

    for titulo, texto in prompts:
        pdf.add_subsection(titulo)
        pdf.add_code(texto)
        pdf.add_spacer(0.05)

    pdf.add_page_break()

    # ── Seccion 3: Claude vs Copilot ──────────────────────────────
    pdf.add_section("Claude vs Copilot: Tabla Comparativa")
    pdf.add_text(
        "Ambas herramientas son complementarias. Esta tabla resume las diferencias "
        "principales para ayudarte a decidir cual usar segun la tarea."
    )
    pdf.add_spacer(0.15)

    comparison_data = [
        ["Caracteristica", "Claude (Anthropic)", "Copilot (Microsoft)"],
        ["Integracion en Excel", "Add-in desde Marketplace", "Nativo en M365"],
        ["Analisis profundo de datos", "Excelente", "Bueno"],
        ["Generacion de formulas", "Excelente (con explicacion)", "Muy bueno"],
        ["Automatizacion rapida", "Buena (via prompts)", "Excelente (nativo)"],
        ["Generacion de graficos", "Sugiere configuracion", "Crea directamente"],
        ["Razonamiento complejo", "Superior", "Bueno"],
        ["Codigo VBA", "Genera y explica", "Genera"],
        ["Lenguaje natural en espaniol", "Excelente", "Muy bueno"],
        ["Privacidad de datos", "No almacena datos", "Segun plan M365"],
        ["Costo", "Incluido en M365 Pro", "Incluido en M365 Pro/Copilot"],
        ["Mejor para", "Analisis, razonamiento, auditoria", "Tareas rapidas, automatizacion"],
        ["Conectores externos (MCP)", "Si, via protocolo MCP", "Si, via Microsoft Graph"],
    ]
    pdf.add_table(comparison_data)

    pdf.add_spacer(0.2)

    pdf.add_text(
        "**Recomendacion:** Usa Copilot para tareas rapidas de automatizacion "
        "y creacion de graficos. Usa Claude para analisis profundo, razonamiento "
        "sobre datos, auditoria de formulas, y generacion de codigo VBA con "
        "explicaciones detalladas. Juntos cubren el 95% de las necesidades."
    )

    pdf.add_page_break()

    # ── Seccion 4: Privacidad y Datos ─────────────────────────────
    pdf.add_section("Privacidad y Consideraciones de Datos")
    pdf.add_text(
        "Al usar inteligencia artificial con datos empresariales, es fundamental "
        "entender como se manejan los datos."
    )
    pdf.add_spacer(0.1)

    pdf.add_subsection("Politica de datos de Claude en Excel")
    pdf.add_bullet(
        "Claude **no almacena** los datos de las hojas que procesa en las "
        "sesiones del add-in"
    )
    pdf.add_bullet(
        "Los datos se envian a los servidores de Anthropic para procesamiento "
        "y se descartan despues de generar la respuesta"
    )
    pdf.add_bullet(
        "Anthropic **no usa datos de clientes empresariales** para entrenar "
        "sus modelos (politica vigente desde 2024)"
    )
    pdf.add_bullet(
        "Para suscriptores M365 Pro, los datos transitan por la infraestructura "
        "de Microsoft Azure"
    )
    pdf.add_spacer(0.1)

    pdf.add_subsection("Recomendaciones de seguridad")
    pdf.add_bullet(
        "Consultar con el area de TI antes de usar IA con datos confidenciales"
    )
    pdf.add_bullet(
        "No enviar datos sensibles como contrasenas, numeros de tarjeta, o "
        "informacion personal identificable (PII) a menos que la politica "
        "organizacional lo permita"
    )
    pdf.add_bullet(
        "Para datos altamente confidenciales, considerar Claude en modo "
        "empresarial con politicas de retencion personalizadas"
    )
    pdf.add_bullet(
        "Siempre validar las respuestas de la IA antes de tomar decisiones "
        "basadas en ellas"
    )
    pdf.add_bullet(
        "Documentar el uso de IA en procesos contables para cumplimiento "
        "normativo y auditoria"
    )

    pdf.add_page_break()

    # ── Seccion 5: MCP Overview ───────────────────────────────────
    pdf.add_section("MCP (Model Context Protocol): Vision General")
    pdf.add_text(
        "MCP es un protocolo abierto creado por Anthropic que permite a Claude "
        "conectarse a fuentes de datos externas de forma segura y estandarizada."
    )
    pdf.add_spacer(0.1)

    pdf.add_subsection("Que es MCP?")
    pdf.add_bullet(
        "MCP = Model Context Protocol (Protocolo de Contexto de Modelo)"
    )
    pdf.add_bullet(
        "Permite que Claude acceda a datos fuera de la hoja de Excel: "
        "bases de datos, APIs, archivos en la nube"
    )
    pdf.add_bullet(
        "Protocolo abierto y estandarizado: cualquier proveedor puede crear conectores"
    )
    pdf.add_bullet(
        "Los conectores se configuran a nivel organizacional por el area de TI"
    )
    pdf.add_spacer(0.1)

    pdf.add_subsection("Casos de uso contable")
    pdf.add_bullet(
        "**Base de datos SQL:** Claude consulta directamente el sistema "
        "contable y trae datos a Excel para analisis"
    )
    pdf.add_bullet(
        "**SharePoint/OneDrive:** Acceso a archivos historicos para "
        "comparaciones y consolidaciones"
    )
    pdf.add_bullet(
        "**APIs externas:** Conexion a servicios del SAT, tipo de cambio "
        "del Banco de Mexico, o INPC actualizado"
    )
    pdf.add_bullet(
        "**Correo y calendario:** Consultar fechas limite de declaraciones "
        "y recordatorios de cierre"
    )
    pdf.add_spacer(0.1)

    pdf.add_subsection("Como empezar con MCP")
    pdf.add_bullet("Visitar modelcontextprotocol.io para documentacion completa")
    pdf.add_bullet("Coordinar con el area de TI para configurar conectores organizacionales")
    pdf.add_bullet(
        "Empezar con conectores de solo lectura (consulta) antes de habilitar escritura"
    )
    pdf.add_bullet(
        "Claude Code (terminal) soporta MCP de forma nativa para automatizacion avanzada"
    )

    # ── Seccion final ─────────────────────────────────────────────
    pdf.add_page_break()
    pdf.add_section("Recursos y Enlaces")
    pdf.add_bullet("**Claude para Excel:** Marketplace de Microsoft 365")
    pdf.add_bullet("**Documentacion Claude:** docs.anthropic.com")
    pdf.add_bullet("**Claude Code:** npm install -g @anthropic-ai/claude-code")
    pdf.add_bullet("**MCP Protocol:** modelcontextprotocol.io")
    pdf.add_bullet("**Anthropic:** anthropic.com")
    pdf.add_bullet("**Comunidad del curso:** todoconta.com")
    pdf.add_spacer(0.2)
    pdf.add_text(
        "**Nota final:** La inteligencia artificial es una herramienta que "
        "amplifica la productividad del profesional contable. El criterio humano, "
        "la etica profesional, y el conocimiento normativo siguen siendo "
        "insustituibles. Usen la IA como su segundo cerebro, pero nunca dejen "
        "de ser la brujula."
    )

    pdf.save()


# =====================================================================
# Funcion principal
# =====================================================================

def build():
    """Generate both bonus Markdown guides."""
    _build_vba_guide()
    _build_claude_guide()


if __name__ == "__main__":
    build()
