"""
Generador de PDFs Bonus:
  1. Guia_VBA_con_IA.pdf — Plantillas de prompts y 5 macros listas para copiar
  2. Guia_Claude_en_Excel.pdf — Guia de instalacion, prompts y comparativa

Salida: output/Pack_Excel_Pro/Bonus/
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.colors import HexColor

from scripts.config.constants import PACK, Color, CURSO_NOMBRE, INSTRUCTOR, ANIO
from scripts.generators.pdf_gen import (
    PDFGenerator, RL_AZUL, RL_BLANCO, RL_TEXTO, RL_FONDO, RL_GRIS_BORDE,
    RL_TEXTO_MEDIO, RL_VERDE, RL_ROJO,
)

OUTPUT_DIR = PACK / "Bonus"


# =====================================================================
# PDF 1 — Guia VBA con IA
# =====================================================================

def _build_vba_guide():
    pdf = PDFGenerator(
        filename="Guia_VBA_con_IA.pdf",
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
        "<b>Requisitos:</b> Excel 2019, 2021, o Microsoft 365 en Windows. "
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
    vba1 = (
        "Sub FormatearNomina()<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim ws As Worksheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim rng As Range<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim lastRow As Long, lastCol As Long<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set ws = ActiveSheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Formato numerico a columnas monetarias (E:H)<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 8)).NumberFormat = \"#,##0.00\"<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Bordes delgados<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;rng.Borders.LineStyle = xlContinuous<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;rng.Borders.Weight = xlThin<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Encabezado azul<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.Interior.Color = RGB(37, 99, 235)<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.Font.Color = RGB(255, 255, 255)<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.Font.Bold = True<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;End With<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Autoajuste<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;rng.Columns.AutoFit<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"Nomina formateada correctamente.\", vbInformation<br/>"
        "End Sub"
    )
    pdf.add_code(vba1)
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
    vba2 = (
        "Sub LimpiarVacias()<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim rng As Range<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim i As Long<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set rng = Selection<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Recorrer de abajo hacia arriba<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;For i = rng.Rows.Count To 1 Step -1<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;rng.Rows(i).EntireRow.Delete<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Next i<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"Filas vacias eliminadas.\", vbInformation<br/>"
        "End Sub"
    )
    pdf.add_code(vba2)
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
    vba3 = (
        "Sub ReporteMensual()<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim ws As Worksheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim nombreHoja As String<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim meses As Variant<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;meses = Array(\"Ene\", \"Feb\", \"Mar\", \"Abr\", \"May\", \"Jun\", _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        "\"Jul\", \"Ago\", \"Sep\", \"Oct\", \"Nov\", \"Dic\")<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;nombreHoja = \"Reporte_\" &amp; meses(Month(Date) - 1) &amp; \"_\" &amp; Year(Date)<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Verificar si ya existe<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;On Error Resume Next<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set ws = ThisWorkbook.Sheets(nombreHoja)<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;On Error GoTo 0<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;If Not ws Is Nothing Then<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"La hoja \" &amp; nombreHoja &amp; \" ya existe.\", vbExclamation<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Exit Sub<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;End If<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Crear nueva hoja<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set ws = ThisWorkbook.Sheets.Add( _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Name = nombreHoja<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;' Encabezado<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Range(\"A1\").Value = \"Reporte Mensual - \" &amp; Format(Date, \"MMMM YYYY\")<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Range(\"A2\").Value = \"Generado: \" &amp; Format(Now, \"DD/MM/YYYY HH:MM\")<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Range(\"A1\").Font.Bold = True<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.Range(\"A1\").Font.Size = 14<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"Hoja '\" &amp; nombreHoja &amp; \"' creada.\", vbInformation<br/>"
        "End Sub"
    )
    pdf.add_code(vba3)
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
    vba4 = (
        "Sub ActualizarPivots()<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim ws As Worksheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim pt As PivotTable<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim contador As Long<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;contador = 0<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;For Each ws In ThisWorkbook.Worksheets<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For Each pt In ws.PivotTables<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;pt.RefreshTable<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;contador = contador + 1<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Next pt<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Next ws<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;MsgBox contador &amp; \" tabla(s) dinamica(s) actualizada(s).\", vbInformation<br/>"
        "End Sub"
    )
    pdf.add_code(vba4)
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
    vba5 = (
        "Sub ExportarPDF()<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim ws As Worksheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim rutaPDF As String<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Dim rutaLibro As String<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;Set ws = ActiveSheet<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;rutaLibro = ThisWorkbook.Path<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;If rutaLibro = \"\" Then<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"Guarde el libro primero.\", vbExclamation<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Exit Sub<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;End If<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;rutaPDF = rutaLibro &amp; \"\\\" &amp; ws.Name &amp; \"_\" &amp; _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        "Format(Date, \"YYYY-MM-DD\") &amp; \".pdf\"<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;ws.ExportAsFixedFormat _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type:=xlTypePDF, _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Filename:=rutaPDF, _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Quality:=xlQualityStandard, _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IncludeDocProperties:=True, _<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;OpenAfterPublish:=True<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;MsgBox \"PDF exportado: \" &amp; rutaPDF, vbInformation<br/>"
        "End Sub"
    )
    pdf.add_code(vba5)
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
        "<b>Personal.xlsb:</b> Guarda tus macros mas utiles en el libro Personal "
        "(Archivo > Opciones > Guardar macro en > Libro de macros personal). "
        "Asi estan disponibles en todos tus archivos."
    )
    pdf.add_bullet(
        "<b>Seguridad:</b> Solo habilita macros de fuentes confiables. "
        "Configura el Trust Center en: Archivo > Opciones > Centro de confianza."
    )
    pdf.add_bullet(
        "<b>Respaldo:</b> Siempre prueba macros en una copia del archivo, nunca en el original."
    )
    pdf.add_bullet(
        "<b>Errores:</b> Si una macro da error, copia el mensaje de error y el codigo, "
        "pegalo en Claude o ChatGPT, y pide explicacion y correccion."
    )
    pdf.add_bullet(
        "<b>Documentacion:</b> Agrega comentarios (lineas con ') a tus macros para "
        "que tu yo futuro entienda que hacen."
    )

    pdf.save()
    print("PDF 1 - Guia VBA con IA generado correctamente.")


# =====================================================================
# PDF 2 — Guia Claude en Excel
# =====================================================================

def _build_claude_guide():
    pdf = PDFGenerator(
        filename="Guia_Claude_en_Excel.pdf",
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
    pdf.add_bullet("<b>Paso 1:</b> Abrir Excel y crear o abrir cualquier libro")
    pdf.add_bullet("<b>Paso 2:</b> Ir a la pestana Insertar en la cinta de opciones")
    pdf.add_bullet("<b>Paso 3:</b> Clic en 'Obtener complementos' (o 'Get Add-ins')")
    pdf.add_bullet("<b>Paso 4:</b> En el cuadro de busqueda, escribir 'Claude for Excel'")
    pdf.add_bullet("<b>Paso 5:</b> Seleccionar el add-in oficial de Anthropic y clic en 'Agregar'")
    pdf.add_bullet("<b>Paso 6:</b> Aceptar los permisos solicitados")
    pdf.add_bullet("<b>Paso 7:</b> El panel de Claude aparecera en la pestana Inicio")
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
            "desvien mas del 30%% del promedio historico de su categoria. Senala "
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

    comparison_header = [
        Paragraph("<b>Caracteristica</b>", pdf.styles["BodyText2"]),
        Paragraph("<b>Claude (Anthropic)</b>", pdf.styles["BodyText2"]),
        Paragraph("<b>Copilot (Microsoft)</b>", pdf.styles["BodyText2"]),
    ]
    comparison_data = [comparison_header]

    rows = [
        ("Integracion en Excel", "Add-in desde Marketplace", "Nativo en M365"),
        ("Analisis profundo de datos", "Excelente", "Bueno"),
        ("Generacion de formulas", "Excelente (con explicacion)", "Muy bueno"),
        ("Automatizacion rapida", "Buena (via prompts)", "Excelente (nativo)"),
        ("Generacion de graficos", "Sugiere configuracion", "Crea directamente"),
        ("Razonamiento complejo", "Superior", "Bueno"),
        ("Codigo VBA", "Genera y explica", "Genera"),
        ("Lenguaje natural en espaniol", "Excelente", "Muy bueno"),
        ("Privacidad de datos", "No almacena datos", "Segun plan M365"),
        ("Costo", "Incluido en M365 Pro", "Incluido en M365 Pro/Copilot"),
        ("Mejor para", "Analisis, razonamiento, auditoria", "Tareas rapidas, automatizacion"),
        ("Conectores externos (MCP)", "Si, via protocolo MCP", "Si, via Microsoft Graph"),
    ]

    for caract, claude_val, copilot_val in rows:
        comparison_data.append([
            Paragraph(caract, pdf.styles["BodyText2"]),
            Paragraph(claude_val, pdf.styles["BodyText2"]),
            Paragraph(copilot_val, pdf.styles["BodyText2"]),
        ])

    col_widths = [2.2 * inch, 2.4 * inch, 2.4 * inch]
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (-1, -1), RL_TEXTO),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, RL_GRIS_BORDE),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("BACKGROUND", (0, 0), (-1, 0), RL_AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), RL_BLANCO),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]
    for i in range(2, len(comparison_data), 2):
        style_cmds.append(("BACKGROUND", (0, i), (-1, i), RL_FONDO))

    tbl = Table(comparison_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    pdf.elements.append(tbl)
    pdf.add_spacer(0.2)

    pdf.add_text(
        "<b>Recomendacion:</b> Usa Copilot para tareas rapidas de automatizacion "
        "y creacion de graficos. Usa Claude para analisis profundo, razonamiento "
        "sobre datos, auditoria de formulas, y generacion de codigo VBA con "
        "explicaciones detalladas. Juntos cubren el 95%% de las necesidades."
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
        "Claude <b>no almacena</b> los datos de las hojas que procesa en las "
        "sesiones del add-in"
    )
    pdf.add_bullet(
        "Los datos se envian a los servidores de Anthropic para procesamiento "
        "y se descartan despues de generar la respuesta"
    )
    pdf.add_bullet(
        "Anthropic <b>no usa datos de clientes empresariales</b> para entrenar "
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
        "<b>Base de datos SQL:</b> Claude consulta directamente el sistema "
        "contable y trae datos a Excel para analisis"
    )
    pdf.add_bullet(
        "<b>SharePoint/OneDrive:</b> Acceso a archivos historicos para "
        "comparaciones y consolidaciones"
    )
    pdf.add_bullet(
        "<b>APIs externas:</b> Conexion a servicios del SAT, tipo de cambio "
        "del Banco de Mexico, o INPC actualizado"
    )
    pdf.add_bullet(
        "<b>Correo y calendario:</b> Consultar fechas limite de declaraciones "
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
    pdf.add_bullet("<b>Claude para Excel:</b> Marketplace de Microsoft 365")
    pdf.add_bullet("<b>Documentacion Claude:</b> docs.anthropic.com")
    pdf.add_bullet("<b>Claude Code:</b> npm install -g @anthropic-ai/claude-code")
    pdf.add_bullet("<b>MCP Protocol:</b> modelcontextprotocol.io")
    pdf.add_bullet("<b>Anthropic:</b> anthropic.com")
    pdf.add_bullet("<b>Comunidad del curso:</b> todoconta.com")
    pdf.add_spacer(0.2)
    pdf.add_text(
        "<b>Nota final:</b> La inteligencia artificial es una herramienta que "
        "amplifica la productividad del profesional contable. El criterio humano, "
        "la etica profesional, y el conocimiento normativo siguen siendo "
        "insustituibles. Usen la IA como su segundo cerebro, pero nunca dejen "
        "de ser la brujula."
    )

    pdf.save()
    print("PDF 2 - Guia Claude en Excel generado correctamente.")


# =====================================================================
# Funcion principal
# =====================================================================

def build():
    """Generate both bonus PDFs."""
    _build_vba_guide()
    _build_claude_guide()


if __name__ == "__main__":
    build()
