"""
Generador: Bonus_1_VBA_con_IA.pptx + Script_Bonus_1_VBA.md
Bonus 1 — Automatizacion con Macros VBA Asistidas por IA

Slides:
  1. Portada
  2. Que es VBA?  Cuando macros vs formulas
  3. El editor VBA (Alt+F11)
  4. Prompting IA para generar VBA
  5. Ejemplo 1 - Macro formatear nomina
  6. Ejemplo 2 - Macro limpiar celdas vacias
  7. Ejemplo 3 - Macro reporte mensual con fecha automatica
  8. Ejemplo 4 - Boton actualizar TDs
  9. Claude Code como asistente debugging
  10. Seguridad: .xlsm, Trust Center
  11. Recursos
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import SLIDES_DIR, TELEPROMPTER_DIR
from scripts.generators.pptx_gen import SlidesGenerator


def build():
    gen = SlidesGenerator(
        filename="Bonus_1_VBA_con_IA.md",
        output_dir=SLIDES_DIR,
        script_filename="Script_Bonus_1_VBA.md",
        script_dir=TELEPROMPTER_DIR,
    )

    # ── Slide 1 — Portada ─────────────────────────────────────────
    gen.add_title_slide(
        modulo_num="",
        modulo_nombre="Automatizacion con Macros Asistidas por IA",
        subtitulo="Bonus 1 — VBA + ChatGPT / Claude para contadores",
    )
    gen.script_lines.append(
        "Bienvenidos al primer bonus del curso. En esta sesion vamos a ver "
        "como usar VBA, el lenguaje de macros de Excel, pero con un giro "
        "moderno: vamos a pedirle a la inteligencia artificial que nos genere "
        "el codigo. No necesitan ser programadores; solo necesitan saber "
        "describir lo que quieren automatizar.\n"
    )

    # ── Slide 2 — Que es VBA? ─────────────────────────────────────
    gen.add_content_slide(
        title="Que es VBA? Cuando Macros vs Formulas",
        bullets=[
            "VBA = Visual Basic for Applications (lenguaje de macros de Office)",
            "Las formulas resuelven calculos; las macros automatizan procesos",
            "Ejemplo formula: =BUSCARV(A2, Tabla, 3, 0) -> busca un dato",
            "Ejemplo macro: FormatearNomina() -> aplica formato a 500 filas en 1 clic",
            "Regla practica: si lo haces mas de 3 veces, automatizalo con macro",
            "Las macros ahorran horas en tareas repetitivas de cierre mensual",
        ],
        script_text=(
            "VBA significa Visual Basic for Applications. Es el lenguaje que viene "
            "integrado dentro de Excel desde hace mas de 20 anios. La pregunta clave es: "
            "cuando uso una formula y cuando uso una macro?\n\n"
            "Las formulas son para calculos: sumar, buscar, condicionales. Las macros son "
            "para procesos: dar formato a 500 filas, limpiar datos, generar reportes "
            "automaticamente.\n\n"
            "Mi regla es simple: si haces algo mas de 3 veces al mes y toma mas de 5 minutos "
            "cada vez, merece una macro. Piensen en el cierre mensual: formatear la nomina, "
            "limpiar celdas vacias, actualizar tablas dinamicas... todo eso se puede automatizar.\n"
        ),
    )

    # ── Slide 3 — El editor VBA ───────────────────────────────────
    gen.add_content_slide(
        title="El Editor VBA (Alt + F11)",
        bullets=[
            "Alt + F11  ->  abre el editor de Visual Basic",
            "Panel izquierdo: Explorador de proyecto (VBAProject)",
            "Insertar -> Modulo  ->  aqui se escribe el codigo",
            "F5 para ejecutar la macro seleccionada",
            "F8 para ejecutar linea por linea (debug)",
            "Ventana Inmediato (Ctrl+G): pruebas rapidas",
            "No se asusten por la interfaz; solo necesitamos Modulo + F5",
        ],
        script_text=(
            "Vamos a abrir el editor. Presionen Alt + F11. Se abre una ventana que parece "
            "de los anios 90, y si, es de los anios 90. Pero funciona perfectamente.\n\n"
            "A la izquierda ven el Explorador de Proyecto. Ahi aparece su archivo. "
            "Vayan a Insertar, Modulo. Eso crea un espacio en blanco donde pegamos el codigo.\n\n"
            "Para ejecutar: F5. Para ejecutar linea por linea y ver que pasa: F8. "
            "La ventana Inmediato con Ctrl+G les permite probar instrucciones sueltas.\n\n"
            "Lo importante: no necesitan entender cada linea del codigo. La IA nos lo genera. "
            "Ustedes solo necesitan saber donde pegarlo y como ejecutarlo.\n"
        ),
    )

    # ── Slide 4 — Prompting IA para VBA ───────────────────────────
    gen.add_content_slide(
        title="Prompting IA para Generar VBA",
        bullets=[
            "La IA (ChatGPT, Claude) genera macros VBA si describes bien la tarea",
            "Estructura del prompt:",
            "   1. Contexto: 'Tengo una hoja de nomina en Excel...'",
            "   2. Tarea: 'Necesito una macro que formatee las columnas A:H...'",
            "   3. Detalles: 'Separador de miles, bordes, encabezado azul...'",
            "   4. Restriccion: 'Compatible con Excel 2019 o superior'",
            "Pedir que incluya comentarios en espanol",
            "Siempre probar en copia del archivo, nunca en el original",
        ],
        script_text=(
            "Aqui esta la magia. Ustedes no necesitan aprender a programar en VBA. "
            "Necesitan aprender a PEDIR codigo en VBA. Eso se llama prompting.\n\n"
            "Un buen prompt tiene 4 partes: contexto, tarea, detalles y restriccion. "
            "Por ejemplo:\n\n"
            "PROMPT: 'Genera una macro VBA para Excel. Tengo una hoja llamada Nomina con "
            "datos en A1:H50. La macro debe: aplicar formato de miles con 2 decimales a "
            "las columnas E, F, G y H; poner bordes delgados a todo el rango; pintar la "
            "fila 1 de azul con texto blanco. Incluye comentarios en espanol. Compatible "
            "con Excel 2019.'\n\n"
            "Eso es todo. La IA les regresa el codigo listo para copiar y pegar. "
            "Siempre prueben en una copia del archivo, nunca en el original.\n"
        ),
    )

    # ── Slide 5 — Ejemplo 1: FormatearNomina ─────────────────────
    gen.add_content_slide(
        title="Ejemplo 1: Macro FormatearNomina()",
        bullets=[
            "Objetivo: formato profesional en 1 clic",
            "Aplica separador de miles a columnas monetarias",
            "Bordes delgados en todo el rango de datos",
            "Encabezado: fondo azul, texto blanco, negrita",
            "Autoajuste de ancho de columnas",
            "Resultado: de datos crudos a tabla presentable en 2 segundos",
        ],
        script_text=(
            "Nuestro primer ejemplo practico. Vamos a crear la macro FormatearNomina.\n\n"
            "El codigo que vamos a pegar es:\n\n"
            "```vba\n"
            "Sub FormatearNomina()\n"
            "    Dim ws As Worksheet\n"
            "    Dim rng As Range\n"
            "    Dim lastRow As Long, lastCol As Long\n"
            "    \n"
            "    Set ws = ActiveSheet\n"
            "    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row\n"
            "    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column\n"
            "    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))\n"
            "    \n"
            "    ' Formato numerico a columnas monetarias (E:H)\n"
            "    ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 8)).NumberFormat = \"#,##0.00\"\n"
            "    \n"
            "    ' Bordes delgados\n"
            "    rng.Borders.LineStyle = xlContinuous\n"
            "    rng.Borders.Weight = xlThin\n"
            "    \n"
            "    ' Encabezado azul\n"
            "    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))\n"
            "        .Interior.Color = RGB(37, 99, 235)\n"
            "        .Font.Color = RGB(255, 255, 255)\n"
            "        .Font.Bold = True\n"
            "    End With\n"
            "    \n"
            "    ' Autoajuste\n"
            "    rng.Columns.AutoFit\n"
            "    \n"
            "    MsgBox \"Nomina formateada correctamente.\", vbInformation\n"
            "End Sub\n"
            "```\n\n"
            "Copiamos este codigo, abrimos Alt+F11, Insertar Modulo, pegamos, y F5. "
            "En 2 segundos la nomina queda presentable.\n"
        ),
    )

    # ── Slide 6 — Ejemplo 2: LimpiarVacias ───────────────────────
    gen.add_content_slide(
        title="Ejemplo 2: Macro LimpiarVacias()",
        bullets=[
            "Problema comun: datos exportados con filas vacias intercaladas",
            "Solucion manual: seleccionar, eliminar, repetir... (tedioso)",
            "Macro: recorre de abajo hacia arriba eliminando filas vacias",
            "Importante: recorrer de abajo hacia arriba para no saltar filas",
            "Se aplica sobre la seleccion activa del usuario",
            "Resultado: datos limpios en milisegundos",
        ],
        script_text=(
            "Segundo ejemplo. Cuantas veces les ha pasado que exportan datos del SAT "
            "o del sistema contable y vienen filas vacias intercaladas?\n\n"
            "El codigo:\n\n"
            "```vba\n"
            "Sub LimpiarVacias()\n"
            "    Dim rng As Range\n"
            "    Dim i As Long\n"
            "    \n"
            "    Set rng = Selection\n"
            "    \n"
            "    ' Recorrer de abajo hacia arriba\n"
            "    For i = rng.Rows.Count To 1 Step -1\n"
            "        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then\n"
            "            rng.Rows(i).EntireRow.Delete\n"
            "        End If\n"
            "    Next i\n"
            "    \n"
            "    MsgBox \"Filas vacias eliminadas.\", vbInformation\n"
            "End Sub\n"
            "```\n\n"
            "El truco clave es recorrer de abajo hacia arriba. Si van de arriba hacia abajo "
            "y eliminan una fila, las demas se mueven y se saltan una. Clasico error.\n\n"
            "Seleccionen el rango con datos, ejecuten la macro, y listo.\n"
        ),
    )

    # ── Slide 7 — Ejemplo 3: ReporteMensual ──────────────────────
    gen.add_content_slide(
        title="Ejemplo 3: Macro ReporteMensual()",
        bullets=[
            "Crea nueva hoja con nombre automatico: 'Reporte_Feb_2026'",
            "Copia estructura base desde hoja plantilla",
            "Inserta fecha de generacion automaticamente",
            "Ideal para reportes mensuales de cierre",
            "Prompt IA: 'Macro que cree hoja nueva con nombre mes-anio actual'",
        ],
        script_text=(
            "Tercer ejemplo. Cada mes creamos un reporte de cierre. En vez de copiar "
            "la hoja manualmente y renombrarla, la macro lo hace sola.\n\n"
            "```vba\n"
            "Sub ReporteMensual()\n"
            "    Dim ws As Worksheet\n"
            "    Dim nombreHoja As String\n"
            "    Dim meses As Variant\n"
            "    \n"
            "    meses = Array(\"Ene\", \"Feb\", \"Mar\", \"Abr\", \"May\", \"Jun\", _\n"
            "                   \"Jul\", \"Ago\", \"Sep\", \"Oct\", \"Nov\", \"Dic\")\n"
            "    \n"
            "    nombreHoja = \"Reporte_\" & meses(Month(Date) - 1) & \"_\" & Year(Date)\n"
            "    \n"
            "    ' Verificar si ya existe\n"
            "    On Error Resume Next\n"
            "    Set ws = ThisWorkbook.Sheets(nombreHoja)\n"
            "    On Error GoTo 0\n"
            "    \n"
            "    If Not ws Is Nothing Then\n"
            "        MsgBox \"La hoja \" & nombreHoja & \" ya existe.\", vbExclamation\n"
            "        Exit Sub\n"
            "    End If\n"
            "    \n"
            "    ' Crear nueva hoja\n"
            "    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))\n"
            "    ws.Name = nombreHoja\n"
            "    \n"
            "    ' Encabezado\n"
            "    ws.Range(\"A1\").Value = \"Reporte Mensual - \" & Format(Date, \"MMMM YYYY\")\n"
            "    ws.Range(\"A2\").Value = \"Generado: \" & Format(Now, \"DD/MM/YYYY HH:MM\")\n"
            "    ws.Range(\"A1\").Font.Bold = True\n"
            "    ws.Range(\"A1\").Font.Size = 14\n"
            "    \n"
            "    MsgBox \"Hoja '\" & nombreHoja & \"' creada.\", vbInformation\n"
            "End Sub\n"
            "```\n\n"
            "Fijense que la macro verifica si la hoja ya existe para no dar error. "
            "Ese tipo de detalle la IA lo incluye si se lo piden en el prompt.\n"
        ),
    )

    # ── Slide 8 — Ejemplo 4: Actualizar TDs ──────────────────────
    gen.add_content_slide(
        title="Ejemplo 4: Boton Actualizar Tablas Dinamicas",
        bullets=[
            "Problema: 5 tablas dinamicas en el libro, hay que actualizar una por una",
            "Macro: actualiza TODAS las tablas dinamicas del libro",
            "Se puede asignar a un boton con forma en la hoja",
            "Insertar -> Formas -> Rectangulo -> clic derecho -> Asignar macro",
            "Resultado: un boton 'Actualizar Todo' que ahorra 2 minutos cada vez",
        ],
        script_text=(
            "Cuarto ejemplo. Cuando tienen un dashboard con varias tablas dinamicas, "
            "actualizar cada una es tedioso. Esta macro actualiza todas de un golpe.\n\n"
            "```vba\n"
            "Sub ActualizarPivots()\n"
            "    Dim ws As Worksheet\n"
            "    Dim pt As PivotTable\n"
            "    Dim contador As Long\n"
            "    \n"
            "    contador = 0\n"
            "    \n"
            "    For Each ws In ThisWorkbook.Worksheets\n"
            "        For Each pt In ws.PivotTables\n"
            "            pt.RefreshTable\n"
            "            contador = contador + 1\n"
            "        Next pt\n"
            "    Next ws\n"
            "    \n"
            "    MsgBox contador & \" tabla(s) dinamica(s) actualizada(s).\", vbInformation\n"
            "End Sub\n"
            "```\n\n"
            "Ahora, para hacerlo mas profesional, vamos a asignar esta macro a un boton. "
            "Van a Insertar, Formas, eligen un rectangulo redondeado, lo dibujan en la hoja, "
            "le escriben 'Actualizar Todo', clic derecho sobre la forma, Asignar macro, "
            "seleccionan ActualizarPivots, y listo. Un boton profesional en su dashboard.\n"
        ),
    )

    # ── Slide 9 — Claude Code como asistente ──────────────────────
    gen.add_content_slide(
        title="Claude Code como Asistente de Debugging",
        bullets=[
            "Error comun: 'Run-time error 1004' -> referencia invalida",
            "Copiar el error + el codigo y pegarlo en Claude/ChatGPT",
            "Claude Code (terminal): analiza archivos .xlsm directamente",
            "Prompt: 'Este codigo VBA da error 1004 en la linea X. Explicame por que y corrigelo'",
            "La IA explica el error Y da la correccion",
            "Ciclo: Generar -> Probar -> Si falla, pedir correccion -> Probar otra vez",
        ],
        script_text=(
            "Que pasa cuando el codigo da error? Porque va a pasar, especialmente al principio.\n\n"
            "El error mas comun es 'Run-time error 1004'. Significa que la macro esta intentando "
            "acceder a un rango que no existe o esta mal referenciado.\n\n"
            "Lo que hacen es: copian el mensaje de error y el codigo, lo pegan en Claude o ChatGPT, "
            "y le dicen: 'Este codigo VBA da error 1004 en la linea 15. Explicame por que y corrigelo.'\n\n"
            "La IA les va a explicar exactamente cual es el problema y les da el codigo corregido. "
            "Es un ciclo: generar, probar, si falla, pedir correccion, probar otra vez.\n\n"
            "Si usan Claude Code desde la terminal, pueden incluso apuntar al archivo .xlsm "
            "y Claude lo analiza directamente. Es como tener un programador a su lado.\n"
        ),
    )

    # ── Slide 10 — Seguridad ──────────────────────────────────────
    gen.add_content_slide(
        title="Seguridad: .xlsm y Trust Center",
        bullets=[
            ".xlsx NO soporta macros -> .xlsm SI soporta macros",
            "Al guardar: Guardar como -> Libro de Excel habilitado para macros (.xlsm)",
            "Trust Center: Archivo -> Opciones -> Centro de confianza",
            "Recomendacion: 'Deshabilitar macros con notificacion'",
            "Nunca habilitar macros de archivos desconocidos",
            "Firmar macros digitalmente en ambientes corporativos",
            "Respaldo: siempre mantener copia .xlsx sin macros",
        ],
        script_text=(
            "Un punto importante de seguridad. Los archivos .xlsx normales NO pueden contener "
            "macros. Para guardar macros necesitan el formato .xlsm.\n\n"
            "Vayan a Guardar como, y en Tipo cambien a 'Libro de Excel habilitado para macros'. "
            "Si guardan como .xlsx, las macros se pierden.\n\n"
            "En cuanto al Trust Center: vayan a Archivo, Opciones, Centro de confianza, "
            "Configuracion del Centro de confianza, Configuracion de macros. La opcion recomendada "
            "es 'Deshabilitar macros con notificacion'. Asi Excel les pregunta cada vez que abren "
            "un archivo con macros.\n\n"
            "NUNCA habiliten macros de archivos que reciben por correo de fuentes desconocidas. "
            "Las macros pueden ser maliciosas. Solo confien en macros que ustedes mismos crearon "
            "o que provienen de fuentes verificadas.\n"
        ),
    )

    # ── Slide 11 — Recursos y Cierre ──────────────────────────────
    gen.add_closing_slide(
        next_module="Bonus 2 - Claude en Excel: Tu Segundo Cerebro Contable",
        resources=[
            "Editor VBA: Alt + F11 en cualquier Excel",
            "Documentacion VBA: docs.microsoft.com/en-us/office/vba/api/overview/excel",
            "ChatGPT / Claude para generar y depurar macros",
            "Practica: crear las 4 macros de esta sesion en un archivo .xlsm",
            "Tip: guardar macros utiles en un Personal.xlsb para tenerlas siempre disponibles",
        ],
    )
    gen.script_lines.append(
        "Y eso es todo para este bonus. Repasamos que es VBA, vimos como usar la IA "
        "para generar macros sin programar, creamos 4 macros practicas para contadores, "
        "y hablamos de seguridad.\n\n"
        "La clave es: ustedes describen, la IA programa, ustedes verifican. "
        "No necesitan ser ingenieros. Solo necesitan saber que quieren automatizar.\n\n"
        "En el siguiente bonus veremos como usar Claude directamente dentro de Excel "
        "como un asistente de inteligencia artificial integrado. Nos vemos ahi.\n"
    )

    gen.save()
    print("Bonus 1 - VBA con IA generado correctamente.")


if __name__ == "__main__":
    build()
