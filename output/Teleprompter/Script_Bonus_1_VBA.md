# Módulo : Automatizacion con Macros Asistidas por IA

## Slide 1 — Portada

Bienvenidos al primer bonus del curso. En esta sesion vamos a ver como usar VBA, el lenguaje de macros de Excel, pero con un giro moderno: vamos a pedirle a la inteligencia artificial que nos genere el codigo. No necesitan ser programadores; solo necesitan saber describir lo que quieren automatizar.

## Slide 2 — Que es VBA? Cuando Macros vs Formulas

VBA significa Visual Basic for Applications. Es el lenguaje que viene integrado dentro de Excel desde hace mas de 20 anios. La pregunta clave es: cuando uso una formula y cuando uso una macro?

Las formulas son para calculos: sumar, buscar, condicionales. Las macros son para procesos: dar formato a 500 filas, limpiar datos, generar reportes automaticamente.

Mi regla es simple: si haces algo mas de 3 veces al mes y toma mas de 5 minutos cada vez, merece una macro. Piensen en el cierre mensual: formatear la nomina, limpiar celdas vacias, actualizar tablas dinamicas... todo eso se puede automatizar.


## Slide 3 — El Editor VBA (Alt + F11)

Vamos a abrir el editor. Presionen Alt + F11. Se abre una ventana que parece de los anios 90, y si, es de los anios 90. Pero funciona perfectamente.

A la izquierda ven el Explorador de Proyecto. Ahi aparece su archivo. Vayan a Insertar, Modulo. Eso crea un espacio en blanco donde pegamos el codigo.

Para ejecutar: F5. Para ejecutar linea por linea y ver que pasa: F8. La ventana Inmediato con Ctrl+G les permite probar instrucciones sueltas.

Lo importante: no necesitan entender cada linea del codigo. La IA nos lo genera. Ustedes solo necesitan saber donde pegarlo y como ejecutarlo.


## Slide 4 — Prompting IA para Generar VBA

Aqui esta la magia. Ustedes no necesitan aprender a programar en VBA. Necesitan aprender a PEDIR codigo en VBA. Eso se llama prompting.

Un buen prompt tiene 4 partes: contexto, tarea, detalles y restriccion. Por ejemplo:

PROMPT: 'Genera una macro VBA para Excel. Tengo una hoja llamada Nomina con datos en A1:H50. La macro debe: aplicar formato de miles con 2 decimales a las columnas E, F, G y H; poner bordes delgados a todo el rango; pintar la fila 1 de azul con texto blanco. Incluye comentarios en espaniol. Compatible con Excel 2019.'

Eso es todo. La IA les regresa el codigo listo para copiar y pegar. Siempre prueben en una copia del archivo, nunca en el original.


## Slide 5 — Ejemplo 1: Macro FormatearNomina()

Nuestro primer ejemplo practico. Vamos a crear la macro FormatearNomina.

El codigo que vamos a pegar es:

```vba
Sub FormatearNomina()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Formato numerico a columnas monetarias (E:H)
    ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 8)).NumberFormat = "#,##0.00"
    
    ' Bordes delgados
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    
    ' Encabezado azul
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(37, 99, 235)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Autoajuste
    rng.Columns.AutoFit
    
    MsgBox "Nomina formateada correctamente.", vbInformation
End Sub
```

Copiamos este codigo, abrimos Alt+F11, Insertar Modulo, pegamos, y F5. En 2 segundos la nomina queda presentable.


## Slide 6 — Ejemplo 2: Macro LimpiarVacias()

Segundo ejemplo. Cuantas veces les ha pasado que exportan datos del SAT o del sistema contable y vienen filas vacias intercaladas?

El codigo:

```vba
Sub LimpiarVacias()
    Dim rng As Range
    Dim i As Long
    
    Set rng = Selection
    
    ' Recorrer de abajo hacia arriba
    For i = rng.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
            rng.Rows(i).EntireRow.Delete
        End If
    Next i
    
    MsgBox "Filas vacias eliminadas.", vbInformation
End Sub
```

El truco clave es recorrer de abajo hacia arriba. Si van de arriba hacia abajo y eliminan una fila, las demas se mueven y se saltan una. Clasico error.

Seleccionen el rango con datos, ejecuten la macro, y listo.


## Slide 7 — Ejemplo 3: Macro ReporteMensual()

Tercer ejemplo. Cada mes creamos un reporte de cierre. En vez de copiar la hoja manualmente y renombrarla, la macro lo hace sola.

```vba
Sub ReporteMensual()
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim meses As Variant
    
    meses = Array("Ene", "Feb", "Mar", "Abr", "May", "Jun", _
                   "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
    
    nombreHoja = "Reporte_" & meses(Month(Date) - 1) & "_" & Year(Date)
    
    ' Verificar si ya existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        MsgBox "La hoja " & nombreHoja & " ya existe.", vbExclamation
        Exit Sub
    End If
    
    ' Crear nueva hoja
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nombreHoja
    
    ' Encabezado
    ws.Range("A1").Value = "Reporte Mensual - " & Format(Date, "MMMM YYYY")
    ws.Range("A2").Value = "Generado: " & Format(Now, "DD/MM/YYYY HH:MM")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    MsgBox "Hoja '" & nombreHoja & "' creada.", vbInformation
End Sub
```

Fijense que la macro verifica si la hoja ya existe para no dar error. Ese tipo de detalle la IA lo incluye si se lo piden en el prompt.


## Slide 8 — Ejemplo 4: Boton Actualizar Tablas Dinamicas

Cuarto ejemplo. Cuando tienen un dashboard con varias tablas dinamicas, actualizar cada una es tedioso. Esta macro actualiza todas de un golpe.

```vba
Sub ActualizarPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim contador As Long
    
    contador = 0
    
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
            contador = contador + 1
        Next pt
    Next ws
    
    MsgBox contador & " tabla(s) dinamica(s) actualizada(s).", vbInformation
End Sub
```

Ahora, para hacerlo mas profesional, vamos a asignar esta macro a un boton. Van a Insertar, Formas, eligen un rectangulo redondeado, lo dibujan en la hoja, le escriben 'Actualizar Todo', clic derecho sobre la forma, Asignar macro, seleccionan ActualizarPivots, y listo. Un boton profesional en su dashboard.


## Slide 9 — Claude Code como Asistente de Debugging

Que pasa cuando el codigo da error? Porque va a pasar, especialmente al principio.

El error mas comun es 'Run-time error 1004'. Significa que la macro esta intentando acceder a un rango que no existe o esta mal referenciado.

Lo que hacen es: copian el mensaje de error y el codigo, lo pegan en Claude o ChatGPT, y le dicen: 'Este codigo VBA da error 1004 en la linea 15. Explicame por que y corrigelo.'

La IA les va a explicar exactamente cual es el problema y les da el codigo corregido. Es un ciclo: generar, probar, si falla, pedir correccion, probar otra vez.

Si usan Claude Code desde la terminal, pueden incluso apuntar al archivo .xlsm y Claude lo analiza directamente. Es como tener un programador a su lado.


## Slide 10 — Seguridad: .xlsm y Trust Center

Un punto importante de seguridad. Los archivos .xlsx normales NO pueden contener macros. Para guardar macros necesitan el formato .xlsm.

Vayan a Guardar como, y en Tipo cambien a 'Libro de Excel habilitado para macros'. Si guardan como .xlsx, las macros se pierden.

En cuanto al Trust Center: vayan a Archivo, Opciones, Centro de confianza, Configuracion del Centro de confianza, Configuracion de macros. La opcion recomendada es 'Deshabilitar macros con notificacion'. Asi Excel les pregunta cada vez que abren un archivo con macros.

NUNCA habiliten macros de archivos que reciben por correo de fuentes desconocidas. Las macros pueden ser maliciosas. Solo confien en macros que ustedes mismos crearon o que provienen de fuentes verificadas.


## Slide 11 — Cierre

Y eso es todo para este bonus. Repasamos que es VBA, vimos como usar la IA para generar macros sin programar, creamos 4 macros practicas para contadores, y hablamos de seguridad.

La clave es: ustedes describen, la IA programa, ustedes verifican. No necesitan ser ingenieros. Solo necesitan saber que quieren automatizar.

En el siguiente bonus veremos como usar Claude directamente dentro de Excel como un asistente de inteligencia artificial integrado. Nos vemos ahi.
