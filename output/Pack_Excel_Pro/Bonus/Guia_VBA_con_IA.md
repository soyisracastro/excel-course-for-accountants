# Guia VBA con IA

**Bonus - Material Complementario**

*5 macros listas para copiar + plantillas de prompts*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Introduccion

Esta guia contiene 5 macros VBA listas para copiar y pegar en Excel. Cada macro incluye el codigo completo, instrucciones de uso, y el prompt exacto que puedes usar en ChatGPT o Claude para generar o modificar la macro.


**Requisitos:** Excel 2019, 2021, o Microsoft 365 en Windows. Guardar el archivo como .xlsm (habilitado para macros). Abrir el editor VBA con Alt + F11, insertar un Modulo, y pegar el codigo.

## Plantillas de Prompts para Generar VBA

Usa estas plantillas como punto de partida para pedirle a la IA que genere macros personalizadas.


### Prompt basico

```
Genera una macro VBA para Excel que [TAREA]. Los datos estan en la hoja [NOMBRE] en el rango [RANGO]. Incluye comentarios en espaniol. Compatible con Excel [VERSION].
```


### Prompt con formato

```
Crea una macro VBA que formatee la hoja activa: separador de miles en columnas [X:Y], bordes delgados en todo el rango, encabezado con fondo azul y texto blanco. Autoajustar ancho de columnas.
```


### Prompt de limpieza

```
Genera una macro VBA que limpie datos: elimine filas vacias en la seleccion actual, quite espacios extra con Trim, y convierta texto a mayusculas en la columna [X].
```


### Prompt de reporte

```
Crea una macro VBA que genere un reporte mensual: nueva hoja con nombre 'Reporte_[Mes]_[Anio]', copie la estructura de la hoja Plantilla, e inserte la fecha de generacion en A1.
```


### Prompt de debugging

```
Este codigo VBA da el error '[ERROR]' en la linea [N]. Explicame que causa el error y dame el codigo corregido. El codigo es: [PEGAR CODIGO]
```


---

## Macro 1: FormatearNomina()

### Proposito

Aplica formato profesional a una tabla de nomina: separador de miles con 2 decimales en columnas monetarias (E:H), bordes delgados en todo el rango, encabezado con fondo azul y texto blanco, y autoajuste de columnas.

### Codigo VBA

```
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

### Como usar

- Abrir el archivo de nomina en Excel
- Alt + F11 -> Insertar -> Modulo -> Pegar el codigo
- Posicionar el cursor dentro de Sub FormatearNomina()
- Presionar F5 para ejecutar
### Prompt para generarla

```
Genera una macro VBA llamada FormatearNomina para Excel. Los datos estan en la hoja activa. La macro debe: aplicar formato #,##0.00 a las columnas E:H desde fila 2, bordes delgados a todo el rango con datos, pintar la fila 1 de azul RGB(37,99,235) con texto blanco y negrita, y autoajustar columnas. Mostrar MsgBox al terminar. Comentarios en espaniol.
```

---

## Macro 2: LimpiarVacias()

### Proposito

Elimina filas completamente vacias dentro de la seleccion actual del usuario. Recorre de abajo hacia arriba para evitar saltar filas al eliminar.

### Codigo VBA

```
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

### Como usar

- Seleccionar el rango de datos con posibles filas vacias
- Ejecutar la macro con F5 desde el editor VBA
- Verificar que las filas vacias fueron eliminadas
### Prompt para generarla

```
Genera una macro VBA llamada LimpiarVacias. Debe recorrer la seleccion actual del usuario de abajo hacia arriba. Si una fila esta completamente vacia (CountA = 0), eliminarla con EntireRow.Delete. Mostrar MsgBox al terminar. Comentarios en espaniol.
```

---

## Macro 3: ReporteMensual()

### Proposito

Crea una nueva hoja con nombre automatico basado en el mes y anio actual (ej: Reporte_Feb_2026). Verifica que no exista previamente. Inserta encabezado con fecha de generacion.

### Codigo VBA

```
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
    Set ws = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nombreHoja
    
    ' Encabezado
    ws.Range("A1").Value = "Reporte Mensual - " & Format(Date, "MMMM YYYY")
    ws.Range("A2").Value = "Generado: " & Format(Now, "DD/MM/YYYY HH:MM")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    MsgBox "Hoja '" & nombreHoja & "' creada.", vbInformation
End Sub
```

### Como usar

- Abrir el libro donde se necesita el reporte mensual
- Ejecutar la macro; crea la hoja automaticamente
- Si la hoja ya existe, muestra advertencia y no la sobrescribe
### Prompt para generarla

```
Genera una macro VBA llamada ReporteMensual. Debe crear una hoja nueva con nombre Reporte_[Mes]_[Anio] usando la fecha actual. Verificar que la hoja no exista antes de crearla. En A1 poner titulo con formato MMMM YYYY, en A2 poner fecha y hora de generacion. A1 en negrita tamano 14. MsgBox al terminar. Comentarios en espaniol.
```

---

## Macro 4: ActualizarPivots()

### Proposito

Recorre todas las hojas del libro y actualiza cada tabla dinamica encontrada. Muestra un conteo de cuantas tablas fueron actualizadas.

### Codigo VBA

```
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

### Como usar

- Tener un libro con una o mas tablas dinamicas
- Ejecutar la macro; actualiza TODAS las tablas de todas las hojas
- Asignar a un boton: Insertar > Formas > clic derecho > Asignar macro
### Prompt para generarla

```
Genera una macro VBA llamada ActualizarPivots. Debe recorrer todas las hojas del libro activo, y para cada hoja recorrer todas sus tablas dinamicas y ejecutar RefreshTable. Contar cuantas tablas se actualizaron y mostrarlo en un MsgBox. Comentarios en espaniol.
```

---

## Macro 5: ExportarPDF()

### Proposito

Exporta la hoja activa como archivo PDF. El nombre del archivo incluye el nombre de la hoja y la fecha actual. Guarda en la misma carpeta del libro.

### Codigo VBA

```
Sub ExportarPDF()
    Dim ws As Worksheet
    Dim rutaPDF As String
    Dim rutaLibro As String
    
    Set ws = ActiveSheet
    rutaLibro = ThisWorkbook.Path
    
    If rutaLibro = "" Then
        MsgBox "Guarde el libro primero.", vbExclamation
        Exit Sub
    End If
    
    rutaPDF = rutaLibro & "\" & ws.Name & "_" & _
             Format(Date, "YYYY-MM-DD") & ".pdf"
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=True
    
    MsgBox "PDF exportado: " & rutaPDF, vbInformation
End Sub
```

### Como usar

- Abrir la hoja que desea exportar como PDF
- Configurar area de impresion si es necesario (Vista > Saltos de pagina)
- Ejecutar la macro; el PDF se guarda en la misma carpeta del libro
- El archivo se abre automaticamente despues de exportar
### Prompt para generarla

```
Genera una macro VBA llamada ExportarPDF. Debe exportar la hoja activa como PDF usando ExportAsFixedFormat. El nombre del archivo debe ser [NombreHoja]_[Fecha].pdf y guardarse en la misma carpeta del libro. Verificar que el libro este guardado antes de exportar. Abrir el PDF despues de crearlo. MsgBox con la ruta. Comentarios en espaniol.
```

---

## Tips Finales

- **Personal.xlsb:** Guarda tus macros mas utiles en el libro Personal (Archivo > Opciones > Guardar macro en > Libro de macros personal). Asi estan disponibles en todos tus archivos.
- **Seguridad:** Solo habilita macros de fuentes confiables. Configura el Trust Center en: Archivo > Opciones > Centro de confianza.
- **Respaldo:** Siempre prueba macros en una copia del archivo, nunca en el original.
- **Errores:** Si una macro da error, copia el mensaje de error y el codigo, pegalo en Claude o ChatGPT, y pide explicacion y correccion.
- **Documentacion:** Agrega comentarios (lineas con ') a tus macros para que tu yo futuro entienda que hacen.