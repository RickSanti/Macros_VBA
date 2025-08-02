Attribute VB_Name = "Cambiar_formato_LM_AnaliticoPEN"

Private Sub PASO1()
Attribute PASO1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Eliminar las 2 primeras y las 3 ultimas
'
'
    Range("A1").Select
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Range(Cells(n + 1, 1), Cells(n + 3, 30)).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub
Private Sub PASO2()
Attribute PASO2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Desajustar toda la tabla
'

'
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("A1").Select
End Sub
Private Sub PASO3()
'
' Eliminar filas con el registro "Saldo inicial" de la columna B
'
'
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(n, 28)).AutoFilter Field:=2, Criteria1:= _
        "Saldo Inicial"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("A1").Select
End Sub
Private Sub PASO4()
'
' Descombinar las últimas columnas y eliminar las 2 últimas
'
'
    Range("AA1:AB1").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("AB1:AD1").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
End Sub

Private Sub PASO5()
'
' Volver la columna AA en formato numerico
'

'
    Range("AA2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("AA2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
    Range("A1").Select
End Sub
Private Sub PASO6()
'
' Crear columna Monto Funcional Neto
'

'
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Columns("Y:Y").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "Monto Funcional Neto"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=+RC[-1]-RC[-2]"
    Range("Y2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 25), Cells(n, 25))
    Range("Y1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub PASO7()
'
' Crear Tabla Dinamica
'

'
    Range("A1").Select
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(n, 28)), , xlYes).Name = _
        "Tabla1"
    Range("Tabla1[[#Headers],[Periodo]]").Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla1", Version:=7).CreatePivotTable TableDestination:="Hoja1!R3C1", _
        TableName:="TablaDinámica1", DefaultVersion:=7
    Sheets("Hoja1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("TablaDinámica1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("TablaDinámica1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("TablaDinámica1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica1").PivotFields("Centro Costo")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("TablaDinámica1").PivotFields("Cuenta contable")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("TablaDinámica1").AddDataField ActiveSheet.PivotTables( _
        "TablaDinámica1").PivotFields("Monto Funcional Neto"), _
        "Suma de Monto Funcional Neto", xlSum
        
    Range("C1").Select
    
    Sheets("Hoja1").Select
    Sheets("Hoja1").Name = "TD"
    
    
    
End Sub

Private Sub PASO8()
'
' Cambiar la tabla1 a Rango
'
'

    Sheets("Sheet1").Select
    Range("A1").Select
    
    ActiveSheet.ListObjects(1).Unlist
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = False
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    
    Range("A1").Select

End Sub

Private Sub PASO9()
'
' Cambiar el nombre de la hoja sheet1 a "BASE LM + fecha hoy"
'


    y = Year(Now)
    m = Month(Now)
    d = Day(Now)
    
    If d < 10 Then
        d = "0" & d
    End If
    
    If m < 10 Then
        m = "0" & m
    End If
    
    ' Sheets("Sheet1").Range("A1").Select
    
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "BASE LM " & d & "-" & m & "-" & y

End Sub

Sub Cambiar_formato_LM_AnaliticoPEN()
'
' Todos los pasos
'
'
    Range("A1").Select
    
    PASO1
    PASO2
    PASO3
    PASO4
    PASO5
    PASO6
    PASO7
    PASO8
    PASO9
    
    
    
End Sub
