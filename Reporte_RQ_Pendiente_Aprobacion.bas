Attribute VB_Name = "Reporte_RQ_Pendiente_Aprobacion"
Private Sub PASO1()
'
' Eliminar las 7 primeras y las 2 ultimas
'
'
    Dim n As Double
    
    
    Range("A1").Select
    ActiveSheet.Shapes.Range(Array("image1.png")).Select
    Selection.Delete
    
    Range("A1").Select
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    
    Range("A1").Select
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Range(Cells(n + 1, 1), Cells(n + 2, 30)).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select

End Sub

Private Sub PASO2()
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
        .MergeCells = False
    End With
    
    Range("A1").Select
    
End Sub

Private Sub PASO3()
Attribute PASO3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Formato
'
    Range("A1").Select
    Columns("AC:AC").ColumnWidth = 8
    Columns("AD:AD").ColumnWidth = 8
    
    Range("AC1") = "CDG"
    Range("AD1") = "Aceptar/Rechazar"
    Range("AE1") = "Comentarios"
    
'---------------------------------------------------------------------
    Application.Union(Range("A1"), Range("E1"), Range("F1"), Range("K1"), Range("L1"), Range("P1"), Range("Z1"), Range("AC1")).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 7
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
'---------------------------------------------------------------------
    Range("AD1", "AE1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 7
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    Range("A1").Select
    
End Sub

Private Sub PASO4()
Attribute PASO4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Formula CDG
'
    
    Dim n As Double
    
    Range("A1").Select
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("AC2").Select
    
    ActiveCell.Formula2R1C1 = _
        "=IF(AND(RC5=""Pendiente Aprobación"",RC6=""JEFE DE CONTROL DE COSTOS"")," & _
        "IFS(" & _
        "OR(Correos.xlsx!R2C2:R2C26=RC[-13]),""VP Redes - MG""," & _
        "OR(Correos.xlsx!R3C2:R3C26=RC[-13]),""VP TI""," & _
        "OR(Correos.xlsx!R4C2:R4C26=RC[-13]),""VP TI - Logistica""," & _
        "OR(Correos.xlsx!R5C2:R5C26=RC[-13]),""VP TI - Facturación""," & _
        "OR(Correos.xlsx!R6C2:R6C26=RC[-13]),""GC Planeamiento""," & _
        "OR(Correos.xlsx!R7C2:R7C26=RC[-13]),""GC Finanzas""," & _
        "OR(Correos.xlsx!R8C2:R8C26=RC[-13]),""GC Legal""," & _
        "OR(Correos.xlsx!R9C2:R9C26=RC[-13]),""GC Capital Humano""," & _
        "OR(Correos.xlsx!R10C2:R10C26=RC[-13]),""VP Redes - Controller""),"""")"
        
    Range("AC2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 29), Cells(n, 29))
    
    ActiveWorkbook.Save
        
    Range("A1").Select
End Sub
Private Sub PASO5(ByVal Fecha As String)
'
' Guardar archivo
'

    ruta = "C:\...\RQs\Pendiente Aprobación\RQs pendientes de aprobación\"
    'ruta = "C:\...\RQs\Pendiente Aprobación\Prueba\"
    Filename = "TS 03.10 Historico de RQ_" & Fecha
    
    Range("A1").Select
    ActiveWorkbook.SaveAs Filename:=ruta & Filename & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=True
    ActiveWorkbook.RunAutoMacros Which:=xlAutoClose
    Range("A1").Select


End Sub
Private Sub PASO6(ByVal Fecha As String)
'
' Abrir excel a trabajar
'

    ruta = "C:\...\RQs\Pendiente Aprobación\"
    'ruta = "C:\...\RQs\Pendiente Aprobación\Prueba\"
    Filename = "TS 03.10 Historico de RQs_010622-" & Fecha & "_" & Fecha & ".xlsx"
    
    Workbooks.Open (ruta & Filename)



End Sub

Private Sub PASO7(ByVal Fecha As String)
'
' RQs a enviar / RQs a detalle
' Modificar el archivo a enviar
'

    Workbooks("TS 03.10 Historico de RQ_" & Fecha & ".xlsx").Activate
    
    ActiveWorkbook.Sheets("Sheet1").Copy After:=ActiveWorkbook.Sheets("Sheet1")
    Sheets("Sheet1").Activate
    ActiveSheet.Name = "RQs_a_revisar"

    Sheets("Sheet1 (2)").Activate
    ActiveSheet.Name = "RQs_a_detalle"
    
    Sheets("RQs_a_revisar").Activate
    Range("B:B").Columns.Delete
    Range("I:I").Columns.Delete
    Range("T:T").Columns.Delete
    Range("T:T").Columns.Delete
    
    Range("A1").Select
    Selection.AutoFilter
    
    Dim n As Double
    Dim x As Double
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ActiveSheet.Range("$A$1:$AA$" & n).AutoFilter _
    Field:=4, _
    Criteria1:="<>Pendiente Aprobación", _
    Operator:=xlAnd
    
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ActiveSheet.Range("$A$1:$AA$" & n).AutoFilter _
    Field:=5, _
    Criteria1:="<>JEFE DE CONTROL DE COSTOS", _
    Operator:=xlAnd
    
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("$A$1:$AA$" & n).RemoveDuplicates Columns:=1, Header:=xlYes
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    For x = 2 To n
    
        Sheets("RQs_a_revisar").Range("I" & x) = WorksheetFunction.SumIfs(Sheets("RQs_a_detalle").Range("K:K"), _
                                                Sheets("RQs_a_detalle").Range("A:A"), Range("A" & x))
                                                
        Sheets("RQs_a_revisar").Range("J" & x) = WorksheetFunction.SumIfs(Sheets("RQs_a_detalle").Range("L:L"), _
                                                Sheets("RQs_a_detalle").Range("A:A"), Range("A" & x))
                                                
    Next
        
End Sub

Sub Reporte_RQ_Pendiente_Aprobacion()
'
' Todos los pasos
'
'
    tiempo = Now
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    
    Range("A1").Select
    
'FECHA--------------------------------------------
    If Day(Date) < 10 Then
        dia = "0" & Day(Date)
    Else
        dia = Day(Date)
    End If
    
    If Month(Date) < 10 Then
        mes = "0" & Month(Date)
    Else
        mes = Month(Date)
    End If
    
    anio = Right(Year(Date), 2)
    
    Fecha = dia & mes & anio
'FECHA--------------------------------------------
    
    Call PASO1
    Call PASO2
    Call PASO3
    Call PASO4
    Call PASO5(Fecha)
    Call PASO6(Fecha)
    Call PASO7(Fecha)
    
    Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
    tiempo = DateDiff("n", tiempo, Now)
    MsgBox ("LA MACRO HA DURADO: " & tiempo & " MINUTOS")
    
End Sub
