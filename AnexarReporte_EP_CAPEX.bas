Attribute VB_Name = "AnexarReporte_EP_CAPEX"
Sub AnexarReporte_EP_CAPEX()
Attribute AnexarReporte_EP_CAPEX.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro para anexar las semanas
'

' Agrega a la columna AK (Reporte Semanal) la semana a la que pertenece y copia el contenido
' ---------------------------------------------------------------------------------------------------
    'sem es el numero de la semana actual
    sem = WorksheetFunction.WeekNum(Date)
    
    'nombre del archivo recibido de la semana
    archivo = "Ejecución Presupuestal - CAPEX REDES W" & sem & " - 2022.xlsb"
    
    Windows(archivo).Activate
    Sheets("6. Comprometido Financiero").Select
    
    Range("A1").Select
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("AK4").Select
    ActiveCell.FormulaR1C1 = "W" & sem
    Range("AK5").Select
    ActiveCell.FormulaR1C1 = "W" & sem
    Range("AK4:AK5").Select
    Selection.AutoFill Destination:=Range(Cells(4, 37), Cells(n, 37))
    
    Range(Cells(4, 1), Cells(n, 37)).Select
    
    Selection.Copy
    
' Pega el contenido en el archivo requerido
' ---------------------------------------------------------------------------------------------------
    
    sem = WorksheetFunction.WeekNum(Date) - 1
    'nombre del archivo a donde se pegará los datos
    archivo = "Ejecución Presupuestal - CAPEX REDES W" & sem & " - 2022.xlsb"
    
    Windows(archivo).Activate
    Sheets("6. Comprometido Financiero").Select
    
    Range("AK3") = "SemanaReporte"
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Cells(n + 1, 1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub
