Attribute VB_Name = "Prueba_debito_v3"
Private Sub rep_suministros(ByVal nameSU As String, ByVal nameBCP As String, ByVal cell_SU As Range, ByVal cell_BCP As Range, ByVal nSU As Double, ByVal nBCP As Double)
            
'Busqueda de Suministros-------------------------------------------------------------------------------------------------------------------
    
    Dim aux As Double
    
    Workbooks(nameSU).Activate
    'Sheets("Suministros Eléctricos").Select
    
    Range("S1") = "Repeticiones"
    
    'For Each cell_SU In Sheets("Suministros Eléctricos").Range(Cells(2, 13), Cells(nSU, 13))
    For Each cell_SU In Range(Cells(2, 13), Cells(nSU, 13))
    
        aux = 0
        
        If cell_SU.Offset(0, -10) = "CREATED" Or cell_SU.Offset(0, -10) = "ON AIR" Then
        
            Workbooks(nameBCP).Activate
            'Sheets("Sheet0").Select
                    
            For Each cell_BCP In Range(Cells(6, 3), Cells(nBCP, 3))
            
                If InStr(1, cell_BCP, cell_SU.Value, 1) > 0 Then
                    
                    aux = aux + 1
                                        
                End If
                          
            Next
            
            Workbooks(nameSU).Activate
            
            cell_SU.Offset(0, 6) = aux
            
        End If
        
    Next
    
End Sub

Private Sub TextoAColumnas(ByVal nameBCP As String, ByVal cell_BCP As Range, ByVal nBCP As Double)
    
    Workbooks(nameBCP).Activate
    
    'Range("F5:M" & nBCP).NumberFormat = "Texto"
    
    Range("F5") = "Delimitador 1"
    Range("G5") = "Delimitador 2"
    Range("H5") = "Delimitador 3"
    Range("I5") = "Delimitador 4"
    Range("J5") = "Delimitador 5"
    Range("K5") = "Delimitador 6"
    Range("L5") = "Delimitador 7"
    Range("M5") = "Delimitador 8"
    
    Range(Cells(6, 3), Cells(nBCP, 3)).TextToColumns _
    Destination:=Range(Cells(6, 6), Cells(nBCP, 6)), _
    DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=True, _
    Space:=True, _
    FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2))
    
End Sub
Private Sub concesionaria(ByVal nameSU As String, ByVal nameBCP As String, ByVal cell_SU As Range, ByVal cell_BCP As Range, ByVal nSU As Double, ByVal nBCP As Double)
            
'Busqueda de concesionarias-------------------------------------------------------------------------------------------------------------------
    
    Workbooks(nameBCP).Activate
    
    Range("E5") = "Concesionaria"
    
    For Each cell_BCP In Range(Cells(6, 3), Cells(nBCP, 3))
    
        Workbooks(nameSU).Activate
                
        For Each cell_SU In Range(Cells(2, 13), Cells(nBCP, 13))
        
            If InStr(1, cell_BCP, cell_SU.Value, 1) > 0 Then
                
                cell_BCP.Offset(0, 2) = cell_SU.Offset(0, -4).Value
                Exit For
                                    
            End If
                      
        Next
                
        Workbooks(nameSU).Activate
        
    Next
    
End Sub
Private Sub Contar_suministro_x_delimitador(ByVal nameSU As String, ByVal nameBCP As String, ByVal nSU As Double)
    
    Workbooks(nameSU).Activate
    
    Range("T1") = "Delimitador 1"
    Range("U1") = "Delimitador 2"
    Range("V1") = "Delimitador 3"
    Range("W1") = "Delimitador 4"
    Range("X1") = "Delimitador 5"
    Range("Y1") = "Delimitador 6"
    Range("Z1") = "Delimitador 7"
    Range("AA1") = "Delimitador 8"
    
    For x = 2 To nSU
        Range("T" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("F:F"), Range("M" & x).Value)
        Range("U" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("G:G"), Range("M" & x).Value)
        Range("V" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("H:H"), Range("M" & x).Value)
        Range("W" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("I:I"), Range("M" & x).Value)
        Range("X" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("J:J"), Range("M" & x).Value)
        Range("Y" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("K:K"), Range("M" & x).Value)
        Range("Z" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("L:L"), Range("M" & x).Value)
        Range("AA" & x) = Application.WorksheetFunction.CountIf(Workbooks(nameBCP).Worksheets("Sheet0").Range("M:M"), Range("M" & x).Value)
    Next
    
End Sub

Sub main_Verificacion_Cuentas_no_pagadas()
    
    MsgBox "A continuación escribir los nombres de los archivos a usar (Recordar tener ambos archivos abiertos)", , "Verificación débito automático"
    
    nameSU = Application.InputBox(Prompt:="Escriba el nombre del archivo de los suministros", Title:="Archivo de suministros", Default:="Ejemplo1.xlsx o Ejemplo1.xls", Type:=2)
    nameBCP = Application.InputBox(Prompt:="Escriba el nombre del archivo de pagos de débito automático", Title:="Archivo de pagos", Default:="Ejemplo1.xlsx o Ejemplo1.xls", Type:=2)
    
    tiempo = Now
        
'Variables globales-------------------------------------------------------------------------------------------------------------------
    'variables celdas de cada hoja
    Dim cell_SU As Range
    Dim cell_BCP As Range
    
    'variables de cantidad de lineas
    Dim nSU As Double
    Dim nBCP As Double

'Cantidad de lineas-------------------------------------------------------------------------------------------------------------------
    
    'cantidad de lineas de SUMINISTRO
    Workbooks(nameSU).Activate
    'Sheets("Suministros Eléctricos").Select
    nSU = Cells(Rows.Count, "A").End(xlUp).Row
    
    'cantidad de lineas de BCP
    Workbooks(nameBCP).Activate
    'Sheets("Sheet0").Select
    nBCP = Cells(Rows.Count, "A").End(xlUp).Row

'Funciones-------------------------------------------------------------------------------------------------------------------
    
    Call rep_suministros(nameSU, nameBCP, cell_SU, cell_BCP, nSU, nBCP)
    Call TextoAColumnas(nameBCP, cell_BCP, nBCP)
    Call concesionaria(nameSU, nameBCP, cell_SU, cell_BCP, nSU, nBCP)
    Call Contar_suministro_x_delimitador(nameSU, nameBCP, nSU)
    
    tiempo = DateDiff("n", tiempo, Now)
    MsgBox ("LA MACRO HA DURADO: " & tiempo & " MINUTOS")
    
End Sub
