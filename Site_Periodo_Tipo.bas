Attribute VB_Name = "Site_Periodo_Tipo"
Private Sub Sitio_v5(ByVal nameLM As String, ByVal nameADS As String)
'
' Verificación_ADS Macro
'
    'tiempo = Now
    
    'variables de cantidad de lineas
    Dim nLM As Double
    Dim nEN As Double
    Dim nDN As Double
    Dim nAM As Double
    Dim nOP As Double
    Dim nPO As Double
    Dim nDE As Double
    
    'variables celdas de cada hoja
    Dim cell_LM As Range
    Dim cell_EN As Range
    Dim cell_DN As Range
    Dim cell_AM As Range
    Dim cell_OP As Range
    Dim cell_PO As Range
    Dim cell_DE As Range
    
    'variables auxiliares para la busqueda por codigo
    Dim aux_LM As String
    Dim aux_ADS As String


'Cantidad de lineas en LM-------------------------------------------------------------------------------------------------------------------
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    nLM = Cells(Rows.Count, "A").End(xlUp).Row
    
    'cantidad de lineas de ADS
    Workbooks(nameADS).Activate
    
    Sheets("01 Entel").Select
    nEN = Sheets("01 Entel").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("02 Direcnet").Select
    nDN = Sheets("02 Direcnet").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("03 Americatel").Select
    nAM = Sheets("03 Americatel").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("04 Otros Proyectos").Select
    nOP = Sheets("04 Otros Proyectos").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("05 Postes").Select
    nPO = Sheets("05 Postes").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("Desmontados").Select
    nDE = Sheets("Desmontados").Cells(Rows.Count, "A").End(xlUp).Row
    
'Busqueda de NombreSite por NombreSite-------------------------------------------------------------------------------------------------------------------
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    
    For Each cell_LM In Sheets("Sheet1").Range(Cells(2, 22), Cells(nLM, 22))
        
        'Revisa en 01 entel
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("01 Entel").Select
                    
            For Each cell_EN In Sheets("01 Entel").Range(Cells(2, 2), Cells(nEN, 2))
            
                If InStr(1, cell_LM.Value, cell_EN.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_EN.Value
                    cell_LM.Offset(0, 1) = cell_EN.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 02 Direcnet
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("02 Direcnet").Select
                    
            For Each cell_DN In Sheets("02 Direcnet").Range(Cells(2, 2), Cells(nDN, 2))
            
                If InStr(1, cell_LM.Value, cell_DN.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_DN.Value
                    cell_LM.Offset(0, 1) = cell_DN.Offset(0, -1).Value
                    cell_LM.Offset(0, 2) = "Direcnet"
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 03 Americatel
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("03 Americatel").Select
                    
            For Each cell_AM In Sheets("03 Americatel").Range(Cells(2, 2), Cells(nAM, 2))
            
                If InStr(1, cell_LM.Value, cell_AM.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_AM.Value
                    cell_LM.Offset(0, 1) = cell_AM.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 04 Otros Proyectos
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("04 Otros Proyectos").Select
                    
            For Each cell_OP In Sheets("04 Otros Proyectos").Range(Cells(2, 2), Cells(nOP, 2))
            
                If InStr(1, cell_LM.Value, cell_OP.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_OP.Value
                    cell_LM.Offset(0, 1) = cell_OP.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 05 Postes
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("05 Postes").Select
                    
            For Each cell_PO In Sheets("05 Postes").Range(Cells(2, 2), Cells(nPO, 2))
            
                If InStr(1, cell_LM.Value, cell_PO.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_PO.Value
                    cell_LM.Offset(0, 1) = cell_PO.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en Desmontados
        If cell_LM.Offset(0, -3) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("Desmontados").Select
                    
            For Each cell_DE In Sheets("Desmontados").Range(Cells(2, 2), Cells(nDE, 2))
            
                If InStr(1, cell_LM.Value, cell_DE.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_DE.Value
                    'cell_LM.Offset(0, -5) = cell_EN.Offset(0, -1).value
                    cell_LM.Offset(0, 1) = "Desmontado"
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
    Next
        
''Busqueda de NombreSite por CodigoProyecto-------------------------------------------------------------------------------------------------------------------
'
'    Workbooks(nameLM).Activate
'    Sheets("Sheet1").Select
'
'    For Each cell_LM In Sheets("Sheet1").Range(Cells(2, 22), Cells(nLM, 22))
'
'        aux_LM = cell_LM.Offset(0, -7).value
'
'        'Revisa en 01 entel
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("01 Entel").Select
'
'            For Each cell_EN In Sheets("01 Entel").Range(Cells(2, 2), Cells(nEN, 2))
'
'                aux_ADS = cell_EN.Offset(0, 9).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_EN.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'        'Revisa en 02 Direcnet
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("02 Direcnet").Select
'
'            For Each cell_DN In Sheets("02 Direcnet").Range(Cells(2, 2), Cells(nDN, 2))
'
'                aux_ADS = cell_DN.Offset(0, 7).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_DN.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'        'Revisa en 03 Americatel
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("03 Americatel").Select
'
'            For Each cell_AM In Sheets("03 Americatel").Range(Cells(2, 2), Cells(nAM, 2))
'
'                aux_ADS = cell_AM.Offset(0, 7).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_AM.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'        'Revisa en 04 Otros Proyectos
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("04 Otros Proyectos").Select
'
'            For Each cell_OP In Sheets("04 Otros Proyectos").Range(Cells(2, 2), Cells(nOP, 2))
'
'                aux_ADS = cell_OP.Offset(0, 8).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_OP.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'        'Revisa en 05 Postes
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("05 Postes").Select
'
'            For Each cell_PO In Sheets("05 Postes").Range(Cells(2, 2), Cells(nPO, 2))
'
'                aux_ADS = cell_PO.Offset(0, 7).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_PO.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'        'Revisa en Desmontados
'        If cell_LM.Offset(0, -3) = "" Then
'
'            'cell.Select
'            Workbooks(nameADS).Activate
'            Sheets("Desmontados").Select
'
'            For Each cell_DE In Sheets("Desmontados").Range(Cells(2, 2), Cells(nDE, 2))
'
'                aux_ADS = cell_DE.Offset(0, -1).value
'
'                If aux_ADS = aux_LM Then
'
'                    'cell2.Select
'                    Workbooks(nameLM).Activate
'                    Sheets("Sheet1").Select
'                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
'                    cell_LM.Offset(0, -3) = cell_DE.value
'                    Exit For
'
'                End If
'
'            Next
'
'            Workbooks(nameLM).Activate
'            Sheets("Sheet1").Select
'
'        End If
'
'    Next
        
    'tiempo = DateDiff("n", tiempo, Now)
    'MsgBox ("LA MACRO HA DURADO: " & tiempo & " MINUTOS")
    
End Sub

Private Sub PeriodoGasto(ByVal nameLM As String)
'
'  Sacar periodo de la columna Descripción de Linea
'
'       Solo es apoyo, las faltantes se deben hacer a mano
'
    Dim r As Range
    Dim n As Double
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        
    ' COLUMNA PERIODO DEL GASTO
    
    For Each r In Range(Cells(2, 22), Cells(n, 22))
    
        If r.Offset(0, -2) = "" Then
        
            If Left(r.Value, 13) = "REVER. PROV. " Then
            
                mes = Mid(r.Value, 14, 3)
                anio = Mid(r.Value, 18, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
            ElseIf Left(r.Value, 19) = "REVER. SALDO PROV. " Then
            
                mes = Mid(r.Value, 20, 3)
                anio = Mid(r.Value, 24, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
            ElseIf Left(r.Value, 35) = "REVER. Saldo Revers. Revers. PROV. " Then
            
                mes = Mid(r.Value, 36, 3)
                anio = Mid(r.Value, 40, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
            ElseIf Left(r.Value, 6) = "PROV. " Then
            
                mes = Mid(r.Value, 7, 3)
                anio = Mid(r.Value, 11, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
            ElseIf Left(r.Value, 20) = " - NO DEDUC PAGO GD-" Then
            
                mes = Mid(r.Value, 26, 3)
                anio = Mid(r.Value, 30, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                            
            ElseIf Left(r.Value, 7) = " - GD-0" Then
            
                mes = Mid(r.Value, 14, 3)
                anio = Mid(r.Value, 18, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
            ElseIf Left(r.Value, 22) = " - GD-670/18 ALQ.SITE " Then
                
                mes = Mid(r.Value, 23, 3)
                anio = Mid(r.Value, 27, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                'r.Offset(0, -2) = Mid(r.Value, 23, 6)
                
            ElseIf Left(r.Value, 1) = "3" Or Left(r.Value, 1) = "4" Or Left(r.Value, 1) = "5" Then
                
                If InStr(1, r.Value, "RENOV") Then
                    
                    r.Offset(0, -2) = "NA"
                
                Else
                
                    mes = Mid(r.Value, 10, 3)
                    anio = Mid(r.Value, 14, 2)
                    r.Offset(0, -2) = "01/" & mes & "/20" & anio
                    
                End If
            
    '        ElseIf Left(r.Value, 18) = " - DEVENG.ALQ.SITE" Then
    '
    '            mes = Mid(r.Value, 10, 3)
    '            anio = Mid(r.Value, 14, 2)
    '            r.Offset(0, -2) = "01/" & mes & "/20" & anio
                
                
    ' CELDAS TIPO -MES-AÑO / -MES-AÑO A
            ElseIf Left(r.Value, 6) = " - ENE" Or Left(r.Value, 6) = " - Ene" _
                Or Left(r.Value, 6) = " - FEB" Or Left(r.Value, 6) = " - Feb" _
                Or Left(r.Value, 6) = " - MAR" Or Left(r.Value, 6) = " - Mar" _
                Or Left(r.Value, 6) = " - ABR" Or Left(r.Value, 6) = " - Abr" _
                Or Left(r.Value, 6) = " - MAY" Or Left(r.Value, 6) = " - May" _
                Or Left(r.Value, 6) = " - JUN" Or Left(r.Value, 6) = " - Jun" _
                Or Left(r.Value, 6) = " - JUL" Or Left(r.Value, 6) = " - Jul" _
                Or Left(r.Value, 6) = " - AGO" Or Left(r.Value, 6) = " - Ago" _
                Or Left(r.Value, 6) = " - SET" Or Left(r.Value, 6) = " - Set" _
                Or Left(r.Value, 6) = " - OCT" Or Left(r.Value, 6) = " - Oct" _
                Or Left(r.Value, 6) = " - NOV" Or Left(r.Value, 6) = " - Nov" _
                Or Left(r.Value, 6) = " - DIC" Or Left(r.Value, 6) = " - Dic" _
                Then
    
                mes = Mid(r.Value, 4, 3)
                anio = Mid(r.Value, 8, 2)
                r.Offset(0, -2) = "01/" & mes & "/20" & anio
                            
    ' CELDAS SIN FECHA--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf Left(r.Value, 10) = "NC X APLIC" Or Left(r.Value, 13) = "PROV.ABR ALQ." Or Left(r.Value, 26) = "REVER. Saldo Revers. ALQ. " Or Left(r.Value, 12) = "REVER. ALQ. " Then
            
                'mes = Mid(r.Value, 14, 3)
                'anio = Mid(r.Value, 18, 2)
                'r.Offset(0, -2) = "01/" & mes & "/20" & anio
                r.Offset(0, -2) = "NA"
            
            End If
            
        End If
        
    Next
    
    Columns("T:T").Select
    Selection.TextToColumns Destination:=Range("T1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
    Range("T1").Select
       
End Sub

Private Sub TipoRenta(ByVal nameLM As String)
'
'  Sacar tipo de renta de la columna Descripción de Linea
'
'       Falta el tipo RS
'
    Dim r As Range
    Dim n As Double
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    
    n = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        
    ' COLUMNA TIPO DE RENTA
    
    For Each r In Range(Cells(2, 22), Cells(n, 22))
    
        If r.Offset(0, -1) = "" Then
            
            '''soles opex tambien es PISO (OK)
            If InStr(1, r.Value, " PISO ") > 0 Or InStr(1, r.Value, "PISO|") > 0 Or InStr(1, r.Value, "| Piso-") > 0 Or InStr(1, r.Value, "USD  OPEX PISO") > 0 Or InStr(1, r.Value, "| PisoSITE") > 0 Or InStr(1, r.Value, "SOLES OPEX") > 0 Or InStr(1, r.Value, "PEN OPEX") > 0 Or InStr(1, r.Value, "USD OPEX") > 0 Then
            
                r.Offset(0, -1) = "PISO"
            ' validar si dice suelo en la columna siguiente y poner PISO
            
            ElseIf InStr(1, r.Value, "PEN ALBA FASE") > 0 Or InStr(1, r.Value, "USD ALBA FASE") > 0 Then
            
                r.Offset(0, -1) = "ALBA FASE"
                
            ElseIf InStr(1, r.Value, "COLO USD DIFERIDO|") > 0 Or InStr(1, r.Value, "COLO PEN DIFERIDO|") > 0 Then
            
                r.Offset(0, -1) = "COLO"
                
            ElseIf InStr(1, r.Value, "LMT|") > 0 Or InStr(1, r.Value, " LMT -") > 0 Or InStr(1, r.Value, "LMT PEN DIFERIDO|") > 0 Or InStr(1, r.Value, "| LMTSITE") > 0 Or InStr(1, r.Value, "| LMT-") > 0 Or InStr(1, r.Value, "LMT DIF|") > 0 Or InStr(1, r.Value, "LMT DIFERIDO USD|") > 0 Then
            
                r.Offset(0, -1) = "LMT"
            
            ElseIf InStr(1, r.Value, "TORRE|") > 0 Or InStr(1, r.Value, "TORRE PEN DIFERIDO|") > 0 Or InStr(1, r.Value, "| Torre-") > 0 Or InStr(1, r.Value, " TORRE -") > 0 Or InStr(1, r.Value, "TORRE PEN DIF|") > 0 Or InStr(1, r.Value, "| TorreSITE") > 0 Or InStr(1, r.Value, "TORRE DIF|") > 0 Or InStr(1, r.Value, "TORRE DIFERIDO USD|") > 0 Or InStr(1, r.Value, "TORRE DIF USD|") > 0 Or InStr(1, r.Value, "TORRE DIFERIDO PEN|") > 0 Or InStr(1, r.Value, "TORRE DIFERIDO PEN") > 0 Or InStr(1, r.Value, "TORRE USD OPEX|") > 0 Or InStr(1, r.Value, "TORRE PEN DIFERIDO -") > 0 Or InStr(1, r.Value, "TORRE/") > 0 Then
            
                r.Offset(0, -1) = "TORRE"
                
            ElseIf InStr(1, r.Value, "POSTES") > 0 Or InStr(1, r.Value, "| Postes-") > 0 Or InStr(1, r.Value, "PostesSITE") > 0 Then
            
                r.Offset(0, -1) = "POSTE"
                
            ElseIf InStr(1, r.Value, "MANTENIMIENTO") > 0 Or InStr(1, r.Value, "Mantenimiento") > 0 Or InStr(1, r.Value, "mantenimiento") > 0 Then
            
                r.Offset(0, -1) = "MANTENIMIENTO"
                
            ElseIf InStr(1, r.Value, "RANSHARING|") > 0 Or InStr(1, r.Value, " RANSHARING ") > 0 Or InStr(1, r.Value, "| RANSHARING") > 0 Or InStr(1, r.Value, "RAN|") > 0 Or InStr(1, r.Value, "RAN ") > 0 Or InStr(1, r.Value, "| RNS") > 0 Or InStr(1, r.Value, "RANS PEN DIFERIDO|") > 0 Then
            
                r.Offset(0, -1) = "RAN SHARING"
                
            ElseIf InStr(1, r.Value, "Ead PEN DIF|") > 0 Or InStr(1, r.Value, "Ead 2 PEN DIF|") > 0 Or InStr(1, r.Value, "EAD|") > 0 Or InStr(1, r.Value, "/EAD") > 0 Or InStr(1, r.Value, "| EAD") > 0 Or InStr(1, r.Value, "EAD DIFERIDO USD|") > 0 Then
     
                 r.Offset(0, -1) = "EAD"
            
            'tomarlo como EAD (OK)
            ElseIf InStr(1, r.Value, " EA ") > 0 Or InStr(1, r.Value, "EA|") > 0 Or InStr(1, r.Value, "| EA-") > 0 Or InStr(1, r.Value, "EA-NOTA DEBITO|") > 0 Then
     
                 r.Offset(0, -1) = "EAD"
                
    '        'tomarlo como RanSharing (OK)
    '        ElseIf InStr(1, r.Value, "RAN|") > 0 Or InStr(1, r.Value, "RAN -") > 0 Then
    '
    '             r.Offset(0, -1) = "RAN SHARING"
                 
            ElseIf InStr(1, r.Value, "BA|") > 0 Then
    
                 r.Offset(0, -1) = "BA"
            
    '        'tomarlo como RanSharing (OK)
    '        ElseIf InStr(1, r.Value, "| RNS") > 0 Then
    '
    '            r.Offset(0, -1) = "RAN SHARING"
                
            ElseIf InStr(1, r.Offset(0, 2), " suelo ") Or InStr(1, r.Offset(0, 2), " suelocorrespondiente ") Or InStr(1, r.Offset(0, 2), " sueloCorrespondiente ") Then
            
                r.Offset(0, -1) = "PISO"
                
            End If
            
        End If
            
    Next
    
End Sub
Private Sub Empresa(ByVal nameLM As String, ByVal nameDS As String)
'
'  Sacar empresa
'
'
    Dim cell_LM As Range
    Dim cell_DS As Range
    Dim n_LM As Double
    Dim n_DS As Double
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    
    n_LM = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        
    'COLUMNA EMPRESA
    
    For Each cell_LM In Range(Cells(2, 22), Cells(n_LM, 22))
    
        If cell_LM.Offset(0, 2) = "" Then
            
            Workbooks(nameDS).Activate
            Sheets("Base").Select
                    
            For Each cell_DS In Sheets("Base").Range(Cells(2, 2), Cells(nEN, 2))
            
                If InStr(1, cell_LM.Value, cell_EN.Value, 1) > 0 Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    'cell.Offset(0, 12) = cell2.Offset(0, -1).Value
                    cell_LM.Offset(0, -3) = cell_EN.Value
                    cell_LM.Offset(0, 1) = cell_EN.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
            
    Next
    
    For Each cell_LM In Range(Cells(2, 21), Cells(n, 21))
    
        If cell_LM.Offset(0, 3) = "" Then
            
            'Empresa POSTE
            'Revisar las empresas Entel que son poste
            If cell_LM.Value = "POSTE" Then
            
                cell_LM.Offset(0, 3) = "POSTE"
                            
            End If
            
        End If
            
    Next
    
End Sub
Private Sub ID_ADS(ByVal nameLM As String, ByVal nameADS As String)
'
' Verificación_ADS Macro
'
    'tiempo = Now
    
    'variables de cantidad de lineas
    Dim nLM As Double
    Dim nEN As Double
    Dim nDN As Double
    Dim nAM As Double
    Dim nOP As Double
    Dim nPO As Double
    Dim nDE As Double
    
    'variables celdas de cada hoja
    Dim cell_LM As Range
    Dim cell_EN As Range
    Dim cell_DN As Range
    Dim cell_AM As Range
    Dim cell_OP As Range
    Dim cell_PO As Range
    Dim cell_DE As Range
    
    'variables auxiliares para la busqueda por codigo
    Dim aux_LM As String
    Dim aux_ADS As String


'Cantidad de lineas en LM-------------------------------------------------------------------------------------------------------------------
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    nLM = Cells(Rows.Count, "A").End(xlUp).Row
    
    'cantidad de lineas de ADS
    Workbooks(nameADS).Activate
    
    Sheets("01 Entel").Select
    nEN = Sheets("01 Entel").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("02 Direcnet").Select
    nDN = Sheets("02 Direcnet").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("03 Americatel").Select
    nAM = Sheets("03 Americatel").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("04 Otros Proyectos").Select
    nOP = Sheets("04 Otros Proyectos").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("05 Postes").Select
    nPO = Sheets("05 Postes").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheets("Desmontados").Select
    nDE = Sheets("Desmontados").Cells(Rows.Count, "A").End(xlUp).Row
    
'Busqueda de NombreSite por NombreSite-------------------------------------------------------------------------------------------------------------------
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
    
    For Each cell_LM In Sheets("Sheet1").Range(Cells(2, 19), Cells(nLM, 19))
        
        'Revisa en 01 entel
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("01 Entel").Select
                    
            For Each cell_EN In Sheets("01 Entel").Range(Cells(2, 2), Cells(nEN, 2))
            
                If cell_LM.Value = cell_EN.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = cell_EN.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 02 Direcnet
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("02 Direcnet").Select
                    
            For Each cell_DN In Sheets("02 Direcnet").Range(Cells(2, 2), Cells(nDN, 2))
            
                If cell_LM.Value = cell_DN.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = cell_DN.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 03 Americatel
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("03 Americatel").Select
                    
            For Each cell_AM In Sheets("03 Americatel").Range(Cells(2, 2), Cells(nAM, 2))
            
                If cell_LM.Value = cell_AM.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = cell_AM.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 04 Otros Proyectos
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("04 Otros Proyectos").Select
                    
            For Each cell_OP In Sheets("04 Otros Proyectos").Range(Cells(2, 2), Cells(nOP, 2))
            
                If cell_LM.Value = cell_OP.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = cell_OP.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en 05 Postes
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("05 Postes").Select
                    
            For Each cell_PO In Sheets("05 Postes").Range(Cells(2, 2), Cells(nPO, 2))
            
                If cell_LM.Value = cell_PO.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = cell_PO.Offset(0, -1).Value
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
        'Revisa en Desmontados
        If cell_LM.Offset(0, 4) = "" Then
        
            'cell.Select
            Workbooks(nameADS).Activate
            Sheets("Desmontados").Select
                    
            For Each cell_DE In Sheets("Desmontados").Range(Cells(2, 2), Cells(nDE, 2))
            
                If cell_LM.Value = cell_DE.Value Then
                    
                    'cell2.Select
                    Workbooks(nameLM).Activate
                    Sheets("Sheet1").Select
                    cell_LM.Offset(0, 4) = "Desmontado"
                    Exit For
                    
                End If
                          
            Next
            
            Workbooks(nameLM).Activate
            Sheets("Sheet1").Select
            
        End If
        
    Next
            
End Sub

Sub Site_Periodo_Tipo()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    tiempo = Now
    
    nameLM = "LM ALQUILERES - JUL22.xlsx"
    nameADS = "07 Matriz 3.6 ADS_ a Jul 2022_27jul22.xlsx"
    
    'Call ID_ADS(nameLM, nameADS)
    Call Sitio_v5(nameLM, nameADS)
    Call PeriodoGasto(nameLM)
    Call TipoRenta(nameLM)
    'Call Empresa(nameLM)
    
    Workbooks(nameLM).Activate
    Sheets("Sheet1").Select
            
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
            
    tiempo = DateDiff("n", tiempo, Now)
    MsgBox ("LA MACRO HA DURADO: " & tiempo & " MINUTOS")
End Sub

