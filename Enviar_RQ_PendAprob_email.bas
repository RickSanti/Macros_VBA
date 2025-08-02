Attribute VB_Name = "Enviar_RQ_PendAprob_email"
Sub EnviarEmail()

'
' Declaramos variables
'
Dim OutlookApp As Outlook.Application
Dim MItem As Outlook.MailItem
Dim cell As Range
Dim cell2 As Range
Dim Asunto As String
Dim Correo As String
Dim Destinatario As String
Dim Saldo As String
Dim Msg As String
Dim n As Double
    '
    Sheets("Sheet1").Range("A1").Select
    Set OutlookApp = New Outlook.Application
    '
    ' Recorremos la columna EMAIL
    '
    n = Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
    
    For Each cell In Sheets("Correos").Range("B17:B23")

' UNO------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If cell.Value = "nombre@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - MG") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - Controller", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
            If Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - MG", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - Controller", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Then
                Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
                        
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "VP Redes - MG" And cell2.Offset(0, -3).Value = "CAPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                        
                    ElseIf cell2.Value = "VP Redes - Controller" And cell2.Offset(0, -3).Value = "CAPEX" Then

                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine

                    End If
                    
                Next
                
            End If
                                    
            If Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - MG", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Then
                Msg = Msg & vbNewLine
                Msg = Msg & "OPEX" & vbNewLine & vbNewLine
                        
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "VP Redes - MG" And cell2.Offset(0, -3).Value = "OPEX" Then
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                    End If
                    
                Next
            End If
            
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
 'DOS------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre2@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Finanzas", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Capital Humano", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value

            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine

            ' Msg = Msg & "OPEX" & vbNewLine & vbNewLine

            If (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Finanzas", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Capital Humano", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0) Then
            
                ' Msg = Msg & vbNewLine
                Msg = Msg & "OPEX" & vbNewLine & vbNewLine
                
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "GC Finanzas" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                                            
                    ElseIf cell2.Value = "GC Capital Humano" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                        
                    End If
                    
                Next
            End If

            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
' TRES------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre3@compania.pe" And Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - Controller", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
'            If Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - Controller", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Then
'
'                Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
'
'                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
'
'                    If cell2.Value = "VP Redes - Controller" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    End If
'
'                Next
'
'            End If
            
            If Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP Redes - Controller", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Then
                
                ' Msg = Msg & vbNewLine
                Msg = Msg & "OPEX" & vbNewLine & vbNewLine
                
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "VP Redes - Controller" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                        
                    End If
                    
                Next
                
            End If
            
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
            
' CUATRO------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Logistica", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Facturación", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
            Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
            
            For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                
                If cell2.Value = "VP TI" And cell2.Offset(0, -3).Value = "CAPEX" Then
                    
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                
                ElseIf cell2.Value = "VP TI - Logistica" And cell2.Offset(0, -3).Value = "CAPEX" Then
                
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                
                ElseIf cell2.Value = "VP TI - Facturación" And cell2.Offset(0, -3).Value = "CAPEX" Then
                
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                    
                End If
                
            Next
                                   
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
        
' CINCO------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre1@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Capital Humano", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Finanzas", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Planeamiento", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
            Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
            
            For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                
                If cell2.Value = "GC Capital Humano" And cell2.Offset(0, -3).Value = "CAPEX" Then
                    
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                    
                ElseIf cell2.Value = "GC Finanzas" And cell2.Offset(0, -3).Value = "CAPEX" Then
                    
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                    
                ElseIf cell2.Value = "GC Planeamiento" And cell2.Offset(0, -3).Value = "CAPEX" Then
                    
                    Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                                         
                End If
                
            Next
                                   
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                '.CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
' SEIS------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre4@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Logistica", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
'            If (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Logistica", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Finanzas", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0) Then
'
'                Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
'
'                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
'
'                    If cell2.Value = "VP TI - Logistica" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    ElseIf cell2.Value = "GC Finanzas" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    End If
'
'                Next
'
'            End If
            
            If (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Logistica", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0) Then
            
                ' Msg = Msg & vbNewLine
                Msg = Msg & "OPEX" & vbNewLine & vbNewLine
                
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "VP TI - Logistica" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                                                            
                    '''
                    ElseIf cell2.Value = "VP TI" And cell2.Offset(0, -3).Value = "OPEX" Then
    
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                    '''
                    End If
                    
                Next
            End If
            
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
        
' SIETE------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf cell.Value = "nombre5@compania.pe" And (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Facturación") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Planeamiento") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Legal") <> 0) Then
            '
            'Asignamos valor a las variables
            '
            Asunto = cell.Offset(0, 3).Value
            Destinatario = cell.Offset(0, 2).Value
            Correo = cell.Value
            
            '
            'Cuerpo del mensaje
            '
            Msg = "Buenas " & Destinatario & vbNewLine & vbNewLine
            Msg = Msg & "Un apoyo indicándonos si valido / rechazo las siguientes RQs pendiente de aprobación. " & vbNewLine & vbNewLine
            
'            If (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Facturación", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Planeamiento", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Legal", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "CAPEX") <> 0) Then
'                Msg = Msg & "CAPEX" & vbNewLine & vbNewLine
'
'                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
'
'                    If cell2.Value = "VP TI - Facturación" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    ElseIf cell2.Value = "GC Planeamiento" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    ElseIf cell2.Value = "GC Legal" And cell2.Offset(0, -3).Value = "CAPEX" Then
'
'                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
'
'                    End If
'
'                Next
'            End If
            
            If (Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "VP TI - Facturación", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Planeamiento", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0 Or Application.WorksheetFunction.CountIfs(Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29)), "GC Legal", Sheets("Sheet1").Range(Cells(2, 26), Cells(n, 26)), "OPEX") <> 0) Then
                
                ' Msg = Msg & vbNewLine
                Msg = Msg & "OPEX" & vbNewLine & vbNewLine
                
                For Each cell2 In Sheets("Sheet1").Range(Cells(2, 29), Cells(n, 29))
                    
                    If cell2.Value = "VP TI - Facturación" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                                                            
                    ElseIf cell2.Value = "GC Planeamiento" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                                                            
                    ElseIf cell2.Value = "GC Legal" And cell2.Offset(0, -3).Value = "OPEX" Then
                        
                        Msg = Msg & cell2.Offset(0, -28).Value & vbNewLine
                        
                    End If
                
                Next
            
            End If
            
            Msg = Msg & vbNewLine & "Estamos adjuntado la base de RQs para cualquier consulta." & vbNewLine
            Msg = Msg & vbNewLine & "https://RUTA" & vbNewLine
            Msg = Msg & vbNewLine & "Saludos," & vbNewLine
            Msg = Msg & "Ricardo."
            '
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .To = Correo
                .CC = "nombre1@compania.pe"
                .Subject = Asunto
                .Body = Msg
                '.Attachments.Add (Sheets("Correos").Range("B12").Value)
                .Send
                '
            End With
            '
                
        End If
        
    Next
    
    Sheets("Sheet1").Range("A1").Select
End Sub
