Attribute VB_Name = "Reporte_Demora_RQ_v2"
Private Sub RQ_unicos()

    'Declaramos variables
    Dim nRQ As Long, i As Long, RQ_un As Variant
    
    Sheets("RQ").Select
    nRQ = Sheets("RQ").Cells(Rows.Count, "B").End(xlUp).Row
    
    RQ_un = WorksheetFunction.Unique(Range("B2:B" & nRQ))
    'RQ_un = WorksheetFunction.Unique(Range(Cells(2, 2), Cells(nRQ, 2)))
    
    Sheets("BASE").Select
    For i = LBound(RQ_un) To UBound(RQ_un)

        Cells(i + 1, 1) = RQ_un(i, 1)
        
    Next
                
End Sub

Private Sub RQ_data(ByVal nBASE As Long)
Attribute RQ_data.VB_ProcData.VB_Invoke_Func = " \n14"

'DATOS--------------------------------------------------

    'Se pone la formula
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC1,RQ!C2:C42,MATCH(R1C,RQ!R1C2:R1C42,0),0)=0,"""",VLOOKUP(RC1,RQ!C2:C42,MATCH(R1C,RQ!R1C2:R1C42,0),0))"
        
    'Copiamos la celda con la formula
    Range("B2").Select
    Selection.Copy
    
    'Lo pegamos en el rango correspondiente
    Range("B2:P" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
'MONTOS---------------------------------------------------
    
    'Se pone la formula
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(RQ!C[11],RQ!C2,BASE!RC1)"
    
    'Copiamos la celda con la formula
    Range("I2").Select
    Selection.Copy
    
    'Lo pegamos en el rango correspondiente
    Range("I2:J" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("I2").Select
    
End Sub

Private Sub OC_data(ByVal nBASE As Long)
    
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC8,OC!C3:C43,MATCH(R1C,OC!R1C3:R1C43,0),0),"""")=0,"""",IFERROR(VLOOKUP(RC8,OC!C3:C43,MATCH(R1C,OC!R1C3:R1C43,0),0),""""))"
        
    Range("Q2").Select
    Selection.Copy
    
    Range("Q2:S" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub

Private Sub Acuerdo_data(ByVal nBASE As Long)
    
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-5]="""",RC[-5]=0),""Sin Acuerdo"",""Con Acuerdo"")"
        
    Range("T2").Select
    Selection.Copy
    
    Range("T2:T" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub

Private Sub Dias_data(ByVal nBASE As Long)
    
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-18]="""",RC[-18]=0),0,NETWORKDAYS.INTL(RC[-19],RC[-18],1))"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-4]="""",RC[-4]=0),0,NETWORKDAYS.INTL(RC[-5],RC[-4],1))"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-5]="""",RC[-5]=0),0,NETWORKDAYS.INTL(RC[-21],RC[-5],1))"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-7]="""",RC[-7]=0),0,NETWORKDAYS.INTL(RC[-21],RC[-7],1))"
        
    Range("U2:X2").Select
    Selection.Copy
    
    Range("U2:X" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
End Sub

Private Sub Condiciones_data(ByVal nBASE As Long)
    
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-23])"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[-24])"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC10<70000,""Hasta S/ 70,000"",IF(RC10>700000,""Mayor a S/ 700,000"",""De S/ 70,001 a S/ 700,000""))"
    Range("AB2").Select
    ActiveCell.Formula2R1C1 = _
        "=IFS(AND(RC[-12]=""CAPEX"",RC[-8]=""Con Acuerdo""),VLOOKUP(RC[-1],'Estimación Tiempo'!R4C4:R6C18,15,0),AND(RC[-12]=""OPEX"",RC[-8]=""Con Acuerdo""),VLOOKUP(RC[-1],'Estimación Tiempo'!R8C4:R10C18,15,0),AND(RC[-12]=""CAPEX"",RC[-8]=""Sin Acuerdo""),VLOOKUP(RC[-1],'Estimación Tiempo'!R12C4:R14C18,15,0),AND(RC[-12]=""OPEX"",RC[-8]=""Sin Acuerdo""),VLOOKUP(RC[-1],'Estima" & _
        "ción Tiempo'!R16C4:R18C18,15,0))" & _
        ""
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]-RC[-1]"
    
    Range("Y2:AC2").Select
    Selection.Copy
    
    Range("Y2:AC" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    
End Sub

Private Sub Segmentacion_data(ByVal nBASE As Long)
    
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC1,Niveles!C1:C4,2,0),"""")"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC1,Niveles!C1:C4,4,0),"""")"
    
    Range("AD2:AE2").Select
    Selection.Copy
    
    Range("AD2:AE" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
End Sub

Private Sub Status_data(ByVal nBASE As Long)

    Range("AF2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC4=""Pendiente Aprobación"",IFERROR(VLOOKUP(RC1,ReporteRequisiciones!C2:C7,5,0),""""),"""")"
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC4=""Pendiente Aprobación"",IFERROR(VLOOKUP(RC1,ReporteRequisiciones!C2:C7,6,0),""""),"""")"
    
    Range("AF2:AG2").Select
    Selection.Copy
    Range("AF2:AG" & nBASE).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
End Sub

Private Sub EnValores()

    Columns("B:P").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("H:H").Select
    Selection.TextToColumns Destination:=Range("H1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
        
    Columns("P:AG").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub BASE_Demora_RQs()
Attribute BASE_Demora_RQs.VB_ProcData.VB_Invoke_Func = " \n14"
    
    tiempo = Now
    
    Dim nBASE As Long
    
    Call RQ_unicos
    
    Sheets("BASE").Select
    nBASE = Sheets("BASE").Cells(Rows.Count, "A").End(xlUp).Row
            
    Call RQ_data(nBASE)
    Call OC_data(nBASE)
    Call Acuerdo_data(nBASE)
    Call Dias_data(nBASE)
    Call Condiciones_data(nBASE)
    Call Segmentacion_data(nBASE)
    Call Status_data(nBASE)
    
    Call EnValores
    
    Range("A1").Select
    Application.CutCopyMode = False

    tiempo = DateDiff("n", tiempo, Now)
    MsgBox ("LA MACRO HA DURADO: " & tiempo & " MINUTOS")
    
End Sub
