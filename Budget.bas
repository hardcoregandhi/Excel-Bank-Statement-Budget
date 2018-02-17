Attribute VB_Name = "Module1"
Sub Budget()
Attribute Budget.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' Line 12 > Categories
' Line 65 > Parsing
'

'
If WorksheetFunction.CountA(Range("J2:L13")) <> 0 And Range("J2") <> "Food" Then
    MsgBox "'J2', 'L13' are not Empty"
    Exit Sub
End If

    Range("J2").Select
        ActiveCell.Formula = "Food"
        Range("K2").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Food"",F:F)"
    Range("J3").Select
        ActiveCell.Formula = "Amazon"
        Range("K3").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Amazon"",F:F)"
    Range("J4").Select
        ActiveCell.Formula = "Fun"
        Range("K4").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Fun"",F:F)"
    Range("J5").Select
        ActiveCell.Formula = "Utility"
        Range("K5").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Utility"",F:F)"
    Range("J6").Select
        ActiveCell.Formula = "Takeaway"
        Range("K6").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Takeaway"",F:F)"
    Range("J7").Select
        ActiveCell.Formula = "Subscriptions"
        Range("K7").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Subscriptions"",F:F)"
    Range("J8").Select
        ActiveCell.Formula = "Misc"
        Range("K8").Select
        ActiveCell.Formula = "=SUMIF(D:D,""Misc"",F:F)"
    
    Range("J12").Select
    ActiveCell.Value = "Outgoing"
    Range("J13").Select
    ActiveCell.Value = "=SUM(F:F)"
    Range("K12").Select
    ActiveCell.FormulaR1C1 = "Incoming"
    Range("K13").Select
    ActiveCell.Value = "=SUM(G:G)"
    Range("L12").Select
    ActiveCell.FormulaR1C1 = "Net"
    Range("L13").Select
    ActiveCell.Formula = "=K13-J13"
    
    Range("D1").Select
    If ActiveCell.Value <> "Category" Then
        Columns("D:D").Select
        Selection.ClearContents
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Category"
    End If
    
    For Each c In ActiveSheet.UsedRange.Columns("D").Cells
        c.Select
        If IsEmpty(c) = True Then
            Select Case ActiveCell.Offset(0, 1).Value
                Case "PLACEYOUGETFOODFROM", "OTHERPLACEYOUGETFOODFROM"
                    ActiveCell.FormulaR1C1 = "Food"
                Case "MYGASSUPPLIER", "MYELECPROVIDER", "MYINTERNETPROVIDER", "MYGASANDELECPROVIDER", "MYESTATEAGENT", "MYMOBILEPLAN", "COUNCILTAX"
                    ActiveCell.FormulaR1C1 = "Utility"
                Case "YOURWORKPLACE", "INTEREST (GROSS)"
                    ActiveCell.FormulaR1C1 = "Income"
                Case "SUBSCRIPTION1", "SUBSCRIPTION2", "SUBSCRIPTION3", "SUBSCRIPTION4", "SUBSCRIPTION5", "SUBSCRIPTION6", "SUBSCRIPTION7", "SUBSCRIPTION8"
                    ActiveCell.FormulaR1C1 = "Subscriptions"
                Case "THINGIDOEVEERYMONTH1", "THINGIDOEVERYMONTH2"
                    ActiveCell.FormulaR1C1 = "Fun"
                Case Else
                    If InStr(ActiveCell.Offset(0, 1).Value, "PAYPAL") Then
                        ActiveCell.FormulaR1C1 = "Fun"
                    End If
                    If InStr(ActiveCell.Offset(0, 1).Value, "AMAZON") Or InStr(ActiveCell.Offset(0, 1).Value, "Amazon") Then
                        ActiveCell.FormulaR1C1 = "Amazon"
                    End If
                    If InStr(ActiveCell.Offset(0, 1).Value, "UBER EATS") Or InStr(ActiveCell.Offset(0, 1).Value, "JUST EAT") Or InStr(ActiveCell.Offset(0, 1).Value, "PIZZA HUT") Then
                        ActiveCell.FormulaR1C1 = "Takeaway"
                    End If
                    If InStr(ActiveCell.Offset(0, 1).Value, "Spotify") Then
                        ActiveCell.FormulaR1C1 = "Subscriptions"
                    End If
                     
            End Select
        End If
    Next
    
    
    
    
    Range("A1").Select
    Cells.Columns.AutoFit
End Sub

