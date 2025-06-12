Attribute VB_Name = "AFH_Update"
Sub AFH_Update()
    Dim AFH As Worksheet
    Dim oldAFH As Double
    Dim newAFH As Double
    Dim i As Long
    
    'Check Sheet Naming
    If Not SheetExists("Update List") Then
        MsgBox "For AFH Update. Ensure the associated Sheet is named 'Update List' (case sensitive). Exiting macro.", vbExclamation
        Exit Sub
    End If
    
    Set AFH = ThisWorkbook.Sheets("Update List")
    
    'Reset colour scheme and repaint cell 6
    AFH.Range(AFH.Cells(2, 2), AFH.Cells(80, 6)).Interior.ColorIndex = xlNone
    AFH.Range(AFH.Cells(2, 6), AFH.Cells(80, 6)).Interior.Color = RGB(255, 255, 0)
    
    'We actually only use 2 to 40, the rest of the rows are future contingency
    For i = 2 To 80
        newAFH = AFH.Cells(i, 6).Value
        oldAFH = AFH.Cells(i, 4).Value
            If newAFH > oldAFH Then
                AFH.Cells(i, 4) = newAFH
                AFH.Cells(i, 5).Value = Now
                AFH.Cells(i, 5).NumberFormat = "dd/mm/yyyy hh:mm"
                'Colours the Cell Green to show an increase
                AFH.Range(AFH.Cells(i, 2), AFH.Cells(i, 6)).Interior.Color = RGB(146, 208, 80)
                If (newAFH - oldAFH) > 6 Then
                    'Colours the Cell Purple to show Anomalous increase
                    AFH.Range(AFH.Cells(i, 2), AFH.Cells(i, 6)).Interior.Color = RGB(225, 153, 225)
                    MsgBox "Warning: Anomalous input noted (AFH Increase is Greater than 6.00). AFH overwriten, check value before saving"
                End If
            ElseIf newAFH < oldAFH Then
                'Colours the Cell Red to show an Erronous input, a decrease in value
                AFH.Range(AFH.Cells(i, 2), AFH.Cells(i, 6)).Interior.Color = RGB(255, 0, 0)
                MsgBox "Warning: Lower AFH Input noted. AFH not overwritten. Check red shaded boxes"
            End If
    Next i
    
    SyncFillColors
    
    ThisWorkbook.Sheets("Daily_Hr").Range("F3").Value = Now
    ThisWorkbook.Sheets("Daily_Hr").Range("F3").NumberFormat = "dd/mm/yyyy"
    MsgBox "AFH Data Updated", vbInformation
End Sub

