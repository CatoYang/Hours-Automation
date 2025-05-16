Attribute VB_Name = "TTC_Update"
Sub TTC_Update()
    Dim wsTTC As Worksheet
    Dim wsEngines As Worksheet
    Dim TTClast As Long
    Dim Engineslast As Long
    Dim i As Long
    Dim esnTTC As String
    Dim esnEngines As String
    Dim TTCData As Object
    
    'Check Sheets Naming
    If Not SheetExists("TTC") Or Not SheetExists("Engines") Then
        MsgBox "Sheets Does not Exist. Ensure there are 2 sheets, 'TTC' and 'Engines'(case sensitive). Exiting macro.", vbExclamation
        Exit Sub
    End If
    
    Set wsTTC = ThisWorkbook.Sheets("TTC")
    Set wsEngines = ThisWorkbook.Sheets("Engines")
    TTClast = wsTTC.Cells(wsTTC.Rows.Count, 1).End(xlUp).Row
    Engineslast = wsEngines.Cells(wsEngines.Rows.Count, 2).End(xlUp).Row
    
    ' Create Dictionary for TTC data
    Set TTCData = CreateObject("Scripting.Dictionary")
    
    ' Load TTC data into Dictionary
    For j = 2 To TTClast
        esnTTC = wsTTC.Cells(j, 1).Value
        If TTCData.exists(esnTTC) Then
            ' Compare EOT and keep the higher one
            If wsTTC.Cells(j, 17).Value > TTCData(esnTTC)(1, 18) Then
                TTCData(esnTTC) = wsTTC.Rows(j).Value
            End If
        Else
            ' Add new key if it doesn't exist
            TTCData.Add esnTTC, wsTTC.Rows(j).Value
        End If
    Next j
    
    'Reset Fill colour
    wsEngines.Range("B2:AV117").Interior.ColorIndex = xlNone
    
    For i = 2 To Engineslast
        esnEngines = wsEngines.Cells(i, 2).Value
        If TTCData.exists(esnEngines) Then
            Dim TTC As Variant
            TTC = TTCData(esnEngines)
            
            ' Column 3 = Download time , 18 = COT
            If TTC(1, 2) > wsEngines.Cells(i, 3).Value And _
               TTC(1, 17) > wsEngines.Cells(i, 18).Value Then
                
                ' Overwrite columns 3 to 48
                Dim col As Long
                For col = 3 To 48
                    wsEngines.Cells(i, col).Value = TTC(1, col - 1)
                    wsEngines.Cells(i, col).Interior.Color = RGB(51, 204, 51)
                Next col
                
                'Reset Time formatting (older excel versions/TTC data mess with the formatting)
                wsEngines.Cells(i, 3).NumberFormat = "dd/mm/yyyy hh:mm"
                wsEngines.Cells(i, 49).Value = Now
                wsEngines.Cells(i, 49).NumberFormat = "dd/mm/yyyy hh:mm"
            End If
        End If
    Next i
    
    ThisWorkbook.Sheets("Daily_Hr").Range("F4").Value = Now
    ThisWorkbook.Sheets("Daily_Hr").Range("F4").NumberFormat = "dd/mm/yyyy"
    
    Application.CutCopyMode = False
    MsgBox "Engine Data Updated", vbInformation
End Sub



