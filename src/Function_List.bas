Attribute VB_Name = "Function_List"
'This Function is used in several modules to test for sheets name before execution
Function SheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
'This sub is for copying fill colour from Update list to Daily_hr
Sub SyncFillColors()
    Dim wsA As Worksheet, wsB As Worksheet
    Dim rngA As Range, rngB As Range
    Dim i As Long

    Set wsA = ThisWorkbook.Sheets("Daily_Hr")
    Set wsB = ThisWorkbook.Sheets("Update List")

    Set rngA = wsA.Range("C8:C47")
    Set rngB = wsB.Range("C2:C41")
    
    rngA.Interior.ColorIndex = xlNone
    
    For i = 1 To rngA.Rows.Count
        rngA.Cells(i, 1).Interior.Color = rngB.Cells(i, 1).Interior.Color
    Next i
End Sub
