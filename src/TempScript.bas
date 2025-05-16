Attribute VB_Name = "TempScript"
Sub Hours_sort()
    Dim Due As Worksheet
    Dim Update As Worksheet
    Dim Duelast As Long
    Dim Updatelast As Long
    Dim i As Long, j As Long
    Dim esnDue As String
    Dim esnUpdate As String
    Dim Duehours As Object
    Dim col As Long
    
    ' Set worksheets
    Set Due = ThisWorkbook.Sheets("Due")
    Set Update = ThisWorkbook.Sheets("LRU")
    
    ' Find the last rows in Due and Update sheets
    Duelast = Due.Cells(Due.Rows.Count, 1).End(xlUp).Row
    Updatelast = Update.Cells(Update.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize dictionary
    Set Duehours = CreateObject("Scripting.Dictionary")
    
    ' Populate dictionary with "Due" data
    For j = 2 To Duelast
        esnDue = Due.Cells(j, 1).Value
        If Not Duehours.exists(esnDue) Then
            ' Store row data in the dictionary
            Duehours.Add esnDue, Due.Cells(j, 2).Value
        End If
    Next j
    
    ' Define the column in Update sheet to write the data
    col = 5 ' Example column, update as needed
    
    ' Compare and update data
    For i = 2 To Updatelast
        esnUpdate = Update.Cells(i, 1).Value
        If Duehours.exists(esnUpdate) Then
            ' Write the corresponding value from the dictionary
            Update.Cells(i, col).Value = Duehours(esnUpdate)
        End If
    Next i
    
    ' Clean up
    Set Duehours = Nothing
    
    MsgBox "Sort Completed", vbInformation
End Sub

