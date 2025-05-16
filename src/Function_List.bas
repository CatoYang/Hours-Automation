Attribute VB_Name = "Function_List"
'This Function is used in several modules to test for sheets name before execution
Function SheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

