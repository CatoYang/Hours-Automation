VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A1")) Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!TTC_Update.TTC_Update"
    End If
End Sub


