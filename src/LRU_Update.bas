Attribute VB_Name = "LRU_Update"
Sub LRU_Update()
    Dim PSU As Worksheet
    Dim LRU As Worksheet
    Dim PSUlast As Long, LRUlast As Long
    Dim i As Long, j As Long
    Dim PSUserial As String, LRUserial As String

    ' Check Sheets Naming
    If Not SheetExists("PSU") Or Not SheetExists("LRU") Then
        MsgBox "Sheets do not exist. Ensure there are 2 sheets, 'PSU' and 'LRU' (case sensitive). Exiting macro.", vbExclamation
        Exit Sub
    End If

    Set PSU = ThisWorkbook.Sheets("PSU")
    Set LRU = ThisWorkbook.Sheets("LRU")
    PSUlast = PSU.Cells(PSU.Rows.Count, 3).End(xlUp).Row
    LRUlast = LRU.Cells(LRU.Rows.Count, 4).End(xlUp).Row

    Dim PSUData As Object
    Set PSUData = CreateObject("Scripting.Dictionary")
    Set LRUData = CreateObject("Scripting.Dictionary")
    
    ' Populate dictionary with PSU row numbers
    For i = 2 To PSUlast
        PSUserial = PSU.Cells(i, 3).Value
        If Not PSUData.exists(PSUserial) And PSUserial <> "" Then
            PSUData(PSUserial) = i
        End If
    Next i
    
    ' Populate dictionary with LRU serials
    For j = 2 To LRUlast
        LRUserial = LRU.Cells(j, 4).Value
        If Not LRUData.exists(LRUserial) And LRUserial <> "" Then
            LRUData(LRUserial) = True
        End If
    Next j
    
    ' Reset fill highlights in designated columns
    LRU.Range("B:V").Interior.ColorIndex = xlNone
    LRU.Range("E2:E" & LRUlast & ",K2:K" & LRUlast).Interior.Color = RGB(255, 255, 0)

    For j = 2 To LRUlast
        LRUserial = LRU.Cells(j, 4).Value
        If PSUData.exists(LRUserial) Then
            Dim PSURow As Long
            PSURow = PSUData(LRUserial)
            If PSU.Cells(PSURow, 10).Value > LRU.Cells(j, 11).Value Then
                LRU.Rows(j).Columns("B:V").Value = PSU.Rows(PSURow).Columns("A:U").Value
                LRU.Rows(j).Columns("B:V").Interior.Color = RGB(51, 204, 51)
                LRU.Cells(j, 5).Interior.Color = RGB(255, 255, 0)   ' Column E
                LRU.Cells(j, 11).Interior.Color = RGB(255, 255, 0)  ' Column K
                LRU.Cells(j, 23).Value = Now
                LRU.Cells(j, 23).NumberFormat = "dd/mm/yyyy hh:mm"
                LRU.Cells(j, 25).Value = LRU.Cells(j, 26).Value
            End If
        End If
    Next j
    
    ' Clear previous highlights in PSU
    PSU.Range("A2:Z" & PSUlast).Interior.ColorIndex = xlNone
    ' Highlight unmatched serial numbers in PSU
    For i = 2 To PSUlast
        PSUserial = PSU.Cells(i, 3).Value
        If PSUserial <> "" And Not LRUData.exists(PSUserial) Then
            PSU.Cells(i, 3).Interior.Color = RGB(255, 199, 206)
        End If
    Next i

    
    ThisWorkbook.Sheets("Daily_Hr").Range("F5").Value = Now
    ThisWorkbook.Sheets("Daily_Hr").Range("F5").NumberFormat = "dd/mm/yyyy"

    MsgBox "LRU Data Reconciled", vbInformation
End Sub




