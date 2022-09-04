Sub loadFileToSheet(strSheet As String, strFileName As String)
    Dim targetWorkbook As Workbook
    Dim activeWorkbook As Workbook
    Dim myactiveSheet As Worksheet
    
    Set activeWorkbook = Application.activeWorkbook
    Set myactiveSheet = Application.ActiveSheet

    Set targetWorkbook = Application.Workbooks.Open(strFileName)
    targetWorkbook.Activate
    'MsgBox "The current sheet is " & targetWorkbook.activeSheet.Name
    ActiveSheet.UsedRange.Copy (activeWorkbook.Sheets(strSheet).Range("A1"))
    
    targetWorkbook.Close (SaveChanges = False)
    
    activeWorkbook.Activate
    activeWorkbook.Sheets(myactiveSheet.Name).Select
End Sub

Sub clearOrCreateSheet(strName As String)
    Dim mySheetNameTest As String
    Dim returnSheet As String
    
    returnSheet = ActiveSheet.Name
    
    On Error Resume Next
    activeWorkbook.Sheets(strName).Visible = xlSheetVisible
    activeWorkbook.Sheets(strName).Select
    If Err.Number = 0 Then
        activeWorkbook.Sheets(strName).UsedRange.Clear
    Else
        Err.Clear
        Worksheets.Add.Name = strName
        activeWorkbook.Sheets(strName).Move after:=Worksheets(Worksheets.Count)
    End If
    activeWorkbook.Sheets(strName).Visible = xlSheetHidden
    activeWorkbook.Sheets(returnSheet).Select
End Sub
