Private pactiveWorkbook As Workbook
Private pactiveSheet As String
    
Public Sub restore()
  Application.Goto Range("A1"), Scroll:=True
  pactiveWorkbook.Activate
  pactiveWorkbook.Sheets(pactiveSheet).Select
End Sub
Private Sub Class_Initialize()
  Set pactiveWorkbook = Application.activeWorkbook
  pactiveSheet = Application.ActiveSheet.Name
End Sub

