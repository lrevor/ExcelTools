Sub createTemplateReport()
  Dim scoutKey As Variant
  Dim myScout As scoutRecord
  Dim i As Integer
  Dim j As Integer
  Dim rng As Range
  
  Dim returnState As ExcelState
  Set returnState = New ExcelState
 
  clearOrCreateSheet ("Template")
  activeWorkbook.Sheets("Template").Visible = xlSheetVisible
  activeWorkbook.Sheets("Template").Select
  ActiveSheet.UsedRange.Interior.ColorIndex = 0
  
  'Set up header rows.  Dont forget to set the value below
  i = 1
  ActiveSheet.Cells(1, i).Value = "BSA Member ID"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Last Name"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "First Name"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Status"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "DOB"
  
  ActiveWindow.FreezePanes = False
  ActiveSheet.Cells(2, i + 1).Select
  ActiveWindow.FreezePanes = True

  'Add code for additional Titles here
  
  'Start the report
  Set rng = ActiveSheet.UsedRange
  j = 1
  For Each scoutKey In scouts.Keys
    j = j + 1
    Set myScout = scouts(scoutKey)
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "BSA Member ID")).Value = myScout.id.BSAMemberID
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Last Name")).Value = myScout.id.LastName
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "First Name")).Value = myScout.id.FirstName
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Status")).Value = myScout.id.Archive
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "DOB")).Value = myScout.idDetail.DOB
  
    'Add Code for additional Title Data Here
  Next
  
  ActiveSheet.Range("1:1").Orientation = xlUpward
  ActiveSheet.Range("1:1").HorizontalAlignment = xlCenter
  ActiveSheet.UsedRange.Columns.AutoFit
  returnState.restore
End Sub

