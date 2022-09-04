Sub createMBReport()
  Dim mbKey As Variant
  Dim mbname As String
  Dim scoutKey As Variant
  Dim myScout As scoutRecord
  Dim i As Integer
  Dim j As Integer
  Dim headerRows As Integer
  Dim rng As Range
  
  Dim returnState As ExcelState
  Set returnState = New ExcelState
  
  clearOrCreateSheet ("MeritBadges")
  activeWorkbook.Sheets("MeritBadges").Visible = xlSheetVisible
  activeWorkbook.Sheets("MeritBadges").Select
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
  
  For Each mbKey In eaglembnames.Keys
    mbname = eaglembnames(mbKey)
    i = i + 1
    ActiveSheet.Cells(1, i).Value = mbname
  Next
  For Each mbKey In mbnames.Keys
    If Not (eaglembnames.Exists(mbKey)) Then
      mbname = mbnames(mbKey)
      i = i + 1
      ActiveSheet.Cells(1, i).Value = mbname
    End If
  Next
  
  'Build the Repot
  Set rng = ActiveSheet.UsedRange
  j = 2
  For Each scoutKey In scouts.Keys
    Set myScout = scouts(scoutKey)
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "BSA Member ID")).Value = myScout.id.BSAMemberID
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Last Name")).Value = myScout.id.LastName
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "First Name")).Value = myScout.id.FirstName
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Status")).Value = myScout.id.Archive
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "DOB")).Value = myScout.idDetail.DOB
    For Each mbKey In mbnames.Keys
      mbname = mbnames(mbKey)
      If Not (myScout.MBDate(mbname) = "") Then
        ActiveSheet.Cells(j, getColumnFromHeader(rng, mbname)).Value = "x"
      End If
    Next
    j = j + 1
  Next
  'ActiveSheet.Cells(1, 1).Interior.ColorIndex = 5
  'ActiveSheet.Range(Columns(1), Columns(4)).Interior.ColorIndex = 5
  'ActiveSheet.Range(Columns(5), Columns(5 + numEagleReq - 1)).Interior.ColorIndex = 10
  'ActiveSheet.Range(Columns(5 + numEagleReq), Columns(5 + numEagleReq + numNonEagleReq - 1)).Interior.ColorIndex = 15
  
  ActiveSheet.Range("1:1").Orientation = xlUpward
  ActiveSheet.Range("1:1").HorizontalAlignment = xlCenter
  ActiveSheet.UsedRange.Columns.AutoFit
  returnState.restore
End Sub
