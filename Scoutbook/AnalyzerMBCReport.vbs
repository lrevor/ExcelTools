Sub createMBCReport()
  Dim mbc As Variant
  Dim mbname As Variant
  Dim i As Integer
  Dim j As Integer
  
  Dim returnState As ExcelState
  Set returnState = New ExcelState
 
  clearOrCreateSheet ("MBCReport")
  activeWorkbook.Sheets("MBCReport").Visible = xlSheetVisible
  activeWorkbook.Sheets("MBCReport").Select
  ActiveSheet.UsedRange.Interior.ColorIndex = 0
  
  ActiveSheet.Cells(1, 1).Value = "MBC Name"
  
  ActiveWindow.FreezePanes = False
  ActiveSheet.Cells(2, 2).Select
  ActiveWindow.FreezePanes = True
  
  'Add code for additional Titles here
  i = 1
  For Each mbname In mbcmbnames.Keys
    i = i + 1
    ActiveSheet.Cells(i, 1).Value = mbname
  Next mbname
  'Start the report
  j = 1
  For Each mbc In mbcs.Keys
    j = j + 1
    ActiveSheet.Cells(1, j).Value = mbc
    i = 1
    For Each mbname In mbcmbnames.Keys
      i = i + 1
      If InStr(mbcs(mbc), mbname) > 0 Then
        ActiveSheet.Cells(i, j).Value = "x"
      End If
    Next mbname

    'Add Code for additional Title Data Here
  Next mbc
  
  ActiveSheet.Range("1:1").Orientation = xlUpward
  ActiveSheet.Range("1:1").HorizontalAlignment = xlCenter
  ActiveSheet.UsedRange.Columns.AutoFit
  returnState.restore
End Sub


