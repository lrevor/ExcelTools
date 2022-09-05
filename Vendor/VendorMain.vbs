Public vendorExports As New Dictionary
Public headings As New Dictionary

Sub loadAll()
  ' Turn on optimizations
  OptimizeCode_Begin

  ' Clear the headings cache
  headings.RemoveAll

  ' Load the data sheets
  loadSheet ("VendorExport")

  ' Process the data
  validateVendors("VendorBooks")
  
  ' Turn off the optimizations
  OptimizeCode_End

  ' End the routine
  MsgBox "Finished Loading"
End Sub


Sub refreshVendorExport()
  Dim fileName As String
  'VendorExport
  fileName = getCSVFileNameDialog("Filename for Vendor Export Report")
  Range("B8").Value = fileName
  clearOrCreateSheet ("VendorExport")
  loadFileToSheet "VendorExport", fileName
End Sub
