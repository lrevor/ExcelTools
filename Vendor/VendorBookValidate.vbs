' This is a helper routine
Sub validateUpdateField(rng As Range, i As Integer, upc As String, field As String, value As String, isNumber As Boolean)
  Dim j As Integer, data as String, boxTitle As String, boxData as String, cValue As Currency, cData As Currency
  
  j = getColumnFromHeader(rng, field)
  data = rng.Cells(i, j)
  boxTitle = "Update " & field & " from VendorExport for Case UPC (" & upc & ")?"

  If (isNumber) Then
    cValue = value
    cData = data
    boxData = "VendorBooks: " & cData & " VendorExport: " & cValue
    If (cValue <> cData) Then
      If (MsgBox(boxData, vbYesNo, boxTitle) = vbYes) Then
        Debug.Print "Yes Selected"
        rng.Cells(i, j) = cValue
      End If
    End If
  Else
    boxData = "VendorBooks: " & data & " VendorExport: " & value
    If (value <> data) Then
      If (MsgBox(boxData, vbYesNo, boxTitle) = vbYes) Then
        Debug.Print "Yes Selected"
        rng.Cells(i, j) = value
      End If
    End If
  End If
End Sub

' This is the main processing routine
Sub validateVendors(mySheet As String)

  Dim rows As Integer, cols As Integer, i As Integer, j As Integer, rng As Range, data As String, found As Boolean, result As Integer

  ' Initializae and sheet specific records
  Dim vendorKey as Variant
  Dim myVendor As VendorRecord
    
  ' Save the current state so it can be put back at the end of the call
  Dim returnState As ExcelState
  Set returnState = New ExcelState
  
  Sheets(mySheet).Activate

  Set rng = ActiveSheet.UsedRange

  cols = rng.Columns.Count
  rows = rng.rows.Count

  ' Add Processing Here
  For Each vendorKey In vendorExports.Keys
    Dim CaseUPC As String, boxData As String, boxTitle As String
    Set myVendor = vendorExports.Item(vendorKey)
    found = False
    For i = 2 To rows
     j = getColumnFromHeader(rng, "Case UPC")
     data = rng.Cells(i, j)
     CaseUPC = data
     
     If (myVendor.CaseID = data) Then
        'Debug.Print "Found Case UPC: " & data
        found = True

        'Validate the Data
        validateUpdateField rng, i, CaseUPC, "Supplier #", myVendor.SupplierNumber, False
        validateUpdateField rng, i, CaseUPC, "Supplier Name", myVendor.SupplierName, False
        validateUpdateField rng, i, CaseUPC, "Cost Link", myVendor.CostLink, False
        validateUpdateField rng, i, CaseUPC, "Retail Link", myVendor.RetailLink, False
        validateUpdateField rng, i, CaseUPC, "HEB IC", myVendor.EntityID, False
        validateUpdateField rng, i, CaseUPC, "UPC", myVendor.UPC, False
        validateUpdateField rng, i, CaseUPC, "MFC", myVendor.MFC, False
        validateUpdateField rng, i, CaseUPC, "Description", myVendor.Description, False
        validateUpdateField rng, i, CaseUPC, "Master Pack", myVendor.MasterPack, False
        validateUpdateField rng, i, CaseUPC, "Pack", myVendor.Pack, False
        validateUpdateField rng, i, CaseUPC, "Item Size", myVendor.ItemSize, False
        validateUpdateField rng, i, CaseUPC, "Commodity", myVendor.Commodity, False
        validateUpdateField rng, i, CaseUPC, "Commodity Description", myVendor.CommodityDescription, False
        validateUpdateField rng, i, CaseUPC, "Sub Commodity", myVendor.SubCommodity, False
        validateUpdateField rng, i, CaseUPC, "Sub Commodity Description", myVendor.SubCommodityDescription, False
        validateUpdateField rng, i, CaseUPC, "Pack Cost", myVendor.CostAmt, True
      End If
    Next i
    
    If (found = False) Then
      Debug.Print "Did not find match for CaseID " & myVendor.CaseID
    End If
  Next


  ' MsgBox "Finished Processing " & mySheet
  Debug.Print "Finished Processing " & mySheet

  ' Restore the state and return
  returnState.restore
End Sub