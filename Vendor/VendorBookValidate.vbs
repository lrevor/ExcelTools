' This is a helper routine
Sub validateUpdateField(rng As Range, i As Integer, upc As String, field As String, value As String)
  Dim j As Integer, data as String, boxTitle As String, boxData as String
  
  j = getColumnFromHeader(rng, field)
  data = rng.Cells(i, j)
  If (value <> data) Then
    boxTitle = "Update " & field & " from VendorExport for Case UPC (" & upc & ")?"
    boxData = "VendorBooks: " & data & " VendorExport: " & value
    If (MsgBox(boxData, vbYesNo, boxTitle) = vbYes) Then
      Debug.Print "Yes Selected"
      rng.Cells(i, j) = value
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
        Debug.Print "Found Case UPC: " & data
        found = True

        'Validate the Data
        validateUpdateField rng, i, CaseUPC, "Supplier #", myVendor.SupplierNumber
        validateUpdateField rng, i, CaseUPC, "Supplier Name", myVendor.SupplierName
        validateUpdateField rng, i, CaseUPC, "Cost Link", myVendor.CostLink
        validateUpdateField rng, i, CaseUPC, "Retail Link", myVendor.RetailLink
        validateUpdateField rng, i, CaseUPC, "HEB IC", myVendor.EntityID
        validateUpdateField rng, i, CaseUPC, "UPC", myVendor.UPC
        validateUpdateField rng, i, CaseUPC, "MFC", myVendor.MFC
        validateUpdateField rng, i, CaseUPC, "Description", myVendor.Description
        validateUpdateField rng, i, CaseUPC, "Master Pack", myVendor.MasterPack
        validateUpdateField rng, i, CaseUPC, "Pack", myVendor.Pack
        validateUpdateField rng, i, CaseUPC, "Item Size", myVendor.ItemSize
        validateUpdateField rng, i, CaseUPC, "Commodity", myVendor.Commodity
        validateUpdateField rng, i, CaseUPC, "Commodity Description", myVendor.CommodityDescription
        validateUpdateField rng, i, CaseUPC, "Sub Commodity", myVendor.SubCommodity
        validateUpdateField rng, i, CaseUPC, "Sub Commodity Description", myVendor.SubCommodityDescription
        validateUpdateField rng, i, CaseUPC, "Pack Cost", myVendor.CostAmt
      End If
    Next i
    
    If (found = False) Then
      Debug.Print "Did not find match for CaseID " & data
    End If
  Next


  ' MsgBox "Finished Processing " & mySheet
  Debug.Print "Finished Processing " & mySheet

  ' Restore the state and return
  returnState.restore
End Sub