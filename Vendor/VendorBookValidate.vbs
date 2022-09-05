' This is the main processing routine
Sub validateVendors(mySheet As String)

  Dim rows As Integer, cols As Integer, i As Integer, j As Integer, rng As Range, data As String, found As Boolean

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
    Dim CaseUPC As String
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
        j = getColumnFromHeader(rng, "Supplier #")
        data = rng.Cells(i, j)
        If (myVendor.SupplierNumber = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Supplier # - VendorBooks: " & data & " VendorExport: " & myVendor.SupplierNumber
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Supplier # - VendorBooks: " & data & " VendorExport: " & myVendor.SupplierNumber
        End If
        j = getColumnFromHeader(rng, "Supplier Name")
        data = rng.Cells(i, j)
        If (myVendor.SupplierName = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Supplier Name - VendorBooks: " & data & " VendorExport: " & myVendor.SupplierName
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Supplier Name - VendorBooks: " & data & " VendorExport: " & myVendor.SupplierName
        End If
        j = getColumnFromHeader(rng, "Cost Link")
        data = rng.Cells(i, j)
        If (myVendor.CostLink = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Cost Link - VendorBooks: " & data & " VendorExport: " & myVendor.CostLink
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Cost Link - VendorBooks: " & data & " VendorExport: " & myVendor.CostLink
        End If
        j = getColumnFromHeader(rng, "Retail Link")
        data = rng.Cells(i, j)
        If (myVendor.RetailLink = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Retail Link - VendorBooks: " & data & " VendorExport: " & myVendor.RetailLink
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Retail Link - VendorBooks: " & data & " VendorExport: " & myVendor.RetailLink
        End If
        j = getColumnFromHeader(rng, "HEB IC")
        data = rng.Cells(i, j)
        If (myVendor.EntityID = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") HEB IC - VendorBooks: " & data & " VendorExport: " & myVendor.EntityID
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") HEB IC - VendorBooks: " & data & " VendorExport: " & myVendor.EntityID
        End If
        j = getColumnFromHeader(rng, "UPC")
        data = rng.Cells(i, j)
        If (myVendor.UPC = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") UPC - VendorBooks: " & data & " VendorExport: " & myVendor.UPC
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") UPC - VendorBooks: " & data & " VendorExport: " & myVendor.UPC
        End If
        j = getColumnFromHeader(rng, "MFC")
        data = rng.Cells(i, j)
        If (myVendor.MFC = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") MFC - VendorBooks: " & data & " VendorExport: " & myVendor.MFC
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") MFC - VendorBooks: " & data & " VendorExport: " & myVendor.MFC
        End If
        j = getColumnFromHeader(rng, "Description")
        data = rng.Cells(i, j)
        If (myVendor.Description = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Description - VendorBooks: " & data & " VendorExport: " & myVendor.Description
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Description - VendorBooks: " & data & " VendorExport: " & myVendor.Description
        End If
        j = getColumnFromHeader(rng, "Master Pack")
        data = rng.Cells(i, j)
        If (myVendor.MasterPack = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Master Pack - VendorBooks: " & data & " VendorExport: " & myVendor.MasterPack
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Master Pack - VendorBooks: " & data & " VendorExport: " & myVendor.MasterPack
        End If
        j = getColumnFromHeader(rng, "Pack")
        data = rng.Cells(i, j)
        If (myVendor.Pack = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Pack - VendorBooks: " & data & " VendorExport: " & myVendor.Pack
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Pack - VendorBooks: " & data & " VendorExport: " & myVendor.Pack
        End If
        j = getColumnFromHeader(rng, "Item Size")
        data = rng.Cells(i, j)
        If (myVendor.ItemSize = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Item Size - VendorBooks: " & data & " VendorExport: " & myVendor.ItemSize
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Item Size - VendorBooks: " & data & " VendorExport: " & myVendor.ItemSize
        End If
        j = getColumnFromHeader(rng, "Commodity")
        data = rng.Cells(i, j)
        If (myVendor.Commodity = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Commodity - VendorBooks: " & data & " VendorExport: " & myVendor.Commodity
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Commodity - VendorBooks: " & data & " VendorExport: " & myVendor.Commodity
        End If
        j = getColumnFromHeader(rng, "Commodity Description")
        data = rng.Cells(i, j)
        If (myVendor.CommodityDescription = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Commodity Description - VendorBooks: " & data & " VendorExport: " & myVendor.CommodityDescription
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Commodity Description - VendorBooks: " & data & " VendorExport: " & myVendor.CommodityDescription
        End If
        j = getColumnFromHeader(rng, "Sub Commodity")
        data = rng.Cells(i, j)
        If (myVendor.SubCommodity = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Sub Commodity - VendorBooks: " & data & " VendorExport: " & myVendor.SubCommodity
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Sub Commodity - VendorBooks: " & data & " VendorExport: " & myVendor.SubCommodity
        End If
        j = getColumnFromHeader(rng, "Sub Commodity Description")
        data = rng.Cells(i, j)
        If (myVendor.SubCommodityDescription = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Sub Commodity Description - VendorBooks: " & data & " VendorExport: " & myVendor.SubCommodityDescription
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Sub Commodity Description - VendorBooks: " & data & " VendorExport: " & myVendor.SubCommodityDescription
        End If
        j = getColumnFromHeader(rng, "Pack Cost")
        data = rng.Cells(i, j)
        If (myVendor.CostAmt = data) Then
          Debug.Print "MATCH: Case UPC (" & CaseUPC & ") Pack Cost - VendorBooks: " & data & " VendorExport: " & myVendor.CostAmt
        Else
          Debug.Print "DIFFERENCE: Case UPC (" & CaseUPC & ") Pack Cost - VendorBooks: " & data & " VendorExport: " & myVendor.CostAmt
        End If
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