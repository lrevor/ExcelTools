' Helper routine here for speed/performance
Function getColumnFromHeader(rng As Range, heading As String) As Integer
  Dim myIndex As Integer
  Dim myKey As String
  
  myKey = ActiveSheet.Name & heading
  
  If (headings.Exists(myKey)) Then
    myIndex = headings.Item(myKey)
  Else
    myIndex = rng.Find(heading, LookIn:=xlValues, LookAt:=xlWhole).Column
    headings.Add myKey, myIndex
  End If
    
  getColumnFromHeader = myIndex
End Function

' This is the main processing routine
Sub loadSheet(mySheet As String)
  Dim rows As Integer, cols As Integer, i As Integer, j As Integer, rng As Range

  ' Initializae and sheet specific records
  Dim myVendor As VendorRecord
    
  ' Save the current state so it can be put back at the end of the call
  Dim returnState As ExcelState
  Set returnState = New ExcelState
  
  Sheets(mySheet).Activate

  Set rng = ActiveSheet.UsedRange

  cols = rng.Columns.Count
  rows = rng.rows.Count
  
  Select Case mySheet

     ' Add the names of the sheets and the required processing
     Case "VendorExport"
     vendorExports.RemoveAll
     For i = 6 To rows
       Set myVendor = New VendorRecord
       myVendor.loadFromSheet rng, i
       vendorExports.Add myVendor.CaseID, myVendor
    Next i
     
     ' Add processing required for this sheet
     Debug.Print "Add some processing to make this interesting"

     Case Else
        MsgBox "Not Expected Condition"
  End Select

  ' MsgBox "Finished Processing " & mySheet
  Debug.Print "Finished Processing " & mySheet

  ' Restore the state and return
  returnState.restore
End Sub