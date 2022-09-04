Function getColumnFromHeader(rng As Range, heading As String) As Integer
  Dim myIndex As Integer
  Dim myKey As String
  
  myKey = ActiveSheet.Name & heading
  
  If (headings.Exists(myKey)) Then
    myIndex = headings(myKey)
  Else
    myIndex = rng.Find(heading, LookIn:=xlValues, LookAt:=xlWhole).Column
    headings.Add myKey, myIndex
  End If
    
  getColumnFromHeader = myIndex
End Function
Function getScoutidFromUnknown(fname As String, lname As String) As String
  Dim i As Integer, myScout As scoutRecord, scoutID As String, found As Boolean
    
  getScoutidFromUnknown = ""
  found = False
  i = 2
  scoutID = "UnkScout" & i
  
  If (((fname = "") And (lname = "")) = False) Then
    While (found = False)
      If scouts.Exists(scoutID) Then
        Set myScout = scouts(scoutID)
        If ((myScout.id.FirstName = fname) And (myScout.id.LastName = lname)) Then
          getScoutidFromUnknown = myScout.id.BSAMemberID
          found = True
          'Debug.Print myScout.id.BSAMemberID & lname & fname
        End If
      End If
      i = i + 1
      scoutID = "UnkScout" & i
      If (i > scouts.Count) Then
        'Give up - We have seen them all
        found = True
      End If
    Wend
  End If
End Function

Sub loadSheet(mySheet As String)
  Dim rows As Integer, cols As Integer, i As Integer, j As Integer, rng As Range
  Dim myAdult As adultRecord, myScout As scoutRecord, myAdv As advancementRecord, myLog As logRecord, myPayment As paymentRecord
  
  Dim returnState As ExcelState
  Set returnState = New ExcelState
  
  Sheets(mySheet).Activate

  Set rng = ActiveSheet.UsedRange

  cols = rng.Columns.Count
  rows = rng.rows.Count
  
  Select Case mySheet

     Case "Adults"
      adults.RemoveAll
       For i = 2 To rows
         Set myAdult = New adultRecord
         myAdult.loadFromSheet rng, i
         If myAdult.id.BSAMemberID = "" Then
           myAdult.id.BSAMemberID = "UnkAdult" & i
         End If
         adults.Add myAdult.id.BSAMemberID, myAdult
       Next i

     Case "Scouts", "ScoutsArchive"
       If mySheet = "Scouts" Then
         scouts.RemoveAll
       End If
       For i = 2 To rows
         Set myScout = New scoutRecord
         myScout.loadFromSheet rng, i
         If myScout.id.BSAMemberID = "" Then
           myScout.id.BSAMemberID = "UnkScout" & i
         End If
         If mySheet = "Scouts" Then
           myScout.id.Archive = "Current"
         Else
           myScout.id.Archive = "Archive"
           If myScout.scoutDetail.PatrolName = "Historical" Then
             myScout.id.Archive = "Historical"
           End If
         End If
         scouts.Add myScout.id.BSAMemberID, myScout
       Next i

     Case "Advancement", "AdvancementArchive"
       If mySheet = "Advancement" Then
         mbnames.RemoveAll
         eaglembnames.RemoveAll
     
         eaglembnames("Camping") = "Camping"
         eaglembnames("Citizenship in the Community") = "Citizenship in the Community"
         eaglembnames("Citizenship in the Nation") = "Citizenship in the Nation"
         eaglembnames("Citizenship in the World") = "Citizenship in the World"
         eaglembnames("Communication") = "Communication"
         eaglembnames("Cooking") = "Cooking"
         eaglembnames("Cycling") = "Cycling"
         eaglembnames("Emergency Preparedness") = "Emergency Preparedness"
         eaglembnames("Environmental Science") = "Environmental Science"
         eaglembnames("Family Life") = "Family Life"
         eaglembnames("First Aid") = "First Aid"
         eaglembnames("Hiking") = "Hiking"
         eaglembnames("Lifesaving") = "Lifesaving"
         eaglembnames("Personal Fitness") = "Personal Fitness"
         eaglembnames("Personal Management") = "Personal Management"
         eaglembnames("Swimming") = "Swimming"
         eaglembnames("Sustainability") = "Sustainability"
       End If
       
       For i = 2 To rows
         Set myAdv = New advancementRecord
         myAdv.loadFromSheet rng, i
         If myAdv.AdvancementType = "Merit Badge" Then
           mbnames(myAdv.Advancement) = myAdv.Advancement
         End If
         If (scouts.Exists(myAdv.id.BSAMemberID) = False) Then
           myAdv.id.BSAMemberID = getScoutidFromUnknown(myAdv.id.FirstName, myAdv.id.LastName)
         End If
         If scouts.Exists(myAdv.id.BSAMemberID) Then
           Set myScout = scouts(myAdv.id.BSAMemberID)
           myScout.addAdvancement myAdv
         'Else
           'Debug.Print "Scout Not Found for Advancement BSAMemberID: " & myAdv.id.BSAMemberID
         End If
       Next i

     Case "Logs", "LogsArchive"
       For i = 2 To rows
         Set myLog = New logRecord
         myLog.loadFromSheet rng, i
         If (scouts.Exists(myLog.id.BSAMemberID) = False) Then
           myLog.id.BSAMemberID = getScoutidFromUnknown(myLog.id.FirstName, myLog.id.LastName)
         End If
         If scouts.Exists(myLog.id.BSAMemberID) Then
           Set myScout = scouts(myLog.id.BSAMemberID)
           myScout.addLog myLog
         'Else
           'Debug.Print "Scout Not Found for Log BSAMemberID: (" & myLog.id.BSAMemberID & ")"
         End If
       Next i

     Case "Payments"
       For i = 2 To rows
         Set myPayment = New paymentRecord
         myPayment.loadFromSheet rng, i
         If (scouts.Exists(myPayment.id.BSAMemberID) = False) Then
           myPayment.id.BSAMemberID = getScoutidFromUnknown(myPayment.id.FirstName, myPayment.id.LastName)
         End If
         If scouts.Exists(myPayment.id.BSAMemberID) Then
           Set myScout = scouts(myPayment.id.BSAMemberID)
           myScout.addPayment myPayment
         'Else
           'Debug.Print "Scout Not Found for Payment BSAMemberID: (" & myPayment.id.BSAMemberID & ")"
         End If
       Next i

     Case "MBCounselors"
       Dim mbcName As String
       Dim mbclist As String
       Dim mbitem As Variant

       mbcs.RemoveAll
       mbcmbnames.RemoveAll
       For i = 1 To rows
         mbcName = rng.Cells(i, 1)
         mbclist = rng.Cells(i, 2)
         mbcs(mbcName) = mbclist
         mbclist = Replace(mbclist, ", ", ",")
         For Each mbitem In Split(mbclist, ",")
           mbcmbnames(mbitem) = mbitem
         Next mbitem
       Next i

     Case Else
        MsgBox "Not Expected Condition"
  End Select

  ' MsgBox "Finished Processing " & mySheet
  Debug.Print "Finished Processing " & mySheet
  returnState.restore
End Sub

