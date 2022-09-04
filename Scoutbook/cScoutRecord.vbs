Private pID As IDRecord
Private pIDDetail As IDDetailRecord
Private pScoutDetail As scoutDetailRecord

Private pAdvs As New Collection
Private pLogs As New Collection
Private pPayments As New Collection
Private pMeritBadges As advCollection
Private pRanks As advCollection

Private pPaymentBalance As Double
Private pCampNights As Integer
Private pCampDays As Integer
Private pCampouts As Integer
Private pFrostPoints As Integer
Private pActivities As Integer
Private pServiceActivities As Integer
Private pServiceHours As Integer
Private pHikes As Integer
Private pHikeMiles As Integer
Private pTwoYearShortTempCampNights As Integer
Private pTwoYearLongTempCampNights As Integer
Private pEligibleCampOA
Private pEligibleOA
Private pEligibleReqs

Public Property Get campNights() As String
  campNights = pCampNights
End Property
Public Property Get campDays() As String
  campDays = pCampDays
End Property
Public Property Get campouts() As String
  campouts = pCampouts
End Property
Public Property Get frostPoints() As String
  frostPoints = pFrostPoints
End Property
Public Property Get activities() As String
  activities = pActivities
End Property
Public Property Get serviceActivities() As String
  serviceActivities = pServiceActivities
End Property
Public Property Get serviceHours() As String
  serviceHours = pServiceHours
End Property
Public Property Get hikes() As String
  hikes = pHikes
End Property
Public Property Get hikeMiles() As String
  hikeMiles = pHikeMiles
End Property
Public Property Get twoYearShortTempCampNights() As String
  twoYearShortTempCampNights = pTwoYearShortTempCampNights
End Property
Public Property Get twoYearLongTempCampNights() As String
  twoYearLongTempCampNights = pTwoYearLongTempCampNights
End Property
Public Property Get eligibleCampOA() As String
  eligibleCampOA = pEligibleCampOA
End Property
Public Property Get eligibleOA() As String
  eligibleOA = pEligibleOA
End Property
Public Property Get eligibleReqs() As String
  eligibleReqs = pEligibleReqs
End Property
Public Function id() As IDRecord
  Set id = pID
End Function
Public Function idDetail() As IDDetailRecord
  Set idDetail = pIDDetail
End Function
Public Function scoutDetail() As scoutDetailRecord
  Set scoutDetail = pScoutDetail
End Function

Public Sub addAdvancement(adv As advancementRecord)
  pAdvs.Add adv
  If adv.AdvancementType = "Merit Badge" Then
    pMeritBadges.advAdd adv
  End If
  If adv.AdvancementType = "Rank" Then
    pRanks.advAdd adv
  End If
End Sub
Public Function getAdvancementDate(advType As String, adv As String) As String
  Dim myAdv As advancementRecord
  
  getAdvancementDate = ""
  For Each myAdv In pAdvs
    If ((myAdv.AdvancementType = advType) And (myAdv.Advancement = adv)) Then
      getAdvancementDate = myAdv.DateCompleted
    End If
  Next
End Function

Public Function rankDate(Key As String) As String
  rankDate = pRanks.advDate(Key)
End Function

Public Function MBDate(Key As String) As String
  MBDate = pMeritBadges.advDate(Key)
End Function

Public Sub addLog(log As logRecord)
  Dim logType As String
  Dim myDate As String

  logType = log.logType
  pLogs.Add log

  Select Case logType
     Case "Camping"
       'Scoutbook does not track non-camping activities, so we use 0 day camping instead.
       If log.Nights > 0 Then
         pCampNights = pCampNights + log.Nights
         pCampDays = pCampDays + log.Days
         pCampouts = pCampouts + 1
         pFrostPoints = pFrostPoints + log.frostPoints
       Else
         pActivities = pActivities + 1
       End If
       myDate = Date
       If ((DateDiff("m", log.logDate, myDate) <= 24) And (log.Nights <= 3)) Then
         pTwoYearShortTempCampNights = pTwoYearShortTempCampNights + log.Nights
       End If
       If ((DateDiff("m", log.logDate, myDate) <= 24) And (log.Nights >= 5)) Then
         pTwoYearLongTempCampNights = pTwoYearLongTempCampNights + log.Nights
       End If
     Case "Service"
       pServiceActivities = pServiceActivities + 1
       pServiceHours = pServiceHours + log.Hours
     Case "Hiking"
       pHikeMiles = pHikeMiles + log.Miles
       pHikes = pHikes + 1
     Case Else
        MsgBox "Not Expected Condition"
  End Select
End Sub

Public Sub addPayment(pay As paymentRecord)
  pPayments.Add pay
  pPaymentBalance = pPaymentBalance + pay.Amount
  'Debug.Print pBSAMemberID & " " & pPaymentBalance
End Sub

Public Sub checkAdvancements()
  Dim advDate As String
  If ((pCampNights > 0) And (rankDate("Tenderfoot") = "") And (getAdvancementDate("Tenderfoot Rank Requirement", "1b") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "Tenderfoot Rank Requirement 1b"
  End If
  If ((pServiceHours > 0) And (rankDate("Tenderfoot") = "") And (getAdvancementDate("Tenderfoot Rank Requirement", "7b") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "Tenderfoot Rank Requirement 7b"
  End If
  If ((pActivities > 4) And (pCampouts > 2) And (rankDate("Second Class") = "") And (getAdvancementDate("Second Class Rank Requirement", "1a") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "Second Class Rank Requirement 1a"
  End If
  If ((pServiceHours > 1) And (rankDate("Second Class") = "") And (getAdvancementDate("Second Class Rank Requirement", "8e") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "Second Class Rank Requirement 8e"
  End If
  If ((pActivities > 9) And (pCampouts > 2) And (rankDate("First Class") = "") And (getAdvancementDate("First Class Rank Requirement", "1a") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "First Class Rank Requirement 1a"
  End If
  If ((pIDDetail.SwimmingClassification = "Swimmer") And (rankDate("First Class") = "") And (getAdvancementDate("First Class Rank Requirement", "6a") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "First Class Rank Requirement 6a"
  End If
  If ((pServiceHours > 2) And (pServiceActivities > 1) And (rankDate("First Class") = "") And (getAdvancementDate("First Class Rank Requirement", "9d") = "")) Then
    If Not (pEligibleReqs = "") Then pEligibleReqs = pEligibleReqs & vbCrLf
    pEligibleReqs = pEligibleReqs & "First Class Rank Requirement 9d"
  End If
  If ((pTwoYearShortTempCampNights >= 10) And (pTwoYearLongTempCampNights >= 5)) Then pEligibleCampOA = "TRUE"
  If ((Not (rankDate("First Class") = "")) And (pEligibleCampOA = "TRUE")) Then pEligibleOA = "TRUE"
End Sub
Sub debugAdv()
  pRanks.debugAdv
  pMeritBadges.debugAdv
End Sub

Public Sub debugScout()
  Debug.Print "Scout " & pID.pFirstName & " " & pID.pLastName
  debugAdv
End Sub

Public Sub loadFromSheet(rng As Range, row As Integer)
  pID.BSAMemberID = rng.Cells(row, getColumnFromHeader(rng, "BSA Member ID"))
  pID.FirstName = rng.Cells(row, getColumnFromHeader(rng, "First Name"))
  pID.MiddleName = rng.Cells(row, getColumnFromHeader(rng, "Middle Name"))
  pID.LastName = rng.Cells(row, getColumnFromHeader(rng, "Last Name"))
  pIDDetail.Suffix = rng.Cells(row, getColumnFromHeader(rng, "Suffix"))
  pIDDetail.Nickname = rng.Cells(row, getColumnFromHeader(rng, "Nickname"))
  pIDDetail.Address1 = rng.Cells(row, getColumnFromHeader(rng, "Address 1"))
  pIDDetail.Address2 = rng.Cells(row, getColumnFromHeader(rng, "Address 2"))
  pIDDetail.City = rng.Cells(row, getColumnFromHeader(rng, "City"))
  pIDDetail.State = rng.Cells(row, getColumnFromHeader(rng, "State"))
  pIDDetail.Zip = rng.Cells(row, getColumnFromHeader(rng, "Zip"))
  pIDDetail.HomePhone = rng.Cells(row, getColumnFromHeader(rng, "Home Phone"))
  pIDDetail.Gender = rng.Cells(row, getColumnFromHeader(rng, "Gender"))
  pIDDetail.DOB = rng.Cells(row, getColumnFromHeader(rng, "DOB"))
  pScoutDetail.SchoolGrade = rng.Cells(row, getColumnFromHeader(rng, "School Grade"))
  pScoutDetail.SchoolName = rng.Cells(row, getColumnFromHeader(rng, "School Name"))
  pScoutDetail.LDS = rng.Cells(row, getColumnFromHeader(rng, "LDS"))
  pIDDetail.SwimmingClassification = rng.Cells(row, getColumnFromHeader(rng, "Swimming Classification"))
  pIDDetail.SwimmingClassificationDate = rng.Cells(row, getColumnFromHeader(rng, "Swimming Classification Date"))
  pScoutDetail.UnitNumber = rng.Cells(row, getColumnFromHeader(rng, "Unit Number"))
  pScoutDetail.UnitType = rng.Cells(row, getColumnFromHeader(rng, "Unit Type"))
  pScoutDetail.DateJoinedScouts = rng.Cells(row, getColumnFromHeader(rng, "Date Joined Boy Scouts"))
  pScoutDetail.DenType = rng.Cells(row, getColumnFromHeader(rng, "Den Type"))
  pScoutDetail.DenNumber = rng.Cells(row, getColumnFromHeader(rng, "Den Number"))
  pScoutDetail.DateJoinedDen = rng.Cells(row, getColumnFromHeader(rng, "Date Joined Den"))
  pScoutDetail.PatrolName = rng.Cells(row, getColumnFromHeader(rng, "Patrol Name"))
  pScoutDetail.DateJoinedPatrol = rng.Cells(row, getColumnFromHeader(rng, "Date Joined Patrol"))
  pScoutDetail.Parent1Email = rng.Cells(row, getColumnFromHeader(rng, "Parent 1 Email"))
  pScoutDetail.Parent2Email = rng.Cells(row, getColumnFromHeader(rng, "Parent 2 Email"))
  pScoutDetail.Parent3Email = rng.Cells(row, getColumnFromHeader(rng, "Parent 3 Email"))
End Sub

Private Sub Class_Initialize()
  pPaymentBalance = 0
  pCampNights = 0
  pCampDays = 0
  pCampouts = 0
  pFrostPoints = 0
  pActivities = 0
  pServiceActivities = 0
  pServiceHours = 0
  pHikes = 0
  pHikeMiles = 0
  pTwoYearShortTempCampNights = 0
  pTwoYearLongTempCampNights = 0
  pEligibleCampOA = ""
  pEligibleOA = ""
  pEligibleReqs = ""
  Set pRanks = New advCollection
  Set pMeritBadges = New advCollection
  Set pID = New IDRecord
  Set pIDDetail = New IDDetailRecord
  Set pScoutDetail = New scoutDetailRecord
End Sub

