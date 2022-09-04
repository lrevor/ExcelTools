Sub createAdvancementReport()
  Dim scoutKey As Variant
  Dim myScout As scoutRecord
  Dim i As Integer
  Dim j As Integer
  Dim rng As Range
  
  Dim returnState As ExcelState
  Set returnState = New ExcelState
 
  clearOrCreateSheet ("Advancements")
  activeWorkbook.Sheets("Advancements").Visible = xlSheetVisible
  activeWorkbook.Sheets("Advancements").Select
  ActiveSheet.UsedRange.Interior.ColorIndex = 0
  
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
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Scout"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Tenderfoot"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Second Class"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "First Class"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Star Scout"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Life Scout"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Eagle Scout"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Campouts"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "CampNights"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "CampDays"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "twoYearShortTempCampNights"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "twoYearLongTempCampNights"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "FrostPoints"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Hikes"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "HikeMiles"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Activities"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Service Activities"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Service Hours"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Eligible Camp OA"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Eligible OA"
  i = i + 1
  ActiveSheet.Cells(1, i).Value = "Eligible Requirements"
  
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
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Scout")).Value = myScout.rankDate("Scout")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Tenderfoot")).Value = myScout.rankDate("Tenderfoot")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Second Class")).Value = myScout.rankDate("Second Class")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "First Class")).Value = myScout.rankDate("First Class")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Star Scout")).Value = myScout.rankDate("Star Scout")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Life Scout")).Value = myScout.rankDate("Life Scout")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Eagle Scout")).Value = myScout.rankDate("Eagle Scout")
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Campouts")).Value = myScout.campouts
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "CampNights")).Value = myScout.campNights
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "CampDays")).Value = myScout.campDays
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "twoYearShortTempCampNights")).Value = myScout.twoYearShortTempCampNights
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "twoYearLongTempCampNights")).Value = myScout.twoYearLongTempCampNights
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "FrostPoints")).Value = myScout.frostPoints
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Hikes")).Value = myScout.hikes
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "HikeMiles")).Value = myScout.hikeMiles
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Activities")).Value = myScout.activities
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Service Activities")).Value = myScout.serviceActivities
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Service Hours")).Value = myScout.serviceHours
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Eligible Camp OA")).Value = myScout.eligibleCampOA
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Eligible OA")).Value = myScout.eligibleOA
    ActiveSheet.Cells(j, getColumnFromHeader(rng, "Eligible Requirements")).Value = myScout.eligibleReqs
  Next
  
  ActiveSheet.Range("1:1").Orientation = xlUpward
  ActiveSheet.Range("1:1").HorizontalAlignment = xlCenter
  ActiveSheet.UsedRange.Columns.AutoFit
  returnState.restore
End Sub


