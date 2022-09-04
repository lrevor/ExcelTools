Private pID As IDRecord
Private pIDDetail As IDDetailRecord
Private pAdultDetail As adultDetailRecord

Public Function id() As IDRecord
  Set id = pID
End Function
Public Function idDetail() As IDDetailRecord
  Set idDetail = pIDDetail
End Function
Public Function adultDetail() As adultDetailRecord
  Set adultDetail = pAdultDetail
End Function
Public Sub debugAdult()
  Debug.Print "Adult " & pFirstName & " " & pLastName
End Sub

Public Sub loadFromSheet(rng As Range, row As Integer)
  pID.BSAMemberID = rng.Cells(row, getColumnFromHeader(rng, "BSA Member ID"))
  pID.FirstName = rng.Cells(row, getColumnFromHeader(rng, "First Name"))
  pID.MiddleName = rng.Cells(row, getColumnFromHeader(rng, "Middle Name"))
  pID.LastName = rng.Cells(row, getColumnFromHeader(rng, "Last Name"))
  'MsgBox "name " & pFirstName & pMiddleName & pLastName
  pIDDetail.Suffix = rng.Cells(row, getColumnFromHeader(rng, "Suffix"))
  pIDDetail.Nickname = rng.Cells(row, getColumnFromHeader(rng, "Nickname"))
  pIDDetail.ScouterTitle = rng.Cells(row, getColumnFromHeader(rng, "Scouter Title"))
  pIDDetail.Email = rng.Cells(row, getColumnFromHeader(rng, "Email"))
  pIDDetail.Address1 = rng.Cells(row, getColumnFromHeader(rng, "Address 1"))
  pIDDetail.Address2 = rng.Cells(row, getColumnFromHeader(rng, "Address 2"))
  pIDDetail.City = rng.Cells(row, getColumnFromHeader(rng, "City"))
  pIDDetail.State = rng.Cells(row, getColumnFromHeader(rng, "State"))
  pIDDetail.Zip = rng.Cells(row, getColumnFromHeader(rng, "Zip"))
  pIDDetail.HomePhone = rng.Cells(row, getColumnFromHeader(rng, "Home Phone"))
  pIDDetail.MobilePhone = rng.Cells(row, getColumnFromHeader(rng, "Mobile Phone"))
  pIDDetail.WorkPhone = rng.Cells(row, getColumnFromHeader(rng, "Work Phone"))
  pIDDetail.Gender = rng.Cells(row, getColumnFromHeader(rng, "Gender"))
  pIDDetail.DOB = rng.Cells(row, getColumnFromHeader(rng, "DOB"))
  pIDDetail.SwimmingClassification = rng.Cells(row, getColumnFromHeader(rng, "Swimming Classification"))
  pIDDetail.SwimmingClassificationDate = rng.Cells(row, getColumnFromHeader(rng, "Swimming Classification Date"))
  pAdultDetail.Occupation = rng.Cells(row, getColumnFromHeader(rng, "Occupation"))
  pAdultDetail.Employer = rng.Cells(row, getColumnFromHeader(rng, "Employer"))
  pAdultDetail.POS1 = rng.Cells(row, getColumnFromHeader(rng, "Leader Position 1"))
  pAdultDetail.POS1Date = rng.Cells(row, getColumnFromHeader(rng, "Position 1 Start Date"))
  pAdultDetail.POS2 = rng.Cells(row, getColumnFromHeader(rng, "Leader Position 2"))
  pAdultDetail.POS2Date = rng.Cells(row, getColumnFromHeader(rng, "Position 2 Start Date"))
  pAdultDetail.POS3 = rng.Cells(row, getColumnFromHeader(rng, "Leader Position 3"))
  pAdultDetail.POS3Date = rng.Cells(row, getColumnFromHeader(rng, "Position 3 Start Date"))
  pAdultDetail.POS4 = rng.Cells(row, getColumnFromHeader(rng, "Leader Position 4"))
  pAdultDetail.POS4Date = rng.Cells(row, getColumnFromHeader(rng, "Position 4 Start Date"))
  pAdultDetail.POS5 = rng.Cells(row, getColumnFromHeader(rng, "Leader Position 5"))
  pAdultDetail.POS5Date = rng.Cells(row, getColumnFromHeader(rng, "Position 5 Start Date"))
End Sub
Private Sub Class_Initialize()
  Set pID = New IDRecord
  Set pIDDetail = New IDDetailRecord
  Set pAdultDetail = New adultDetailRecord
End Sub
