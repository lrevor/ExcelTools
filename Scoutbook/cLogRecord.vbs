Private pID As IDRecord
Private pLogType As String
Private pDate As String
Private pNights As String
Private pDays As String
Private pMiles As String
Private pHours As String
Private pFrostPoints As String
Private pLocationName As String
Private pNotes As String

Public Property Get logType() As String
  logType = pLogType
End Property
Public Property Let logType(Value As String)
  pLogType = Value
End Property
Public Property Get logDate() As String
  logDate = pDate
End Property
Public Property Let logDate(Value As String)
  pDate = Value
End Property
Public Property Get Nights() As String
  Nights = pNights
End Property
Public Property Let Nights(Value As String)
  pNights = Value
End Property
Public Property Get Days() As String
  Days = pDays
End Property
Public Property Let Days(Value As String)
  pDays = Value
End Property
Public Property Get Miles() As String
  Miles = pMiles
End Property
Public Property Let Miles(Value As String)
  pMiles = Value
End Property
Public Property Get Hours() As String
  Hours = pHours
End Property
Public Property Let Hours(Value As String)
  pHours = Value
End Property
Public Property Get frostPoints() As String
  frostPoints = pFrostPoints
End Property
Public Property Let frostPoints(Value As String)
  pFrostPoints = Value
End Property
Public Property Get LocationName() As String
  LocationName = pLocationName
End Property
Public Property Let LocationName(Value As String)
  pLocationName = Value
End Property
Public Property Get Notes() As String
  Notes = pNotes
End Property
Public Property Let Notes(Value As String)
  pNotes = Value
End Property


Public Sub loadFromSheet(rng As Range, row As Integer)
  pID.BSAMemberID = rng.Cells(row, getColumnFromHeader(rng, "BSA Member ID"))
  pID.FirstName = rng.Cells(row, getColumnFromHeader(rng, "First Name"))
  pID.MiddleName = rng.Cells(row, getColumnFromHeader(rng, "Middle Name"))
  pID.LastName = rng.Cells(row, getColumnFromHeader(rng, "Last Name"))
  pLogType = rng.Cells(row, getColumnFromHeader(rng, "Log Type"))
  pDate = rng.Cells(row, getColumnFromHeader(rng, "Date"))
  pNights = rng.Cells(row, getColumnFromHeader(rng, "Nights"))
  pDays = rng.Cells(row, getColumnFromHeader(rng, "Days"))
  pMiles = rng.Cells(row, getColumnFromHeader(rng, "Miles"))
  pHours = rng.Cells(row, getColumnFromHeader(rng, "Hours"))
  pFrostPoints = rng.Cells(row, getColumnFromHeader(rng, "Frost Points"))
  pLocationName = rng.Cells(row, getColumnFromHeader(rng, "Location/Name"))
  pNotes = rng.Cells(row, getColumnFromHeader(rng, "Notes"))
End Sub
Public Function id() As IDRecord
  Set id = pID
End Function
Private Sub Class_Initialize()
  Set pID = New IDRecord
End Sub


