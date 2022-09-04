Private pID As IDRecord
Private pAdvancementType As String
Private pAdvancement As String
Private pVersion As String
Private pDateCompleted As String
Private pApproved As String
Private pAwarded As String

Public Property Get AdvancementType() As String
  AdvancementType = pAdvancementType
End Property
Public Property Let AdvancementType(Value As String)
  pAdvancementType = Value
End Property
Public Property Get Advancement() As String
  Advancement = pAdvancement
End Property
Public Property Let Advancement(Value As String)
  pAdvancement = Value
End Property
Public Property Get Version() As String
  Version = pVersion
End Property
Public Property Let Version(Value As String)
  pVersion = Value
End Property
Public Property Get DateCompleted() As String
  DateCompleted = pDateCompleted
End Property
Public Property Let DateCompleted(Value As String)
  pDateCompleted = Value
End Property
Public Property Get Approved() As String
  Approved = pApproved
End Property
Public Property Let Approved(Value As String)
  pApproved = Value
End Property
Public Property Get Awarded() As String
  Awarded = pAwarded
End Property
Public Property Let Awarded(Value As String)
  pAwarded = Value
End Property

Public Sub debugAdv()
  Debug.Print "Adv " & pFirstName & " " & pLastName & " " & pAdvancementType & " " & pAdvancement & " " & pDateCompleted
End Sub

Public Sub loadFromSheet(rng As Range, row As Integer)
  pID.BSAMemberID = rng.Cells(row, getColumnFromHeader(rng, "BSA Member ID"))
  pID.FirstName = rng.Cells(row, getColumnFromHeader(rng, "First Name"))
  pID.MiddleName = rng.Cells(row, getColumnFromHeader(rng, "Middle Name"))
  pID.LastName = rng.Cells(row, getColumnFromHeader(rng, "Last Name"))
  pAdvancementType = rng.Cells(row, getColumnFromHeader(rng, "Advancement Type"))
  pAdvancement = rng.Cells(row, getColumnFromHeader(rng, "Advancement"))
  pVersion = rng.Cells(row, getColumnFromHeader(rng, "Version"))
  pDateCompleted = rng.Cells(row, getColumnFromHeader(rng, "Date Completed"))
  pApproved = rng.Cells(row, getColumnFromHeader(rng, "Approved"))
  pAwarded = rng.Cells(row, getColumnFromHeader(rng, "Awarded"))
End Sub
Public Function id() As IDRecord
  Set id = pID
End Function
Private Sub Class_Initialize()
  Set pID = New IDRecord
End Sub

