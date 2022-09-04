Private pBSAMemberID As String
Private pFirstName As String
Private pMiddleName As String
Private pLastName As String
Private pArchive As String

Public Property Get BSAMemberID() As String
  BSAMemberID = pBSAMemberID
End Property
Public Property Let BSAMemberID(Value As String)
  pBSAMemberID = Trim(Value)
End Property
Public Property Get FirstName() As String
  FirstName = pFirstName
End Property
Public Property Let FirstName(Value As String)
  pFirstName = Value
End Property
Public Property Get MiddleName() As String
  MiddleName = pMiddleName
End Property
Public Property Let MiddleName(Value As String)
  pMiddleName = Value
End Property
Public Property Get LastName() As String
  LastName = pLastName
End Property
Public Property Let LastName(Value As String)
  pLastName = Value
End Property
Public Property Get Archive() As String
  Archive = pArchive
End Property
Public Property Let Archive(Value As String)
  pArchive = Value
End Property

