Private pID As IDRecord
Private pPaymentType As String
Private pDate As Date
Private pDescription As String
Private pAmount As String
Private pTransactionID As String
Private pCategory As String
Private pNotes As String

Public Property Get PaymentType() As String
  PaymentType = pPaymentType
End Property
Public Property Let PaymentType(Value As String)
  pPaymentType = Value
End Property
Public Property Get PaymentDate() As Date
  PaymentDate = pDate
End Property
Public Property Let PaymentDate(Value As Date)
  pDate = Value
End Property
Public Property Get Description() As String
  Description = pDescription
End Property
Public Property Let Description(Value As String)
  pDescription = Value
End Property
Public Property Get Amount() As String
  Amount = pAmount
End Property
Public Property Let Amount(Value As String)
  pAmount = Value
End Property
Public Property Get TransactionID() As String
  TransactionID = pTransactionID
End Property
Public Property Let TransactionID(Value As String)
  pTransactionID = Value
End Property
Public Property Get Category() As String
  Category = pCategory
End Property
Public Property Let Category(Value As String)
  pCategory = Value
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
  pID.LastName = rng.Cells(row, getColumnFromHeader(rng, "Last Name"))
  pPaymentType = rng.Cells(row, getColumnFromHeader(rng, "Payment Type"))
  pDate = DateValue(rng.Cells(row, getColumnFromHeader(rng, "Date")))
  pDescription = rng.Cells(row, getColumnFromHeader(rng, "Description"))
  pAmount = rng.Cells(row, getColumnFromHeader(rng, "Amount"))
  pTransactionID = rng.Cells(row, getColumnFromHeader(rng, "Transaction ID"))
  pCategory = rng.Cells(row, getColumnFromHeader(rng, "Category"))
  pNotes = rng.Cells(row, getColumnFromHeader(rng, "Notes"))
End Sub
Public Function id() As IDRecord
  Set id = pID
End Function
Private Sub Class_Initialize()
  Set pID = New IDRecord
End Sub


