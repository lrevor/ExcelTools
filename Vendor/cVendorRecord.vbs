Private pSupplierNumber As String
Private pSupplierName As String
Private pLevel As String
Private pCostLink As String
Private pRetailLink As String
Private pEntityID As String
Private pCaseID As String
Private pUPC As String
Private pMFC As String
Private pDescription As String
Private pMasterPack As String
Private pPack As String
Private pItemSize As String
Private pCommodity As String
Private pCommodityDescription As String
Private pSubCommodity As String
Private pSubCommodityDescription As String
Private pLocType As String
Private pLocID As String
Private pLocationDescription As String
Private pTotalInventory As String
Private pStartDt As String
Private pEndDt As String
Private pType As String
Private pCostAmt As String
Private pBasis As String
Private pStatus As String
Private pOfferID As String

Public Property Get SupplierNumber() As String
  SupplierNumber = pSupplierNumber
End Property
Public Property Let SupplierNumber(Value As String)
  pSupplierNumber = Value
End Property
Public Property Get SupplierName() As String
  SupplierName = pSupplierName
End Property
Public Property Let SupplierName(Value As String)
  pSupplierName = Value
End Property
Public Property Get Level() As String
  Level = pLevel
End Property
Public Property Let Level(Value As String)
  pLevel = Value
End Property
Public Property Get CostLink() As String
  CostLink = pCostLink
End Property
Public Property Let CostLink(Value As String)
  pCostLink = Value
End Property
Public Property Get RetailLink() As String
  RetailLink = pRetailLink
End Property
Public Property Let RetailLink(Value As String)
  pRetailLink = Value
End Property
Public Property Get EntityID() As String
  HoEntityIDurs = pEntityID
End Property
Public Property Let EntityID(Value As String)
  pEntityID = Value
End Property
Public Property Get CaseID() As String
  CaseID = pCaseID
End Property
Public Property Let CaseID(Value As String)
  pCaseID = Value
End Property
Public Property Get UPC() As String
  UPC = pUPC
End Property
Public Property Let UPC(Value As String)
  pUPC = Value
End Property
Public Property Get MFC() As String
  MFC = pMFC
End Property
Public Property Let MFC(Value As String)
  pMFC = Value
End Property
Public Property Get Description() As String
  Description = pDescription
End Property
Public Property Let Description(Value As String)
  pDescription = Value
End Property
Public Property Get MasterPack() As String
  MasterPack = pMasterPack
End Property
Public Property Let MasterPack(Value As String)
  pMasterPack = Value
End Property
Public Property Get Pack() As String
  Pack = pPack
End Property
Public Property Let Pack(Value As String)
  pPack = Value
End Property
Public Property Get ItemSize() As String
  ItemSize = pItemSize
End Property
Public Property Let ItemSize(Value As String)
  pItemSize = Value
End Property
Public Property Get Commodity() As String
  Commodity = pCommodity
End Property
Public Property Let Commodity(Value As String)
  pCommodity = Value
End Property
Public Property Get CommodityDescription() As String
  CommodityDescription = pCommodityDescription
End Property
Public Property Let CommodityDescription(Value As String)
  pCommodityDescription = Value
End Property
Public Property Get SubCommodity() As String
  SubCommodity = pSubCommodity
End Property
Public Property Let SubCommodity(Value As String)
  pSubCommodity = Value
End Property
Public Property Get SubCommodityDescription() As String
  SubCommodityDescription = pSubCommodityDescription
End Property
Public Property Let SubCommodityDescription(Value As String)
  pSubCommodityDescription = Value
End Property
Public Property Get LocType() As String
  LocType = pLocType
End Property
Public Property Let LocType(Value As String)
  pLocType = Value
End Property
Public Property Get LocID() As String
  LocID = pLocID
End Property
Public Property Let LocID(Value As String)
  pLocID = Value
End Property
Public Property Get LocationDescription() As String
  LocationDescription = pLocationDescription
End Property
Public Property Let LocationDescription(Value As String)
  pLocationDescription = Value
End Property
Public Property Get TotalInventory() As String
  TotalInventory = pTotalInventory
End Property
Public Property Let TotalInventory(Value As String)
  pTotalInventory = Value
End Property
Public Property Get StartDt() As String
  StartDt = pStartDt
End Property
Public Property Let StartDt(Value As String)
  pStartDt = Value
End Property
Public Property Get EndDt() As String
  EndDt = pEndDt
End Property
Public Property Let EndDt(Value As String)
  pEndDt = Value
End Property
Public Property Get Type() As String
  Type = pType
End Property
Public Property Let Type(Value As String)
  pType = Value
End Property
Public Property Get CostAmt() As String
  CostAmt = pCostAmt
End Property
Public Property Let CostAmt(Value As String)
  pCostAmt = Value
End Property
Public Property Get Basis() As String
  Basis = pBasis
End Property
Public Property Let Basis(Value As String)
  pBasis = Value
End Property
Public Property Get Status() As String
  Status = pStatus
End Property
Public Property Let Status(Value As String)
  pStatus = Value
End Property
Public Property Get OfferID() As String
  OfferID = pOfferID
End Property
Public Property Let OfferID(Value As String)
  pOfferID = Value
End Property

Public Sub loadFromSheet(rng As Range, row As Integer)
  pSupplierNumber = rng.Cells(row, getColumnFromHeader(rng, "Supplier #"))
  pSupplierName = rng.Cells(row, getColumnFromHeader(rng, "Supplier Name"))
  pLevel = rng.Cells(row, getColumnFromHeader(rng, "Level"))
  pCostLink = rng.Cells(row, getColumnFromHeader(rng, "Cost Link"))
  pRetailLink = rng.Cells(row, getColumnFromHeader(rng, "Retail Link"))
  pEntityID = rng.Cells(row, getColumnFromHeader(rng, "Entity ID"))
  pCaseID = rng.Cells(row, getColumnFromHeader(rng, "Case UPC"))
  pUPC = rng.Cells(row, getColumnFromHeader(rng, "UPC"))
  pMFC = rng.Cells(row, getColumnFromHeader(rng, "MFC"))
  pDescription = rng.Cells(row, getColumnFromHeader(rng, "Description"))
  pMasterPack = rng.Cells(row, getColumnFromHeader(rng, "Master Pack"))
  pPack = rng.Cells(row, getColumnFromHeader(rng, "Pack"))
  pItemSize = rng.Cells(row, getColumnFromHeader(rng, "Item Size"))
  pCommodity = rng.Cells(row, getColumnFromHeader(rng, "Commodity"))
  pCommodityDescription = rng.Cells(row, getColumnFromHeader(rng, "Commodity Description"))
  pSubCommodity = rng.Cells(row, getColumnFromHeader(rng, "Sub Commodity"))
  pSubCommodityDescription = rng.Cells(row, getColumnFromHeader(rng, "Sub Commodity Description"))
  pLocType = rng.Cells(row, getColumnFromHeader(rng, "Loc Type"))
  pLocID = rng.Cells(row, getColumnFromHeader(rng, "Loc Id"))
  pLocationDescription = rng.Cells(row, getColumnFromHeader(rng, "Location Description"))
  pTotalInventory = rng.Cells(row, getColumnFromHeader(rng, "Total Inventory"))
  pStartDt = rng.Cells(row, getColumnFromHeader(rng, "Start Dt"))
  pEndDt = rng.Cells(row, getColumnFromHeader(rng, "End Dt"))
  pType = rng.Cells(row, getColumnFromHeader(rng, "Type"))
  pCostAmt = rng.Cells(row, getColumnFromHeader(rng, "Cost Amt"))
  pBasis = rng.Cells(row, getColumnFromHeader(rng, "Basis"))
  pStatus = rng.Cells(row, getColumnFromHeader(rng, "Status"))
  pOfferID = rng.Cells(row, getColumnFromHeader(rng, "Offer ID"))
End Sub


