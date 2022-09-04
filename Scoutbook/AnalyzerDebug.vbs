Sub debugScouts()
  Dim scoutKey As Variant
  Dim myScout As scoutRecord
  For Each scoutKey In scouts.Keys
    Set myScout = scouts(scoutKey)
    myScout.debugScout
  Next
End Sub

Sub debugAdults()
  Dim adultKey As Variant
  Dim myAdult As adultRecord
  For Each adultKey In adults.Keys
    Set myAdult = adults(adultKey)
    myAdult.debugAdult
  Next
End Sub

