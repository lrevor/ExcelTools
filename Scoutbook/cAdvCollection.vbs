Private pCol As New Dictionary


Public Sub advAdd(adv As advancementRecord)
  pCol.Add adv.Advancement, adv
End Sub

Public Function advDate(Key As String) As String
  Dim adv As advancementRecord
  Dim myDate As String
  myDate = ""
  If (pCol.Exists(Key)) Then
     Set adv = pCol(Key)
     myDate = adv.DateCompleted
  End If
  advDate = myDate
End Function

Public Sub debugAdv()
  Dim advKey As Variant
  Dim myAdv As advancementRecord
  For Each advKey In pCol.Keys
    Set myAdv = pCol(advKey)
    myAdv.debugAdv
  Next
End Sub
