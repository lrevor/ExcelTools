Public adults As New Dictionary
Public scouts As New Dictionary
Public mbnames As New Dictionary
Public eaglembnames As New Dictionary
Public mbcs As New Dictionary
Public mbcmbnames As New Dictionary
Public headings As New Dictionary

Sub loadAll()
  Dim scoutKey As Variant
  Dim myScout As scoutRecord
  
  OptimizeCode_Begin
  headings.RemoveAll
  loadSheet ("Adults")
  loadSheet ("Scouts")
  loadSheet ("ScoutsArchive")
  loadSheet ("Payments")
  loadSheet ("Logs")
  loadSheet ("LogsArchive")
  loadSheet ("Advancement")
  loadSheet ("AdvancementArchive")
  loadSheet ("MBCounselors")
  sortMBNames
  sortMBCMBNames
  For Each scoutKey In scouts.Keys
    Set myScout = scouts(scoutKey)
    myScout.checkAdvancements
  Next
  createAdvancementReport
  createMBReport
  createMBCReport
  'createTemplateReport
  'debugScouts
  'debugAdults
  OptimizeCode_End
  MsgBox "Finished Loading"
End Sub

Sub sortMBNames()
  Dim tarray(1000) As String
  Dim i As Integer, j As Integer
  Dim str1 As String
  Dim mbCount As Integer
  
  mbCount = mbnames.Count
  
  i = 0
  For Each mbKey In mbnames.Keys
      tarray(i) = mbnames(mbKey)
      i = i + 1
  Next mbKey
  For i = 0 To mbCount - 2
    For j = i + 1 To mbCount - 1
      If (tarray(j) < tarray(i)) Then
        str1 = tarray(i)
        tarray(i) = tarray(j)
        tarray(j) = str1
      End If
    Next j
  Next i
  mbnames.RemoveAll
  For i = 0 To mbCount - 1
    mbnames(tarray(i)) = tarray(i)
  Next i
End Sub
Sub sortMBCMBNames()
  Dim tarray(1000) As String
  Dim i As Integer, j As Integer
  Dim str1 As String
  Dim mbCount As Integer
  
  mbCount = mbcmbnames.Count
  
  i = 0
  For Each mbKey In mbcmbnames.Keys
      tarray(i) = mbcmbnames(mbKey)
      i = i + 1
  Next mbKey
  For i = 0 To mbCount - 2
    For j = i + 1 To mbCount - 1
      If (tarray(j) < tarray(i)) Then
        str1 = tarray(i)
        tarray(i) = tarray(j)
        tarray(j) = str1
      End If
    Next j
  Next i
  mbcmbnames.RemoveAll
  For i = 0 To mbCount - 1
    mbcmbnames(tarray(i)) = tarray(i)
  Next i
End Sub
Sub refreshAdults()
  Dim fileName As String
  'Adults
  fileName = getCSVFileNameDialog("Filename for Scoutbook Adult Export")
  Range("B7").Value = fileName
  clearOrCreateSheet ("Adults")
  loadFileToSheet "Adults", fileName
End Sub

Sub refreshScouts()
  Dim fileName As String
  'Scouts
  fileName = getCSVFileNameDialog("Filename for Scoutbook Scout Export")
  Range("B8").Value = fileName
  clearOrCreateSheet ("Scouts")
  loadFileToSheet "Scouts", fileName
End Sub

Sub refreshAdvancement()
  Dim fileName As String
  'Advancement
  fileName = getCSVFileNameDialog("Filename for Scoutbook Advancement Export")
  Range("B9").Value = fileName
  clearOrCreateSheet ("Advancement")
  loadFileToSheet "Advancement", fileName
End Sub

Sub refreshLogs()
  Dim fileName As String
  'Logs
  fileName = getCSVFileNameDialog("Filename for Scoutbook Logs Export")
  Range("B10").Value = fileName
  clearOrCreateSheet ("Logs")
  loadFileToSheet "Logs", fileName
End Sub

Sub refreshPayments()
  Dim fileName As String
  'Payments
  fileName = getCSVFileNameDialog("Filename for Scoutbook Payments Export")
  Range("B11").Value = fileName
  clearOrCreateSheet ("Payments")
  loadFileToSheet "Payments", fileName
End Sub

