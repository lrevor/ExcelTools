Function getExcelFileNameDialog(strTitle As String) As String
  Dim strFileToOpen As String
  If Not Application.OperatingSystem Like "*Mac*" Then
        'I am Windows
    strFileToOpen = Application.GetOpenFilename(Title:=strTitle, _
                                                FileFilter:="Excel Files *.xls* (*.xls*),")
  Else
    'I am a Mac and will test if it is Excel 2011 or higher
    If Val(Application.Version) > 14 Then
      strFileToOpen = Application.GetOpenFilename(Title:=strTitle)
    End If
  End If
  getExcelFileNameDialog = strFileToOpen
End Function

Function getCSVFileNameDialog(strTitle As String) As String
  Dim strFileToOpen As String

  If Not Application.OperatingSystem Like "*Mac*" Then
        'I am Windows
    strFileToOpen = Application.GetOpenFilename(Title:=strTitle, _
                                                FileFilter:="CSV Files *.csv* (*.csv*),")
  Else
    'I am a Mac and will test if it is Excel 2011 or higher
    If Val(Application.Version) > 14 Then
      strFileToOpen = Application.GetOpenFilename(Title:=strTitle)
    End If
  End If
  getCSVFileNameDialog = strFileToOpen
End Function


