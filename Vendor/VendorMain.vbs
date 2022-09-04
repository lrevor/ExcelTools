Public headings As New Dictionary

Sub loadAll()
  ' Turn on optimizations
  OptimizeCode_Begin

  ' Clear the headings cache
  headings.RemoveAll

  ' Load the data sheets
  loadSheet ("TEMP")

  ' Process the data
  
  ' Turn off the optimizations
  OptimizeCode_End

  ' End the routine
  MsgBox "Finished Loading"
End Sub
