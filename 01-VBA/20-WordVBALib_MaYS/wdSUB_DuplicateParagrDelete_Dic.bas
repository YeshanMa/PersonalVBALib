Sub DeleteDuplicateParagraphs()

'https://stackoverflow.com/questions/33562468/duplicate-removal-for-vba-word-not-working-effectively
'Created on 04-Jun, seems works fine

  Dim p As Paragraph
  Dim d As Variant
  Dim t As Variant
  Dim i As Integer
  Dim StartTime As Single

  StartTime = Timer

    Set d = CreateObject("Scripting.Dictionary")
    
  ' collect duplicates
  For Each p In ActiveDocument.Paragraphs
  
    t = p.Range.Text
    
    If t <> vbCr Then
    
      If Not d.Exists(t) Then d.Add t, CreateObject("Scripting.Dictionary")
      
      d(t).Add d(t).Count + 1, p
      
    End If
  Next

  ' eliminate duplicates
  Application.ScreenUpdating = False
  
  For Each t In d
    
    For i = 2 To d(t).Count
      d(t)(i).Range.Delete
    Next
  
  Next
  
    Set d = Nothing
    Application.ScreenUpdating = True

  'MsgBox "This code ran successfully in " & Round(Timer - StartTime, 2) & " seconds", vbInformation

Call PostProcessText


    
    
End Sub


