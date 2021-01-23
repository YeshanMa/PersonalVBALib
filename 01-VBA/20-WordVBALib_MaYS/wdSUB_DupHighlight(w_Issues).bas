'https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_win10-mso_365hp/is-it-possible-to-find-duplicate-paragraphs-or/1306ddd2-8a86-4d4c-ba33-fe747e65c37c

Sub DemoA_Paragraph()
Application.ScreenUpdating = False
Options.DefaultHighlightColorIndex = wdYellow
Dim RngA As Range, RngB As Range, i As Long, j As Long, strFnd As String
With ActiveDocument
  For i = 2 To .Paragraphs.Count - 1
    Set RngA = .Paragraphs(i - 1).Range
    Set RngB = .Range(.Paragraphs(i).Range.Start, .Range.End)
    With RngA
      strFnd = Trim(Split(.Text, vbCr)(0))
      If Len(strFnd) > 0 Then
        If .HighlightColorIndex <> wdYellow Then
          With RngB.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = strFnd
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Format = True
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            If .Found = True Then RngA.HighlightColorIndex = wdBrightGreen
          End With
        End If
      End If
    End With
    If i Mod 100 = 0 Then DoEvents
  Next
End With
Application.ScreenUpdating = True
End Sub
'
'For PC macro installation & usage instructions, see: http://www.gmayor.com/installing_macro.htm
'For Mac macro installation & usage instructions, see: https://wordmvp.com/Mac/InstallMacro.html
'
'300 pages is a lot to process, so don't expect instant results. Where duplicates are found, the first occurrences will be highlighted bright green; the duplicates will be highlighted yellow.
'
'After running the macro, selecting any paragraph highlighted bright green and using 'Find' with the reading highlight setting will help you to see both the original and its duplicates for editing. After you've completed that, simply use Word's highlighting tool to remove all remaining highlights.
'
'Sentences are more problematic, because VBA has no idea what a grammatical sentence is. For example, consider the following:
'Mr. Smith spent $1,234.56 at Dr. John's Grocery Store, to buy: 10.25kg of potatoes; 10kg of avocados; and 15.1kg of Mrs. Green's Mt. Pleasant macadamia nuts.
'For you and me, that would count as one sentence; for VBA it counts as 5 sentences. If you're prepared to live with that, you could run the following macro. Initial and duplicate highlights are pink and teal, respectively. Be prepared for a much longer wait.


'Below is a Procedure that basically the same as DemoA, the only difference is Paragraph -> Sentence.
Sub DemoB_Sentence()

Application.ScreenUpdating = False
Options.DefaultHighlightColorIndex = wdTeal
Dim RngA As Range, RngB As Range, i As Long, j As Long, strFnd As String
With ActiveDocument
  For i = 2 To .Sentences.Count - 1
    Set RngA = .Sentences(i - 1)
    Set RngB = .Range(.Sentences(i).Start, .Range.End)
    With RngA
      strFnd = Trim(Split(.Text, vbCr)(0))
      If Len(strFnd) > 0 Then
        If .HighlightColorIndex <> wdTeal Then
          With RngB.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = strFnd
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Format = True
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            If .Found = True Then RngA.HighlightColorIndex = wdPink
          End With
        End If
      End If
    End With
    If i Mod 100 = 0 Then DoEvents
  Next
End With
Application.ScreenUpdating = True
End Sub

' Cheers
' Paul Edstein
' (Fmr MS MVP - Word)
 

