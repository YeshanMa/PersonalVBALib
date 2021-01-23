'Attribute VB_Name = "wdSUB_PreProcessText"
Sub PreProcessText()
'This macro is to Process the Text for easy reading, e.g. remove multiple or Consecutive Line Breakers
'Pre-Process and Post-Process will be combined in one Sub, controlled with Argument of the Sub
'A separate Module for Pre-Process only for call by manual

Dim rngSelectedRange As Range

Set rngSelectedRange = Selection.Range

If rngSelectedRange Is Nothing Then

    Set rngSelectedRange = ActiveDocument.Range
    rngSelectedRange.Select

End If

'Special Characters in Word see below Link
'https://confluence.remc1.net/display/PS/Special+Characters+for+Find+and+Replace+in+Microsoft+Word


'---Automatic break lines to Split to Paragraphs for easy reading---


''----Replace Consecutive Line Breakers with Regular Expression----
'http://www.vbaexpress.com/forum/showthread.php?51480-basic-regex-in-Word-macro
'https://social.msdn.microsoft.com/Forums/office/en-US/b24911b8-071f-4c7e-8cfc-a8b82fecc435/vba-findreplaceexecute-loops-when-replacing-multiple-paragraph-marks

With Selection.Find

    .ClearFormatting
    
    .Text = "[^13]{3,}"
    .Replacement.Text = "^p^p"
    .MatchWildcards = True    'MatchWildcards must be Enabled to True when use RE
  
    .Forward = False
    .Wrap = wdFindContinue

    .Format = False
    .MatchCase = False
    .MatchWholeWord = False

    .MatchSoundsLike = False
    .MatchAllWordForms = False

End With
Selection.Find.Execute Replace:=wdReplaceAll
'-----------------------------------------


'----Replace Consecutive Line Breakers----
Selection.Find.Execute Replace:=wdReplaceAll

With Selection.Find

    .ClearFormatting
    .Text = "^t ^t"
    .Replacement.Text = "^p"

    .Wrap = wdFindContinue
    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

End With

Selection.Find.Execute Replace:=wdReplaceAll
'-----------------------------------------

'----Replace Consecutive Line Breakers----
With Selection.Find

    .ClearFormatting
    .Text = ". ^p"
    .Replacement.Text = ".^p^p"

    .Wrap = wdFindContinue
    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

End With

Selection.Find.Execute Replace:=wdReplaceAll

'-----------------------------------------


'----Replace Consecutive Line Breakers----
With Selection.Find

    .ClearFormatting
    .Text = "^t"
    .Replacement.Text = "^p"

    .Wrap = wdFindContinue
    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

End With

Selection.Find.Execute Replace:=wdReplaceAll
'-----------------------------------------

End Sub
