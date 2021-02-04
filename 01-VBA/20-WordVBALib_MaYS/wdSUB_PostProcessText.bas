'Attribute VB_Name = "wdSUB_PostProcessText"
Sub PostProcessText()
'This macro is to Process the Text for easy reading, e.g. remove multiple or Consecutive Line Breakers

Dim rngSelectedRange As Range
Set rngSelectedRange = Selection.Range

If rngSelectedRange Is Nothing Then
    Set rngSelectedRange = ActiveDocument.Range
    rngSelectedRange.Select
End If

Debug.Print rngSelectedRange.Text

'Special Characters in Word see below Link
'https://confluence.remc1.net/display/PS/Special+Characters+for+Find+and+Replace+in+Microsoft+Word


''----Replace Consecutive Tabs and Line Breakers with Regular Expression----
'http://www.vbaexpress.com/forum/showthread.php?51480-basic-regex-in-Word-macro
'https://social.msdn.microsoft.com/Forums/office/en-US/b24911b8-071f-4c7e-8cfc-a8b82fecc435/vba-findreplaceexecute-loops-when-replacing-multiple-paragraph-marks


'----Clean Consecutive Tabs into One Line Break----
With Selection.Find

    .ClearFormatting

    .Text = "[^t]{1,}"                 'MatchWildcards must be Enabled to True when use RE
    .Replacement.Text = "^p^p"

    .MatchWildcards = True          'MatchWildcards must be Enabled to True when use RE

    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

    '.Wrap = wdFindContinue         '.Wrap = wdFindContinue shall be disabled for only Replace in Selected Range

End With

Selection.Find.Execute Replace:=wdReplaceAll
''-----------------------------------------


'----Clean Consecutive (> 3) White Spaces----
With Selection.Find

    .ClearFormatting

    .Text = " {3,}"                 'MatchWildcards must be Enabled to True when use RE
    .Replacement.Text = " "

    .MatchWildcards = True          'MatchWildcards must be Enabled to True when use RE
    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

    '.Wrap = wdFindContinue         '.Wrap = wdFindContinue shall be disabled for only Replace in Selected Range

End With

Selection.Find.Execute Replace:=wdReplaceAll
''-----------------------------------------


''----Replace Line Breakers with Tabs before----
With Selection.Find

    .ClearFormatting

    .Text = " {1,}^13"             '^p cannot be recognized, use ^13 as alternative
    .Replacement.Text = "^p"

    .MatchWildcards = True          'MatchWildcards must be Enabled to True when use RE

    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

'    .Wrap = wdFindContinue         '.Wrap = wdFindContinue shall be disabled for only Replace in Selected Range


End With
Selection.Find.Execute Replace:=wdReplaceAll
''-----------------------------------------

''----Replace Consecutive (> 3) Line Breakers with RE----
With Selection.Find

    .ClearFormatting

    .Text = "[^13]{2,}"             '^p cannot be recognized, use ^13 as alternative
    .Replacement.Text = "^p^p"

    .MatchWildcards = True          'MatchWildcards must be Enabled to True when use RE

    .Forward = False
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

'    .Wrap = wdFindContinue         '.Wrap = wdFindContinue shall be disabled for only Replace in Selected Range

End With
Selection.Find.Execute Replace:=wdReplaceAll
''-----------------------------------------
With ActiveDocument.Paragraphs

.SpaceBefore = 0
.SpaceBeforeAuto = False
.SpaceAfter = 6
.SpaceAfterAuto = False

End With

End Sub

'A typical Find & Replace Template see:
'https://www.automateexcel.com/vba/word/find-replace
'https://docs.microsoft.com/en-us/office/vba/word/Concepts/Customizing-Word/finding-and-replacing-text-or-formatting

''----Code Example below for refer----
'Dim RegEx_01 As Object
''Set RegEx_01 = CreateObject("VBScript.RegExp")
'Set RegEx_01 = New RegExp
'
'With RegEx_01
'
'    .Pattern = "[^p]{3,}"
'    .Global = False
'
'    .Replace(,Replace)
'
'
'End With
