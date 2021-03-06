'Attribute VB_Name = "wdSUB_MOD_CleanEmptyLines"
Sub CleanEmptyLines()

Dim rngSelectedRange As Range

Set rngSelectedRange = Selection.Range
Debug.Print rngSelectedRange.Text

'Typical Sequence of Clean the Empty Lines

Dim EmptyLinePattenCode() As String


'"[^t]{1,}"
'" {3,}"
' "{1,}[^t]{1,}^13"
' "[^t]{1,} {1,}^13"
' "[^13]{3,}"
''
'
If Len(rngSelectedRange.Text) < 2 Then

    Set rngSelectedRange = ActiveDocument.Range
    rngSelectedRange.Select

End If

'Debug.Print rngSelectedRange.Text

'Special Characters in Word see below Link
'https://confluence.remc1.net/display/PS/Special+Characters+for+Find+and+Replace+in+Microsoft+Word


''----Replace Consecutive Tabs and Line Breakers with Regular Expression----
'http://www.vbaexpress.com/forum/showthread.php?51480-basic-regex-in-Word-macro
'https://social.msdn.microsoft.com/Forums/office/en-US/b24911b8-071f-4c7e-8cfc-a8b82fecc435/vba-findreplaceexecute-loops-when-replacing-multiple-paragraph-marks


'----Clean Consecutive Tabs into One Line Break----

'For i = 0 To Len(EmptyLinePattenCode())

With Selection.Find

    .ClearFormatting

    .Text = "[^t]{1,}"                 'MatchWildcards must be Enabled to True when use RE
    .Replacement.Text = "^p"
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

'Next i

''-----------------------------------------


'----Clean Consecutive (> 3) White Spaces----
With Selection.Find

    .ClearFormatting

    .Text = " {3,}"                 'MatchWildcards must be Enabled to True when use RE
    .Replacement.Text = ""
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

    .Text = " {1,}[^t]{1,}^13"             '^p cannot be recognized, use ^13 as alternative

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

''----Replace Consecutive Line Breakers with Spaces----
With Selection.Find

    .ClearFormatting

    .Text = "[^t]{1,} {1,}^13"             '^p cannot be recognized, use ^13 as alternative

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

    .Text = "[^13]{3,}"             '^p cannot be recognized, use ^13 as alternative

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

Set rngSelectedRange = Nothing
Exit Sub

End Sub


�7
00
+	0
	*�H��
 � �iJ��w�=yY���ۂ���'��=C%�|����_x�[
��"�?�³���>�� ��ia|�p���
�SV[
�`s����m���8���f��Ћ:�޲�/!�"Y�^5�����|�a-g�� ��\���$+ng��</���Mu�����_`�mK�.&m��yZ�i�����֒��P�
��Ұe4����S��o86P7u`I�7\��?Y�t@{��R6ǟ7ܬ���2��dn��cƲ:I�=�����Q;�oHyX����o������1����`7��7^Fd}DIP�-��r,�.�������&xc�>�
�е�I��$�T�B���t<K<
�0�o�u��j��$����o�|[N�������p���z_
��x�
��)�`I��P� �5�Y|�ROH/�$�����t��u�3�ׅb}5�h�O������/H/UգmY��%DJ: