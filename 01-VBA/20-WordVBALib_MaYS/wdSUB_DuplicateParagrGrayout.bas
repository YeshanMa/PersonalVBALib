Sub DuplicateParagraphsGray()

'This Macro is to Find the Duplicate Paragraphs Throughout the Entire Word Document.
'Reference to below which highlight the Duplicated Paragraph.
'https://www.extendoffice.com/documents/word/5450-word-find-duplicate-sentences.html
'https://stackoverflow.com/questions/33562468/duplicate-removal-for-vba-word-not-working-effectively

' I had a very long document to process, the code above would take at least 100 days to finish and blocked everything
' while working at it. The main culprit is the "Set xRng = .Paragraphs(J).Range" which is very slow. I did an alternative
' version which ran in just 4 hours and presents a continuous report on the processing status and time to end.
' (To see the report in real time you have to open the "immediate window" by pressing Ctrl+G in the Microsoft Visual Basic for Applications window.)
' The code works well, except that it predicts a longer time to end than is actually the case (depends on the document). The code is as follows:

Dim StartTime, SecondsElapsed As Date

Dim secondsPerComparison As Double
Dim i, j, k, PC, DupCount As Long
Dim totalComparisons, comparisonsDone, C, secondsToFinish As Long

Dim xRngFind, xRng As Range
Dim xStrg, minutesToFinish As String
Dim currentParag, nextParag As Paragraph

'Options.DefaultHighlightColorIndex = wdYellow
Application.ScreenUpdating = False

With ActiveDocument

StartTime = Now()
C = 0
PC = .Paragraphs.Count
totalComparisons = CLng((PC * (PC + 1)) / 2)
Set currentParag = .Paragraphs(1)

For i = 1 To PC - 1
    'Debug.Print "processing paragraph " & I & " of a total of " & PC & " " & currentParag.Range.Text
    'Debug.Print Len(currentParag) & currentParag
    If currentParag.Range.HighlightColorIndex <> wdGray50 Then
        If currentParag.Range.HighlightColorIndex <> wdBrightGreen Then

            Set nextParag = currentParag

            For j = i + 1 To PC

                Set nextParag = nextParag.Next
                
                If currentParag.Range.Text = nextParag.Range.Text Then
                    currentParag.Range.HighlightColorIndex = wdBrightGreen
                    nextParag.Range.HighlightColorIndex = wdGray50
                    'Debug.Print "found one!! " & " I = " & I & " J = " & J & nextParag.Range.Text
                End If
            
            Next

            End If
                        
    End If
    
    DoEvents

    comparisonsDone = PC * (i - 1) + (j - i)
    
    SecondsElapsed = DateDiff("s", StartTime, Now())
    secondsPerComparison = CLng(SecondsElapsed) / comparisonsDone
    secondsToFinish = CLng(secondsPerComparison * (totalComparisons - comparisonsDone))
    minutesToFinish = Format(secondsToFinish / 86400, "hh:mm:ss")
    
    elapsedTime = Format(SecondsElapsed / 86400, "hh:mm:ss")
    
    Debug.Print "Finished procesing paragraph " & i & " of " & PC & ". Elapsed time = " & elapsedTime & ". Time to finish = " & minutesToFinish
    Set currentParag = currentParag.Next

Next

End With

