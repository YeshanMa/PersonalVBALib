Sub DuplicateParagraphsDelete()
'This Macro is to Remove Duplicate Paragraphs Throughout the Entire Word Document.

'Basically identical with the Procedure wdSUB_DuplicateParagrGrayout.
'A procedure to remove the Duplicated Text was added in the end of the procedure.

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

'----------------------------------------------
'Below to remove the Duplicate Paragraphs.

Set currentParag = ActiveDocument.Paragraphs(1)

DupCount = 0

For k = 0 To PC - 1

    If k < PC - 1 Then Set currentParag = ActiveDocument.Paragraphs(k + 1)
    
    If currentParag.Range.HighlightColorIndex = wdGray50 And _
        currentParag.Range.Text <> vbCr Then
        
'        currentParag.Range.Delete

        DupCount = DupCount + 1
        currentParag.Range.Text = "---DUPLICATED TEXT " & DupCount & " REMOVED---"

        PC = PC - 1
'        this statement of PC = PC -1 cannot be removed

    End If

Next

End Sub
'=================================================