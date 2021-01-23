Attribute VB_Name = "wdSUB_withProblem_01"
Sub DeleteDuplicateParagraphs()
'PURPOSE: Remove Duplicate Paragraphs Throughout the Entire Word Document
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault



'Comments by MaYS, this Procedure works yet not work well when try to remove the Duplicated Text.
'Word always get stucked.


Dim p1 As Paragraph
Dim p2 As Paragraph
Dim DupCount As Long

'Something wrong with the code below that seems some paragraphs will be mistakenly removed and also
'The procedure will be stucked.

DupCount = 0

For Each p1 In ActiveDocument.Paragraphs
  If p1.Range.Text <> vbCr Then 'Ignore blank paragraphs
    
    For Each p2 In ActiveDocument.Paragraphs
    
      If p1.Range.Text = p2.Range.Text Then
      
        DupCount = DupCount + 1
        
        If p1.Range.Text = p2.Range.Text And DupCount > 1 Then
        
            p2.Range.Text = vbCr
            p2.Range.Text = "---DUPLICATED TEXT " & DupCount & " REMOVED---"
            
            'p2.Range.HighlightColorIndex = wdGray50
            
        End If
                
      End If
    Next p2
    
  End If
  
  'Reset Duplicate Counter
    DupCount = 0

Next p1

End Sub
