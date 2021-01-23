Attribute VB_Name = "Sub_CommentsMacrosLib"
Public Sub ConvertToComment()
'Convert Cell Content to Comment
'ShortCutKey: Ctrl + Shift + C

'Author:MaYS, Last Update:2020-12-23
'Different Comment Box Size according to the Length of Text

Dim CelltoAddComment As Range
Dim CellText As String

For Each CelltoAddComment In Selection
    
    With CelltoAddComment
        .ClearComments
    
        CellText = .Value
        
        If Len(CellText) > 1 Then
        
        .AddComment
        .Comment.Text CellText & ""
        
        End If
        
    With .Comment
 
        '.Shape.TextFrame.Characters.Font.Size = 10
        '.Shape.TextFrame.Characters.Font = Tahoma
        
        If Len(CellText) < 100 Then

            '.Shape.Width = 120
            '.Shape.Height = 60
            .Shape.TextFrame.AutoSize = True
            
         ElseIf Len(CellText) < 300 Then
    
            .Shape.Width = 400
            .Shape.Height = 150     'optimized Height by PR 393789
            '.Shape.TextFrame.AutoSize = True
        
        ElseIf Len(CellText) < 600 Then
    
            .Shape.Width = 400
            .Shape.Height = 250     'optimized Height by PR 426178, 408704
            '.Shape.TextFrame.AutoSize = True
            
        ElseIf Len(CellText) < 1000 Then
    
            .Shape.Width = 400
            .Shape.Height = 350
            '.Shape.TextFrame.AutoSize = True
            
        ElseIf Len(CellText) < 2000 Then
        
            .Shape.Width = 450
            .Shape.Height = 400
            '.Shape.TextFrame.AutoSize = True
                               
        ElseIf Len(CellText) < 4000 Then
        
            .Shape.Width = 650
            .Shape.Height = 750
            '.Shape.TextFrame.AutoSize = True
     
        Else

            .Shape.Width = 650
            .Shape.Height = 800
            '.Shape.TextFrame.AutoSize = True
                
    End If

    End With

    End With

Next CelltoAddComment
    
End Sub

Public Sub ResetCommentsPosition()
'Reset the position of all Comments after Sorting, AutoFilter or Reset AutoFilter
'ShortCutKey: Ctrl + Shift + R

'Author:MaYS, Last Update:2020-12-23
'Not to change the size of Comment Box

Dim CommentsPosReset As Comment

For Each CommentsPosReset In Application.ActiveSheet.Comments
   
   CommentsPosReset.Shape.Top = CommentsPosReset.Parent.Top + 5
   CommentsPosReset.Shape.Left = CommentsPosReset.Parent.Offset(0, 1).Left + 5
   
   'If CommentsPosReset.Shape.Width >= 200 And CommentsPosReset.Shape.Height >= 250 Then
   'Only Reset the Size of Comment Flag for the Text Invetigation Summary or Source Notes

    'CommentsPosReset.Shape.Width = 650
    'CommentsPosReset.Shape.Height = 750
   
   'End If
   
Next

End Sub


