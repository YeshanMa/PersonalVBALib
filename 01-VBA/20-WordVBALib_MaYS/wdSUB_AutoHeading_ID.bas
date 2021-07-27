'Attribute VB_Name = "wdSUB_CodeRepository"
Sub HighlightWords()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strWord As String
    Application.ScreenUpdating = False
    Set dbs = DBEngine.OpenDatabase("C:\Databases\MyDatabase.accdb")
    Set rst = dbs.OpenRecordset("tblWords", dbOpenForwardOnly)
    Options.DefaultHighlightColorIndex = wdYellow
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Highlight = True
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        Do While Not rst.EOF
            strWord = rst.Fields("MyWord")
            .Execute FindText:=strWord, ReplaceWith:=strWord, Replace:=wdReplaceAll
            rst.MoveNext
        Loop
    End With
    rst.Close
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing
    Application.ScreenUpdating = True
    
End Sub
