Option Explicit
Public RowLast As Long
Public ColLast As Long

Public Sub IndexAllSheets()
'https://www.thesmallman.com/blog/2020/4/16/list-all-sheets-in-a-excel-workbook
'---------------------------------------------
'Define Variables for how many sheets in the Workbook

Dim nCountofSheets As Integer
Dim i As Integer 'Used for Iterator in Arr Operation
Dim j As Integer

Dim strSheetsNameList() As String

'Define Current Workbook'
Dim wbActivedWorkBook As Workbook
Set wbActivedWorkBook = ThisWorkbook

'Define New Created WorkSheet as Index
Dim shtIndexofAllSheets As Worksheet
Const strIndexSheetName As String = "Index_All_Sheets"

Dim shtEachSheet As Worksheet
Dim bIfSheetProtected As Boolean
Dim bIfSheetHided As Boolean

bIfSheetProtected = False
bIfSheetHided = False
'---------------------------------------------
'Start of Procedure

'Check for if sheet of Index already exist.
'And move the Sheet of Index to the first sheet of the Workbook

If Not Evaluate("ISREF('" & strIndexSheetName & "'!A1)") Then

    Set shtIndexofAllSheets = Worksheets.Add
    shtIndexofAllSheets.Name = strIndexSheetName
     
Else

    Set shtIndexofAllSheets = Worksheets(strIndexSheetName)

End If
    
    shtIndexofAllSheets.Move Before:=Sheets(1)
    Cells.Clear

    Cells(1, 1).Value = "Nr."
    Cells(1, 2).Value = "Index of All Sheets"
    Cells(1, 3).Value = "Link to Each Sheets"
    Cells(1, 4).Value = "Sheet Locked ?"
    Cells(1, 5).Value = "Sheet Hided ?"
    Cells(1, 6).Value = "Sheet Status Reserved_01"
    Cells(1, 7).Value = "Sheet Status Reserved_02"
    Cells(1, 8).Value = "Comments"

    With Range("A1:H1")
        .Interior.Color = RGB(181, 181, 181)
        .Font.Bold = True
    End With
    
Call GetUsedRangeRowCol(RowLast, ColLast)

nCountofSheets = Worksheets.Count - 1  'Did not count the Index Sheet.

ReDim strSheetsNameList(1 To nCountofSheets)

For i = 1 To nCountofSheets
    strSheetsNameList(i) = Sheets(i + 1).Name
Next i

For j = 1 To nCountofSheets

    Set shtEachSheet = Sheets(j + 1)

    'Start from the 2nd Rows
    Cells(j + 1, 1).Value = j
    Cells(j + 1, 2).Value = strSheetsNameList(j)
    'Cells(j + 1, 3).Value will be operated below.
    'Cells(j + 1, 4).Value will be operated below.
    'Cells(j + 1, 5).Value will be operated below.
    
    'Cells(j + 1, 6).Value = "Reserved"
    'Cells(j + 1, 7).Value = "Reserved"
    'Cells(j + 1, 8).Value = "Reserved"
    
    '---------------------------------------------
    'Cells(j + 1, 3).Value was operated along with each other sheet.
    shtIndexofAllSheets.Hyperlinks.Add Anchor:=Cells(j + 1, 3), _
                            Address:="", _
                            SubAddress:="'" & strSheetsNameList(j) & "'" & "!A1", _
                            ScreenTip:="Click to Go to the Sheet", _
                            TextToDisplay:="Link to " & shtEachSheet.Name
    
    If shtEachSheet.ProtectContents = True Then
        bIfSheetProtected = True
        shtEachSheet.Unprotect
        Cells(j + 1, 4).Value = "Y"
    End If

    If shtEachSheet.Visible = False Then
        'bIfSheetHided = True
        Cells(j + 1, 5).Value = "Y"
        Range(Cells(j + 1, 1), Cells(j + 1, ColLast)).Font.Color = RGB(181, 181, 181)
        
    End If

    If shtEachSheet.Cells(1, 1).Value = "" Then
        shtEachSheet.Cells(1, 1).Value = "Click to Index of Sheet"
    End If
    
    shtIndexofAllSheets.Hyperlinks.Add Anchor:=shtEachSheet.Cells(1, 1), _
                                Address:="", _
                                SubAddress:="'" & strIndexSheetName & "'" & "!A1", _
                                ScreenTip:="Go Back to First Sheet of Index", _
                                TextToDisplay:=shtEachSheet.Cells(1, 1).Value

    With shtEachSheet.Cells(1, 1)
            
            .Font.Color = vbBlue
            .Font.Underline = True
            .Font.Bold = True
                
    End With

    'Lock and Hide the Sheet Again

    If bIfSheetProtected = True Then
        shtEachSheet.Protect
        bIfSheetProtected = False
    
    ' ElseIf bIfSheetHided = True Then
    '     shtEachSheet.Visible = False
    '     bIfSheetHided = False
    
    End If
    
Next j

With ActiveSheet.UsedRange

    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    
    .HorizontalAlignment = xlLeft
    '.NumberFormat = "0"
    
    .Font.Name = "Segoe UI"
    .Font.Size = 12
    'Columns(3).Font.Size = 9
    
    .AutoFilter
        
End With

Rows.AutoFit
Columns.AutoFit
Columns(2).ColumnWidth = Columns(2).ColumnWidth * 1.1
Columns(3).ColumnWidth = Columns(3).ColumnWidth * 1.2
Columns(4).HorizontalAlignment = xlCenter
Columns(5).HorizontalAlignment = xlCenter

Set shtIndexofAllSheets = Nothing

End Sub

'Sub ListAllSheets() 'Excel VBA to list sheet names.
''https://www.thesmallman.com/blog/2020/4/16/list-all-sheets-in-a-excel-workbook
'
'Dim i As Integer
'Dim sh As Worksheet
'
' Const txt = "AllSheets"
'
'    If Not Evaluate("ISREF('" & txt & "'!A1)") Then 'Check for AllSheets tab.
'        Set sh = Worksheets.Add
'        sh.Name = txt
'        sh.[A1] = "Index List of All Sheet"
'    End If
'
'Set sh = Sheets(txt)
'    For i = 1 To Worksheets.Count
'        sh.Cells(i + 1, 1) = Sheets(i).Name
'
'        Next i
'
'End Sub

'----------------------------------------
'Below Modules are references.
'Sub CreateIndex()
''updateby Extendoffice
''https://www.extendoffice.com/documents/excel/572-excel-list-worksheet-names.html
'
'
'    Dim xAlerts As Boolean
'    Dim I  As Long
'    Dim xShtIndex As Worksheet
'    Dim xSht As Variant
'    xAlerts = Application.DisplayAlerts
'    Application.DisplayAlerts = False
'    On Error Resume Next
'    Sheets("Index").Delete
'    On Error GoTo 0
'    Set xShtIndex = Sheets.Add(Sheets(1))
'    xShtIndex.Name = "Index"
'    I = 1
'    Cells(1, 1).Value = "INDEX"
'    For Each xSht In ThisWorkbook.Sheets
'        If xSht.Name <> "Index" Then
'            I = I + 1
'            xShtIndex.Hyperlinks.Add Cells(I, 1), "", "'" & xSht.Name & "'!A1", , xSht.Name
'        End If
'    Next
'    Application.DisplayAlerts = xAlerts
'End Sub

'
'Sub ListSheetNamesInNewWorkbook()
''https://www.datanumen.com/blogs/3-quick-ways-to-get-a-list-of-all-worksheet-names-in-an-excel-workbook/
'
'    Dim objNewWorkbook As Workbook
'    Dim objNewWorksheet As Worksheet
'
'    Dim i As Integer
'
'    Set objNewWorkbook = Excel.Application.Workbooks.Add
'    Set objNewWorksheet = objNewWorkbook.Sheets(1)
'
'    For i = 1 To ThisWorkbook.Sheets.Count
'        objNewWorksheet.Cells(i, 1) = i
'        objNewWorksheet.Cells(i, 2) = ThisWorkbook.Sheets(i).Name
'    Next i
'
'    With objNewWorksheet
'         .Rows(1).Insert
'         .Cells(1, 1) = "INDEX"
'         .Cells(1, 1).Font.Bold = True
'         .Cells(1, 2) = "NAME"
'         .Cells(1, 2).Font.Bold = True
'         .Columns("A:B").AutoFit
'    End With
'
'End Sub
''






