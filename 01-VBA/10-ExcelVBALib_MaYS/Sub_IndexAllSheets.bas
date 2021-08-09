Option Explicit
Public RowLast As Long
Public ColLast As Long

Public Sub IndexAllSheets()
' Macro to list the Index of All sheets in this Excel Workbook, by MaYS.

' Ver 1.2, 05-Aug-2021
' Improve Row Format for Protected/Hidden/Sheet with Contents.

' Ver 1.1, 27-Jul-2021
' Add Indicator of If the Workbook is empty on the Index Sheet
' Fix errors when add Link when the A1 content is Number

' Ver 1.0, 14-Dec-2020
' 1st Version
' https://www.thesmallman.com/blog/2020/4/16/list-all-sheets-in-a-excel-workbook


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
Dim bIfSheetEmpty As Boolean
Dim bIfSheetContainsImage As Boolean

bIfSheetProtected = False
bIfSheetHided = False
bIfSheetEmpty = False
bIfSheetContainsImage = False


Const COL_NR As Integer = 1
Const COL_SHEET_NAME As Integer = 2
Const COL_SHEET_LINK As Integer = 3
Const COL_SHEET_EMPTY As Integer = 4
Const COL_SHEET_LOCKED As Integer = 5
Const COL_SHEET_HIDDEN As Integer = 6
Const COL_RESERVED_01 As Integer = 7
Const COL_RESERVED_02 As Integer = 8
Const COL_COMMENTS As Integer = 9

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

    Cells(1, COL_NR).Value = "Nr."
    Cells(1, COL_SHEET_NAME).Value = "Index of All Sheets"
    Cells(1, COL_SHEET_LINK).Value = "Link to Each Sheets"
    Cells(1, COL_SHEET_EMPTY).Value = "Sheet w/ Contents ?"
    Cells(1, COL_SHEET_LOCKED).Value = "Sheet Locked ?"
    Cells(1, COL_SHEET_HIDDEN).Value = "Sheet Hided ?"
    Cells(1, COL_RESERVED_01).Value = "Sheet Status Reserved_01"
    Cells(1, COL_RESERVED_02).Value = "Sheet Status Reserved_02"
    Cells(1, COL_COMMENTS).Value = "Comments"

    With Range("A1:I1")
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
    Cells(j + 1, COL_NR).Value = j
    Cells(j + 1, COL_SHEET_NAME).Value = strSheetsNameList(j)

    'Cells(j + 1, COL_RESERVED_01).Value = "Reserved"
    'Cells(j + 1, COL_RESERVED_02).Value = "Reserved"
    
    '---------------------------------------------
    'Cells(j + 1, COL_SHEET_LINK).Value was operated along with each other sheet.
    shtIndexofAllSheets.Hyperlinks.Add Anchor:=Cells(j + 1, COL_SHEET_LINK), _
                            Address:="", _
                            SubAddress:="'" & strSheetsNameList(j) & "'" & "!A1", _
                            ScreenTip:="Click to Go to the Sheet", _
                            TextToDisplay:="Link to " & shtEachSheet.Name

'---------------------------------------------
'Check if each sheet is protected/locked, hidden, and empty.
    
    If shtEachSheet.ProtectContents = True Then
        bIfSheetProtected = True
        shtEachSheet.Unprotect
        Cells(j + 1, COL_SHEET_LOCKED).Value = "Y"
    End If


    bIfSheetContainsImage = CheckActShtContainsImage(shtEachSheet)
    If bIfSheetContainsImage = True Or WorksheetFunction.CountA(shtEachSheet.Range("A2:H200").Cells) >= 5 Then
    
        Cells(j + 1, COL_SHEET_EMPTY).Value = "Y"
       
        Rows(j + 1).Font.Color = vbBlue
        Rows(j + 1).Font.Bold = True
     
    Else
    
        bIfSheetEmpty = True
        'Cells(j + 1, COL_SHEET_EMPTY).Value = ""
        Rows(j + 1).Font.Color = vbBlack
        Rows(j + 1).Font.Bold = False
                
    End If
    
    
    If shtEachSheet.Visible = False Then
        'bIfSheetHided = True
        Cells(j + 1, COL_SHEET_HIDDEN).Value = "Y"
        
        Range(Cells(j + 1, COL_NR), Cells(j + 1, ColLast)).Font.Color = RGB(181, 181, 181)
        Rows(j + 1).Font.Color = RGB(181, 181, 181)
        Rows(j + 1).Font.Bold = False
        
    End If

'---------------------------------------------
'Add Link on A1 of each sheet, for quick go back to sheet of Index.

'    If shtEachSheet.Cells(1, 1).Value = "" Then
'        shtEachSheet.Cells(1, 1).Value = "Click to Index of Sheet"
'    End If
    
'    shtIndexofAllSheets.Hyperlinks.Add Anchor:=shtEachSheet.Cells(1, 1), _
'                                Address:="", _
'                                SubAddress:="'" & strIndexSheetName & "'" & "!A1", _
'                                ScreenTip:="Go Back to First Sheet of Index", _
'                                TextToDisplay:=Application.Text(shtEachSheet.Cells(1, 1).Value, "0")
                                
                                'If the Contents of A1 is a Number, then convert to Text.
'                                TextToDisplay:=shtEachSheet.Cells(1, 1).Value


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

'---------------------------------------------
'Format the Cell Grid Line, Font, Width for better reading.

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
Columns(COL_SHEET_NAME).ColumnWidth = Columns(COL_SHEET_NAME).ColumnWidth * 1.1
Columns(COL_SHEET_LINK).ColumnWidth = Columns(COL_SHEET_LINK).ColumnWidth * 1.1
Columns(COL_SHEET_EMPTY).HorizontalAlignment = xlCenter
Columns(COL_SHEET_LOCKED).HorizontalAlignment = xlCenter
Columns(COL_SHEET_HIDDEN).HorizontalAlignment = xlCenter

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


