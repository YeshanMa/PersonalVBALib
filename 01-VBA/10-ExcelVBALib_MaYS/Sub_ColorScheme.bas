Attribute VB_Name = "Sub_ColorScheme"
Sub ColorScheme()

'---Define Variables for Range, Row and Column for Start and last---
Dim rngStartCell As Range
Set rngStartCell = Range("A1")
'Set rngStartCell = Application.InputBox("Select the Cell for Position Referenceqq2 ", Type:=8)

Dim nStartRow As Integer
Dim nStartColumn As Integer

nStartRow = rngStartCell.Row
nStartColumn = rngStartCell.Column

Dim nDataRangeStartRow As Integer
nDataRangeStartRow = nStartRow + 1

Dim nLastRow As Integer
Dim nLastColumn As Integer

'UsedRage are not recommended to be used to find the last Row and Column
nLastRow = ActiveSheet.UsedRange.Rows.Count
'nLastColumn = ActiveSheet.UsedRange.Columns.Count

'Excel 2003 and .csv File support max 65535 Rows and 256 columns
nLastRow = ActiveSheet.Range("A65536").End(xlUp).Row
nLastColumn = ActiveSheet.Range("EF1").End(xlToLeft).Column + 2   '2 More Columns Reserved

Dim rngRowofTitle As Range
Set rngRowofTitle = Range(Cells(1, 1), Cells(1, nLastColumn))
With rngRowofTitle

    .Interior.Color = RGB(89, 89, 89) 'Dark Gray
    .Font.Color = vbWhite
    .Font.Size = 11
    .Font.Bold = True
    .HorizontalAlignment = Excel.xlCenter
    .VerticalAlignment = Excel.xlCenter

End With


'---Define Sub Specified Range and Index of Column---
Dim rngColofColorHexCode As Range
Dim rngColofColorFilledArea As Range
Dim rngCellofColorFilledArea As Range

Dim rngColorR As Range
Dim rngColorG As Range
Dim rngColorB As Range

Dim nColofColorHexCode As Integer
Dim nColofColorFilledArea As Integer

Dim nColorR As Integer
Dim nColorG As Integer
Dim nColorB As Integer

'---Define Sub Specified Variables---
Dim nValueofR As Integer
Dim nValueofG As Integer
Dim nValueofB As Integer

Dim sHexColorofCurrentRow As String

'---Define Interators, some are Reserved---
Dim i As Integer
'Dim j As Integer
'Dim k As Integer
'Dim m As Integer
'Dim n As Integer


'---Start of Sub Procedure---

nColofColorHexCode = Application.IfError(Application.Match("Color HexCode", rngRowofTitle, 0), 0)
nColofColorFilledArea = Application.IfError(Application.Match("Color Filled for Example", rngRowofTitle, 0), 0)

If nColofColorHexCode = 0 Then

    Call MsgBox("No Column of Color Hex Code Found " & vbCrLf & _
        vbCrLf & "Color Hex Code shall be in Formart #FFFFFF, e.g. # E6F4AE, etc", _
            vbOKOnly, "No Color Hex Code")

    Exit Sub

End If

nColorR = Application.IfError(Application.Match("R", rngRowofTitle, 0), 0)
nColorG = Application.IfError(Application.Match("G", rngRowofTitle, 0), 0)
nColorB = Application.IfError(Application.Match("B", rngRowofTitle, 0), 0)

Set rngColofColorHexCode = Range(Cells(nDataRangeStartRow, nColofColorHexCode), Cells(nLastRow, nColofColorHexCode))
Set rngColofColorFilledArea = Range(Cells(nDataRangeStartRow, nColofColorFilledArea), Cells(nLastRow, nColofColorFilledArea))
Debug.Print nDataRangeStartRow
rngColofColorFilledArea.Select


Set rngColorR = Range(Cells(nDataRangeStartRow, nColorR), Cells(nLastRow, nColorR))
Set rngColorG = Range(Cells(nDataRangeStartRow, nColorG), Cells(nLastRow, nColorG))
Set rngColorB = Range(Cells(nDataRangeStartRow, nColorB), Cells(nLastRow, nColorB))


If nColorR = 0 Or nColorG = 0 Or nColorB = 0 Then

    Call MsgBox("Columns for Color Code RGB were Not Found" & vbCrLf & _
        vbCrLf & " ""R"",  ""G"",  ""B""", _
            vbOKOnly, "No RGB Column")

    Exit Sub

End If

'---Start of Fill Each Cell with Specific Color---

For Each rngCellofColorFilledArea In rngColofColorFilledArea

With rngCellofColorFilledArea
    
    i = .Row
    
    sHexColorofCurrentRow = Cells(i, nColofColorHexCode).Value
    Cells(i, nColorR).Value = GetRGBFromHex(sHexColorofCurrentRow, "R")
    Cells(i, nColorG).Value = GetRGBFromHex(sHexColorofCurrentRow, "G")
    Cells(i, nColorB).Value = GetRGBFromHex(sHexColorofCurrentRow, "B")
    
    nValueofR = Cells(i, nColorR).Value
    nValueofG = Cells(i, nColorG).Value
    nValueofB = Cells(i, nColorB).Value
    
    .Interior.Color = RGB(nValueofR, nValueofG, nValueofB)

    If nValueofR = 0 And nValueofG = 0 And nValueofB = 0 Then
    
       With Range(Cells(i, nStartColumn), Cells(i, nLastColumn))
       .Interior.Color = RGB(128, 128, 128)
       .Font.Color = RGB(128, 128, 128)
       End With
       
    End If

End With

Next rngCellofColorFilledArea

'---Set the worksheet format---
With ActiveSheet.UsedRange

    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    
    .NumberFormat = "0"
    
    .Font.Name = "Arial"
    .Font.Size = 10
    
    .AutoFilter
        
End With


End Sub


'---Define Functions Specifically for this Sub '
Public Function GetRGBFromHex(sHexColor As String, RGB As String) As String

sHexColor = VBA.Replace(sHexColor, "#", "")
sHexColor = VBA.Right$("000000" & sHexColor, 6)

Select Case RGB

    Case "B"
        GetRGBFromHex = VBA.Val("&H" & VBA.Mid(sHexColor, 5, 2))

    Case "G"
        GetRGBFromHex = VBA.Val("&H" & VBA.Mid(sHexColor, 3, 2))

    Case "R"
        GetRGBFromHex = VBA.Val("&H" & VBA.Mid(sHexColor, 1, 2))

End Select

End Function

'
'Sub DetermineVisualBasicHexColor()
''PURPOSE: Display Visual Basic HEX Color Code next to each cell's Fill Color
'
'Dim cell As Range
'Dim FillHexColor As String
'
''Ensure a cell range is selected
'  If TypeName(Selection) <> "Range" Then Exit Sub
'
''Loop through each cell in selection
'  For Each cell In Selection.Cells
'
'    'Ensure cell has a fill color
'      If cell.Interior.ColorIndex <> xlNone Then
'
'        'Get Hex values (values come through in reverse of what we need)
'          FillHexColor = Right("000000" & Hex(cell.Interior.Color), 6)
'
'        'Convert to the Visual Basic Userform Color Code Format
'          cell.Offset(0, 1).Value = "&H00" & FillHexColor & "&"
'
'      End If
'
'  Next cell

''Select just the ActiveCell
'  ActiveCell.Select
'
'End Sub





