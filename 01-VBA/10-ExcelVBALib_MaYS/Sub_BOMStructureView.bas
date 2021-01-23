Sub BOMStructureView()

'This Macro process a BOM file with automatic Grouping with BOM Levels, Mark and indent for each level, for easy reading and find specific level of parts.
'Author: Ma Yeshan, yeshan.ma@outlook.com

'Ver 2.0, 12-Jan-2020

'Latest Update: Not allowed to translated in Wrong Cell or Range, and if with wrong language Code
'TobeUpdate in Next Version:

        '1. Status Bar to Display the status with Big BOM File.
        '2. Adjust the BOM Level 2 Color for more conspicuous

'Earlier History and Release Notes see the end of the Sub.



'Define Colors for each BOM Level
Dim BOM_LEVEL_COLOR_0 As Long
Dim BOM_LEVEL_COLOR_1 As Long
Dim BOM_LEVEL_COLOR_2 As Long
Dim BOM_LEVEL_COLOR_3 As Long
Dim BOM_LEVEL_COLOR_4 As Long
Dim BOM_LEVEL_COLOR_5 As Long
Dim BOM_LEVEL_COLOR_6 As Long
Dim BOM_LEVEL_COLOR_7 As Long
Dim BOM_LEVEL_COLOR_8 As Long
Dim BOM_LEVEL_COLOR_9 As Long
Dim BOM_LEVEL_COLOR_10 As Long
Dim BOM_LEVEL_COLOR_11 As Long
Dim BOM_LEVEL_COLOR_12 As Long

BOM_LEVEL_COLOR_0 = RGB(255, 255, 255)

'MoRandi Color Scheme
'------------------------------------
BOM_LEVEL_COLOR_1 = RGB(255, 250, 244)
BOM_LEVEL_COLOR_2 = RGB(236, 236, 234)
BOM_LEVEL_COLOR_3 = RGB(224, 229, 223)

BOM_LEVEL_COLOR_4 = RGB(181, 196, 177)
BOM_LEVEL_COLOR_5 = RGB(191, 191, 191)
BOM_LEVEL_COLOR_6 = RGB(175, 176, 178)
BOM_LEVEL_COLOR_7 = RGB(156, 168, 184)
BOM_LEVEL_COLOR_8 = RGB(147, 147, 145)
BOM_LEVEL_COLOR_9 = RGB(134, 150, 167)
BOM_LEVEL_COLOR_10 = RGB(150, 164, 139)
BOM_LEVEL_COLOR_11 = RGB(123, 139, 111)
BOM_LEVEL_COLOR_12 = RGB(101, 101, 101)

'------------------------------------


'Define Variables for Range Define
Dim rngStartCell As Range 'This defines the highest level of assembly, usually 1, and must be the top leftmost cell of concern for outlining, its our starting point for grouping
Dim nStartRow As Integer 'This defines the starting row to beging grouping, based on the row we define from rngStartCell
Dim nLevelCol As Integer 'This is the column that defines the assembly level we're basing our grouping on
Dim nLastRow As Integer 'This is the last row in the sheet that contains information we're grouping
Dim nLastColumn As Integer
Dim bProperBOMFile As Boolean

Dim nCurrentLevel As Integer 'iterative counter'
Dim i As Integer
Dim j As Integer

'Define Variables for AutoIndent and Formatting
Dim rngRowofTitle As Range
Dim rngCellBOMLevel As Range
Dim rngColBOMLevel As Range
Dim sRawNrofBOMLevel As Variant

Dim nStartBOMLevel As Integer
Dim nMaxBOMLevel As Integer
Dim nBOMLevelDepth As Integer

'Define Variables for Parse Rev and ECO/LCO Nr.
Dim rngCellRevECO As Range
Dim rngColofRevNr As Range
Dim rngColofECONr As Range

Dim nColNrofRevECO As Integer
Dim nColNrofRev As Integer
Dim nColNrofECO As Integer

Dim bFlagofRevExist As Boolean


Application.ScreenUpdating = False 'Turns off screen updating while running.

'Remove any pre-existing outlining and formart on worksheet

ActiveSheet.UsedRange.ClearOutline
ActiveSheet.UsedRange.ClearFormats

'Prompts user to select the starting row. It MUST be the highest level of assembly and also the top left cell of the range you want to group/outline"
'Set rngStartCell = Application.InputBox("Select top left cell for highest assembly level", Type:=8)
Set rngStartCell = Range("A1")

nStartRow = rngStartCell.Row
nLevelCol = rngStartCell.Column

'nLastRow = ActiveSheet.UsedRange.Rows.Count
'nLastColumn = ActiveSheet.UsedRange.Columns.Count + 2

'Excel 2003 and .csv File support max 65535 Rows and 256 columns
nLastRow = ActiveSheet.Range("A65536").End(xlUp).Row
nLastColumn = ActiveSheet.Range("EF1").End(xlToLeft).Column + 2

'Check if the Excel is a Proper BOM file that contains BOM data with correct BOM Level Nr in Column A
bProperBOMFile = True

If nLastRow < 3 Then

    bProperBOMFile = False
    
    GoTo NotProperBOMFile
End If


nStartBOMLevel = BOMLevelNrExtract(Cells(nStartRow, nLevelCol).Value)

For i = nStartRow + 1 To nLastRow

    sRawNrofBOMLevel = Cells(i, nLevelCol).Value
    nCurrentLevel = BOMLevelNrExtract(sRawNrofBOMLevel)
    DoEvents

'---Check if the Max BOM level is great than 8---
    If nCurrentLevel > 8 Then
    
        nMaxBOMLevel = nCurrentLevel
            
    End If
    
    If nCurrentLevel > 15 Or nCurrentLevel < 0 Then
    
        bProperBOMFile = False
        
        Exit For
        
    End If
    
Next i

If nMaxBOMLevel > 8 Then

    Call MsgBox("Max BOM Level > 8 will not be grouped." & vbCrLf & _
        "Please manually process the BOM before run this Macro on BOM.", vbOKOnly, "Too Much BOM Levels")

End If

NotProperBOMFile:

If bProperBOMFile = False Then
 
    Call MsgBox("Please run this Macro on a worksheet contains BOM data with correct BOM Level Nr in Column A" & vbCrLf & _
        vbCrLf & "e.g. Number: 1 ~ 12,  or BOM Level Codes Like: "".1"",  ""..2"",  ""....4"", etc.", _
            vbOKOnly, "Not Proper BOM File")
    
    Exit Sub
    
End If


Set rngRowofTitle = Range(Cells(1, 1), Cells(1, nLastColumn))
rngRowofTitle.Interior.Color = vbBlack
rngRowofTitle.Font.Color = vbWhite

'---Find the Column of Rev and Insert 2 Columns separated for Rev and ECO/LCO respectively---

'On Error if Col of Rev is not founded, resume the procesure
If Application.IfError(Application.Match("Rev", rngRowofTitle, 0), 0) > 0 Then
    
    bFlagofRevExist = True
    nColNrofRevECO = Application.Match("Rev", rngRowofTitle, 0)
    
ElseIf Application.IfError(Application.Match("Item Rev", rngRowofTitle, 0), 0) > 0 Then
    
    bFlagofRevExist = True
    nColNrofRevECO = Application.Match("Item Rev", rngRowofTitle, 0)
    
Else
    bFlagofRevExist = False
    Call MsgBox("No Column of ""Rev"" was founded", vbOKOnly, "No Rev Column")

End If

If bFlagofRevExist = True Then
    
nColNrofRev = nColNrofRevECO + 1
nColNrofECO = nColNrofRevECO + 2
    
If Application.IfError(Application.Match("Rev Nr", rngRowofTitle, 0), 0) = 0 Then
    Columns(nColNrofRevECO).Offset(, 1).Insert
    Cells(1, nColNrofRev).Value = "Rev Nr"
End If

If Application.IfError(Application.Match("ECO/LCO Nr", rngRowofTitle, 0), 0) = 0 Then
    Columns(nColNrofRev).Offset(, 1).Insert
    Cells(1, nColNrofECO).Value = "ECO/LCO Nr"
    
End If

End If



'---Start of Operation of Each Row---

For i = nStartRow + 1 To nLastRow

'---Procedure for Parse Rev and ECO/LCO Nr---
    If bFlagofRevExist = True Then
    
        Cells(i, nColNrofRev).Value = VBA.Trim(VBA.Left(Cells(i, nColNrofRevECO).Value, 3))
        Cells(i, nColNrofECO).Value = VBA.Trim(VBA.Right(Cells(i, nColNrofRevECO).Value, 9))

    End If
    
    
'---Procedure for AutoIndent and Formatting---

    sRawNrofBOMLevel = Cells(i, nLevelCol).Value
    nCurrentLevel = BOMLevelNrExtract(sRawNrofBOMLevel)

    With Range(Cells(i, 1), Cells(i, nLastColumn))

    Select Case nCurrentLevel
    
      Case 0
        .IndentLevel = 0
        .Font.Bold = True
        .Interior.Color = BOM_LEVEL_COLOR_0

      Case 1
        .IndentLevel = 0
        .Font.Bold = True
        .Font.Color = vbBlack
        .Interior.Color = BOM_LEVEL_COLOR_1
         
      Case 2
        .IndentLevel = 1
        .Font.Bold = True
        .Font.Color = vbBlue
        .Interior.Color = BOM_LEVEL_COLOR_2

      Case 3
        .IndentLevel = 2
        .Interior.Color = BOM_LEVEL_COLOR_3

      Case 4
        .IndentLevel = 3
        .Interior.Color = BOM_LEVEL_COLOR_4

      Case 5
        .IndentLevel = 4
        .Interior.Color = BOM_LEVEL_COLOR_5

      Case 6
        .IndentLevel = 5
        .Interior.Color = BOM_LEVEL_COLOR_6

      Case 7
        .IndentLevel = 6
        .Interior.Color = BOM_LEVEL_COLOR_7

      Case 8
        .IndentLevel = 7
        .Interior.Color = BOM_LEVEL_COLOR_8

      Case 9
        .IndentLevel = 8
        .Interior.Color = BOM_LEVEL_COLOR_9

      Case 10
        .IndentLevel = 9
        .Interior.Color = BOM_LEVEL_COLOR_10

      Case 11
        .IndentLevel = 10
        .Interior.Color = BOM_LEVEL_COLOR_11

      Case 12
        .IndentLevel = 11
        .Interior.Color = BOM_LEVEL_COLOR_12

      Case Else
         DoEvents
         
    End Select
    
    End With
   
'---Procedure for AutoGrouping---

'https://social.technet.microsoft.com/Forums/en-US/324726d4-9a86-430b-b6e8-9abc890cb645/max-limit-of-excel-grouping#:~:text=The%20maximum%20number%20of%20group%20levels%20in%20an,table%20allows%20for%20more%20than%208%20group%20levels.
'https://stackoverflow.com/questions/8335677/is-there-a-constraint-on-the-depth-level-of-grouping-in-excel
'https://social.technet.microsoft.com/Forums/office/en-US/324726d4-9a86-430b-b6e8-9abc890cb645/max-limit-of-excel-grouping
'The maximum number of group levels in an outline is 8. There is no way to increase that.
'A pivot table allows for more than 8 group levels.

If nCurrentLevel > 8 Then
    nCurrentLevel = 8
End If

    Rows(i).Select
    For j = 1 To nCurrentLevel - 1
        Selection.Rows.Group
    Next j
    
    DoEvents
    
Next i

'---Set the worksheet format---
With ActiveSheet.UsedRange

    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    
    .HorizontalAlignment = xlLeft
    .NumberFormat = "0"
    
    .Font.Name = "Arial"
    .Font.Size = 10
    
    .AutoFilter
        
End With

If bFlagofRevExist = True Then
    'Columns(nColNrofRevECO).Hidden = True
    Columns(nColNrofRev).NumberFormat = "00"

End If

Rows.AutoFit
Columns.AutoFit

Application.ScreenUpdating = True 'Turns on screen updating when done.

End Sub


Public Function BOMLevelNrExtract(vBOMLevelNr As Variant)

Dim sExtractedNrofBOMLevel As String
Dim sNrofBOMLevel As String
Dim nNrofBOMLevel As Integer

sExtractedNrofBOMLevel = vBOMLevelNr

sNrofBOMLevel = Replace(Expression:=sExtractedNrofBOMLevel, Find:="¡­", Replace:="")
sNrofBOMLevel = Replace(Expression:=sNrofBOMLevel, Find:=".", Replace:="")

sNrofBOMLevel = VBA.Right(sNrofBOMLevel, 3)
sNrofBOMLevel = VBA.Trim(sNrofBOMLevel)

nNrofBOMLevel = Val(sNrofBOMLevel)
Debug.Print nNrofBOMLevel

BOMLevelNrExtract = nNrofBOMLevel

End Function


''Color AI_Combo Color Scheme
''------------------------------------
'BOM_LEVEL_COLOR_1 = RGB(254, 246, 235)
'BOM_LEVEL_COLOR_2 = RGB(221, 221, 221)
'BOM_LEVEL_COLOR_3 = RGB(206, 223, 206)
'BOM_LEVEL_COLOR_4 = RGB(189, 207, 189)
'BOM_LEVEL_COLOR_5 = RGB(128, 206, 185)
'BOM_LEVEL_COLOR_6 = RGB(148, 186, 231)
'BOM_LEVEL_COLOR_7 = RGB(99, 154, 206)
'BOM_LEVEL_COLOR_8 = RGB(102, 144, 131)
'BOM_LEVEL_COLOR_9 = RGB(179, 188, 204)
'BOM_LEVEL_COLOR_10 = RGB(141, 150, 168)
'BOM_LEVEL_COLOR_11 = RGB(115, 124, 140)
'BOM_LEVEL_COLOR_12 = RGB(89, 97, 113)

''------------------------------------


''QNAP Color Scheme
''------------------------------------
'BOM_LEVEL_COLOR_1 = RGB(254, 246, 235)
'BOM_LEVEL_COLOR_2 = RGB(206, 223, 206)
'BOM_LEVEL_COLOR_3 = RGB(179, 202, 198)
'BOM_LEVEL_COLOR_4 = RGB(128, 167, 160)
'BOM_LEVEL_COLOR_5 = RGB(148, 186, 231)
'BOM_LEVEL_COLOR_6 = RGB(99, 154, 206)
'BOM_LEVEL_COLOR_7 = RGB(68, 132, 208)
'BOM_LEVEL_COLOR_8 = RGB(124, 149, 155)
'BOM_LEVEL_COLOR_9 = RGB(179, 188, 204)
'BOM_LEVEL_COLOR_10 = RGB(128, 136, 152)
'BOM_LEVEL_COLOR_11 = RGB(115, 124, 140)
'BOM_LEVEL_COLOR_12 = RGB(89, 97, 113)
''------------------------------------



''MoRandi Color Scheme
''------------------------------------
'BOM_LEVEL_COLOR_1 = RGB(255, 250, 244)
'BOM_LEVEL_COLOR_2 = RGB(236, 236, 234)
'BOM_LEVEL_COLOR_3 = RGB(224, 229, 223)
'BOM_LEVEL_COLOR_4 = RGB(181, 196, 177)
'BOM_LEVEL_COLOR_5 = RGB(191, 191, 191)
'BOM_LEVEL_COLOR_6 = RGB(175, 176, 178)
'BOM_LEVEL_COLOR_7 = RGB(156, 168, 184)
'BOM_LEVEL_COLOR_8 = RGB(147, 147, 145)
'BOM_LEVEL_COLOR_9 = RGB(134, 150, 167)
'BOM_LEVEL_COLOR_10 = RGB(150, 164, 139)
'BOM_LEVEL_COLOR_11 = RGB(123, 139, 111)
'BOM_LEVEL_COLOR_12 = RGB(101, 101, 101)
'
''------------------------------------