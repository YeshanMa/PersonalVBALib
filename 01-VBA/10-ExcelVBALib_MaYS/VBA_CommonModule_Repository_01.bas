Attribute VB_Name = "VBA_CommonModule_Repository_01"
Public Sub GetUsedRangeRowCol(RowLast As Long, ColLast As Long)
'This Sub is a common module for all other module to get the Row and Column Nr of Used Range\

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    RowLast = 0
    ColLast = 0
    
    ActiveSheet.UsedRange.Select
    
    Cells(1, 1).Activate
    Selection.End(xlDown).Select
    'Selection.End(xlDown).Select
    
    On Error GoTo -1: On Error GoTo Quit
    
    Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Activate
    
    On Error GoTo -1: On Error GoTo 0
    
    RowLast = Selection.Row
    
    Cells(1, 1).Activate
    Selection.End(xlToRight).Select
    'Selection.End(xlToRight).Select
    
    Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Activate
    ColLast = Selection.Column
    
Quit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo -1: On Error GoTo 0
    
End Sub





Sub subGroupTest()
    Dim sRng As Range, eRng As Range
    Dim groupMap() As Variant
    Dim subGrp As Integer, i As Integer, j As Integer
    Dim startRow As Range, lastRow As Range
    Dim startGrp As Range, lastGrp As Range
    Dim k As Integer
    

    ReDim groupMap(1 To 2, 1 To 1)
    subGrp = 0
    i = 0
    Set startRow = Range("A1")

    ' Create a map of the groups with their cell addresses and an index of the lowest subgrouping
    Do While (startRow.Offset(i).Value <> "")
        groupMap(1, i + 1) = startRow.Offset(i).Address
        groupMap(2, i + 1) = UBound(Split(startRow.Offset(i).Value, "."))
        If subGrp < groupMap(2, i + 1) Then subGrp = groupMap(2, i + 1)
        ReDim Preserve groupMap(1 To 2, 1 To (i + 2))

        Set lastRow = Range(groupMap(1, i + 1))
        i = i + 1
    Loop

    ' Destroy already existing groups, otherwise we get errors
    On Error Resume Next
    For k = 1 To 10
        Rows(startRow.Row & ":" & lastRow.Row).EntireRow.Ungroup
    Next k
    On Error GoTo 0

    ' Create the groups
    ' We do them by levels in descending order, ie. all groups with an index of 3 are grouped individually before we move to index 2
    Do While (subGrp > 0)
        For j = LBound(groupMap, 2) To UBound(groupMap, 2)
            If groupMap(2, j) >= CStr(subGrp) Then
            ' If current value in the map matches the current group index

                ' Update group range references
                If startGrp Is Nothing Then
                    Set startGrp = Range(groupMap(1, j))
                End If
                Set lastGrp = Range(groupMap(1, j))
            Else
                ' If/when we reach this loop, it means we've reached the end of a subgroup

                ' Create the group we found in the previous loops
                If Not startGrp Is Nothing And Not lastGrp Is Nothing Then Range(startGrp, lastGrp).EntireRow.Group

                ' Then, reset the group ranges so they're ready for the next group we encounter
                If Not startGrp Is Nothing Then Set startGrp = Nothing
                If Not lastGrp Is Nothing Then Set lastGrp = Nothing
            End If
        Next j

        ' Decrement the index
        subGrp = subGrp - 1
    Loop
End Sub



