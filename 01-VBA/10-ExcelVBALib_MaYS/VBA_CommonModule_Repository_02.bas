Public Function CheckActShtContainsImage(ByVal thisSht As Worksheet)
'This Sub is a common module for all other module to check if the actived sheet contains picture
Dim xImage As Picture
Dim bImageExist As Boolean
Dim xImageName As String
    
'thisSht = ActiveSheet

On Error Resume Next
    
Application.ScreenUpdating = False

bImageExist = False

For Each xImage In thisSht.Pictures

    Debug.Print xImage.Name
    Debug.Print thisSht.Name
    
    
        If Len(xImage.Name) > 0 Then
            'MsgBox "The Image is on the Active Sheet", vbInformation, "KuTools For Excel"

            bImageExist = True

            Exit For
        End If
Next

CheckActShtContainsImage = bImageExist

Application.ScreenUpdating = True

End Function

Public Function CheckActShtContainsChart(ByVal thisSht As Worksheet)
'This Sub is a common module for all other module to check if the actived sheet contains Chart
Dim xChart As Chart
Dim bChartExist As Boolean
Dim xChartName As String
    
'thisSht = ActiveSheet

On Error Resume Next
    
Application.ScreenUpdating = False

bChartExist = False

For Each xChart In thisSht.ChartObjects

    Debug.Print xChart.Name
    Debug.Print thisSht.Name
    
    
        If Len(xChart.Name) > 0 Then
            'MsgBox "The Chart is on the Active Sheet", vbInformation, "KuTools For Excel"

            bChartExist = True

            Exit For
        End If
Next

CheckActShtContainsChart = bChartExist

Application.ScreenUpdating = True

End Function



