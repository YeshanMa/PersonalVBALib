Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Option Explicit
Public sDetectedLanguageCode_From
Public sDetectedLanguageCode_To

Public Sub MSTranslatorScript()
'
' AutoTranslator Macro, by MaYS.

' Ver 3.6, 20-Apr-2021
' Latest Update:
            '1. Add sLanguageCode_From sLanguageCode_To to allow to Translate to Languge other than EN.
            '2. Translate to EN is still default if sLanguageCode_To is empty
' To be Update:
            '1. Clean the error value before Translating (e.g. "#NAME", "#VALUE", etc..)
            '2. Use Regular Expression for better Duplicated / Redudant Text Remove, e.g. Safety Questions, Empty Lines, etc...
            
            
'Earlier History and Release Notes see the end of the Sub.

' Keyboard Shortcut: Ctrl+Shift+T
' The Raw Data must be ognized in 4 Column.
' 1st Column: Text need to be translated
' 2nd Column: The ISO 639 Language Code
' 3rd Column: A empty column that will call the MSTranslate() function. This column will be filled with "Tranlated and Copied" when the translation is done.
' 4th Column: The translated Column will be copied to the 4th Column.

'------------------------------------------------------
'Define Variables
Dim UnTranslatedTextOrginal As String
Dim UnTranslatedTextSubstitued As String 'Replace the special characters: " \ / in the Original Text

Dim LanguageCode_From As String
Dim LanguageCode_To As String
Dim DetectedLanguageCode_From As String

Const REQUEST_INTERVAL_TIME As Integer = 2000
'Change 2000 since Ver 3.4, for seems every 8 ~ 12 Requests, the server will reject new requests for Error of:
'{"error":{"code":429001,"message":"The server rejected the request because the client has exceeded request limits."}}
'This Const was not used since V3.5, for the time interval was done by the function but not script.

Const MAX_CHAR_LIMIT_PER_REQUEST As Integer = 8000
'Max characters length limit of Microsoft Translator is 3000, Actual value set at 2900 for buffer.


'Declaring Variables for Long Text Segment
Dim i As Integer
Dim nNrSegmentLongText As Integer

Dim nSegementStartPos As Integer

Dim LongTextSegmented() As String
Dim TranslatedLongTextSegmented() As String

Dim LongTextSegmented1 As String

Dim TranslatedLongTextCombined As String

Dim TranslatedText As String
Dim TranslatedTextLineBreaked As String

Dim FlagforSaveWorkbook As Integer
FlagforSaveWorkbook = 60


Dim IfAbortTranslating As String

Dim RangeforTranslate As Range
Set RangeforTranslate = Selection


Dim CountofRows As Integer
Dim RowsleftToBeTranslated As Integer

CountofRows = RangeforTranslate.Rows.Count
RowsleftToBeTranslated = CountofRows

With ActiveSheet.UsedRange

    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin

    .WrapText = False

    .Font.Name = "Tahoma"
    .Font.Size = 9
    
End With


'------------------------------------------------------
'Check if Selected Translation Text if Empty'

Dim IfTranlateTextRangeIsEmpty As Boolean

IfTranlateTextRangeIsEmpty = True

For Each RangeforTranslate In RangeforTranslate

    If IsEmpty(RangeforTranslate.Offset(0, -3)) = False Then
    
        IfTranlateTextRangeIsEmpty = False
        Exit For
        
    End If
    
Next RangeforTranslate

If IfTranlateTextRangeIsEmpty = True Then
    Call MsgBox("Please Select Correct Column." & vbCrLf & vbCrLf & "Column of Text to be Translated and Column of LanguageCode_From" & vbCrLf & "shall be on the Left of the Selected Column.", vbOKOnly, "Wrong Range Selected")
    Exit Sub
End If

'------------------------------------------------------
'Define Function

For Each RangeforTranslate In Selection

With RangeforTranslate
    '.Select
    .Value = ""
    .Interior.Color = RGB(191, 191, 191)
    .Font.Color = RGB(0, 0, 0)

'------------------------------------------------------
'Check if select the wrong range and check if Language code is valid

If Len(.Offset(0, -3).Value) < 2 Then
    
    .Offset(0, 1).Value = .Offset(0, -3).Value
    .Value = "No Text for Translation"

    'Call MsgBox("Please Select Correct Column." & vbCrLf & vbCrLf & "Column of Text to be Translated and Column of LanguageCode_From" & vbCrLf & "shall be on the Left of the Selected Column.", vbOKOnly, "Wrong Range Selected")
    'Exit Sub
    'Not Exit Sub since Ver3.3, display a Message in this Cell and continue with Next Cells
    
ElseIf Len(.Offset(0, -2).Value) > 2 Then
    'Try use enumeration variable to list all legal Langua Code in Next Version
    .Value = "Wrong Language Code"
    .Interior.Color = RGB(255, 217, 102)

    'Call MsgBox("Please Select Correct Language Code or leave it Empty", vbOKOnly, "Wrong Language Code")
    'Exit Sub
    'Not Exit Sub since Ver3.3, display a Message in this Cell and continue with Next Cells
    
Else
'------------------------------------------------------
'Text for translation and language Code is ok, Start Translation procedure then.

    LanguageCode_From = .Offset(0, -2).Value
    LanguageCode_To = .Offset(0, -1).Value

    UnTranslatedTextOrginal = .Offset(0, -3).Value

    UnTranslatedTextOrginal = Replace(Expression:=UnTranslatedTextOrginal, Find:="\t", Replace:="vbCrlf")

    'UnTranslatedTextOrginal = Replace(Expression:=UnTranslatedTextOrginal, Find:="### TBD...", Replace:="TBD...")
    
    '4. Use Regular Expression for better Duplicated / Redudant Text Remove, e.g. Safety Questions, Empty Lines, etc...
        
    If InStr(1, UnTranslatedTextOrginal, """") > 0 Or InStr(1, UnTranslatedTextOrginal, "\") > 0 Or InStr(1, UnTranslatedTextOrginal, "/") > 0 Then

        'Seems Microsoft Translator V3.0 does not accept certain characters like ", \ or /, etc...
        UnTranslatedTextSubstitued = Replace(Expression:=UnTranslatedTextOrginal, Find:="""", Replace:="'")

        UnTranslatedTextSubstitued = Replace(Expression:=UnTranslatedTextSubstitued, Find:="\", Replace:="_")

        UnTranslatedTextSubstitued = Replace(Expression:=UnTranslatedTextSubstitued, Find:="/", Replace:="_")
'        Debug.Print UnTranslatedTextSubstitued
        'UnTranslatedTextSubstitued = Replace(Expression:=UnTranslatedTextSubstitued, Find:="ReservedCharacter", Replace:="_")

    Else
    
        UnTranslatedTextSubstitued = UnTranslatedTextOrginal
        
    End If


'------------------------------------------------------

    If Len(UnTranslatedTextSubstitued) < MAX_CHAR_LIMIT_PER_REQUEST Then
    
        TranslatedText = MSTranslate(UnTranslatedTextSubstitued, LanguageCode_From, LanguageCode_To)
'        Debug.Print TranslatedText
        DoEvents
        'Sleep in MSTranslator Function depending on the length of the Text, No sleep in Script since V3.5

'------------------------------------------------------
'This Segement is for long Text that exceed the 10000 Characters Limit for each Request

    ElseIf Len(UnTranslatedTextSubstitued) < (MAX_CHAR_LIMIT_PER_REQUEST * 2) Then

        nNrSegmentLongText = WorksheetFunction.RoundUp((Len(UnTranslatedTextSubstitued) / MAX_CHAR_LIMIT_PER_REQUEST), 0)
        'Debug.Print nNrSegmentLongText

        ReDim LongTextSegmented(nNrSegmentLongText)
        ReDim TranslatedLongTextSegmented(nNrSegmentLongText)
        
        TranslatedLongTextCombined = ""
        
        For i = 0 To nNrSegmentLongText - 1
        
            nSegementStartPos = MAX_CHAR_LIMIT_PER_REQUEST * i
            
            If nSegementStartPos = 0 Then nSegementStartPos = 1
            
            LongTextSegmented(i) = VBA.Mid(UnTranslatedTextSubstitued, nSegementStartPos, MAX_CHAR_LIMIT_PER_REQUEST)
            'Debug.Print LongTextSegmented(i)
            TranslatedLongTextSegmented(i) = MSTranslate(LongTextSegmented(i), LanguageCode_From, LanguageCode_To)
            
            DoEvents

            'Debug.Print TranslatedLongTextSegmented(i)

            'For long text, Wait for 2 x Interval Time for each segment to avoid Error of:
            '{"error":{"code":429001,"message":"The server rejected the request because the client has exceeded request limits."}}
            'Sleep in MSTranslator Function depending on the length of the Text, No sleep in Script since V3.5
            'Add judge 42900# error in Function, and wait for 8 seconds and retry
            
            TranslatedLongTextCombined = TranslatedLongTextCombined & " " & TranslatedLongTextSegmented(i)
        
        Next i

        TranslatedText = TranslatedLongTextCombined


    Else
    ' IF the text is too long (> 16000), then just translate the 1st 500 Characters and the rest keep not translated.

        LongTextSegmented1 = VBA.Mid(UnTranslatedTextSubstitued, 1, 500)

        TranslatedText = MSTranslate(LongTextSegmented1, LanguageCode_From, LanguageCode_To)
        DoEvents
        
        TranslatedLongTextCombined = TranslatedText & vbCrLf & "---------------------" & vbCrLf & _
            "TEXT TOO LONG, Only the 1st 500 Character Translated " & vbCrLf & "---ORIGINAL TEXT Attached Below---" & vbCrLf & _
                vbCrLf & TranslatedLongTextSegmented(i)
        
        TranslatedText = TranslatedLongTextCombined
        
    End If
    
    TranslatedTextLineBreaked = Replace(Expression:=TranslatedText, Find:="\t", Replace:=vbCrLf)
    TranslatedTextLineBreaked = Replace(Expression:=TranslatedTextLineBreaked, Find:="\n", Replace:=vbCrLf)


    If InStr(1, TranslatedText, "Language Detected:") > 0 Then

        DetectedLanguageCode_From = VBA.Mid(TranslatedText, InStr(1, TranslatedText, "Language Detected:") + 18, 2)
        
        .Offset(0, 2).Value = DetectedLanguageCode_From 'Copy Detected LanguageCode_From to 2nd Column to the right
        .Offset(0, 1).Value = VBA.Mid(TranslatedTextLineBreaked, 49) ' Copy the Text to 1st Column to the right

'        Debug.Print TranslatedTextLineBreaked
        
    Else
    
        .Offset(0, 1).Value = TranslatedTextLineBreaked ' Copy the Text to 1st Column to the right
'        Debug.Print TranslatedTextLineBreaked
        
    End If
    
    
    .Offset(0, 1).WrapText = False
        
        
    If InStr(1, TranslatedTextLineBreaked, "error"":{""code"":400035") > 0 Then
    
        .Value = "Wrong Language Code"
        .Interior.Color = RGB(255, 217, 102)
    
    ElseIf InStr(1, TranslatedTextLineBreaked, "error"":{""code") > 0 Then
    
        .Value = "Translation Fail"
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)

    Else
        .Value = "Translated and Copied"
        .Interior.Color = RGB(102, 255, 102)
    
    End If
    
    
    FlagforSaveWorkbook = FlagforSaveWorkbook - 1
    
    If FlagforSaveWorkbook <= 1 Then
        ActiveWorkbook.Save
        Application.StatusBar = "Saving WorkBook, Please Wait..."
        
        DoEvents
        Sleep 5000
        FlagforSaveWorkbook = 60
    End If
  
End If

'Sleep 100
'Sleep in MSTranslator Function depending on the length of the Text, No sleep in Script since V3.5

End With

DoEvents

RowsleftToBeTranslated = RowsleftToBeTranslated - 1

Application.StatusBar = "Translating... Please Wait...   " & RowsleftToBeTranslated & "of Rows left."

Next RangeforTranslate

If CountofRows > 10 Then

    Call MsgBox(CountofRows & " lines of Text Translation Finished", vbOKOnly, "Translation Finished")
    ActiveWorkbook.Save
    
End If

Application.StatusBar = False

End Sub


' Ver 3.5, 19-Jan-2021
' Latest Update:
            '1. Improve the Strategy both check Nr. of Requests and Total Length of Chrs to avoid be Rejected but not only wait for 2000, and also for When to Save Workbook
            '2. Add Status Bar and Process Bar to display more information
' To be Update:
            '1. Clean the error value before Translating (e.g. "#NAME", "#VALUE", etc..)
            '2. Use Regular Expression for better Duplicated / Redudant Text Remove, e.g. Safety Questions, Empty Lines, etc...
            

' Ver 3.4, 12-Jan-2021
' Latest Update: Change MAX_CHAR_LIMIT_PER_REQUEST to 2900; Disable Warp Text for each row; change the interval time from 200 to 2000.
' TobeUpdate:
            '1. Improve the Strategy both check Nr. of Requests and Total Length of Chrs to avoid be Rejected but not only wait for 2000, and also for When to Save Workbook
            '2. Add Status Bar and Process Bar to display more information
            '3. Clean the error value before Translating (e.g. "#NAME", "#VALUE", etc..)
            '4. Use Regular Expression for better Duplicated / Redudant Text Remove, e.g. Safety Questions, Empty Lines, etc...
            

' Ver 3.3, 14-Dec-2020
' Inputbox for InputAuthenticationCode now works since Ver 3.3, but was disabled that not used.
' TobeUpdate:

' Ver 3.2, 09-Dec-2020
' Latest Update: Not allowed to translated in Wrong Cell or Range, and if with wrong language Code
' TobeUpdate: Inputbox for InputAuthenticationCode, not work with Error:This method cannot be called after the sent method has been called





