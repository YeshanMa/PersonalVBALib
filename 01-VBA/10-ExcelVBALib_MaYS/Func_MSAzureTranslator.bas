Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Function MSTranslate(sText As String, sLanguageFrom As String, Optional sLanguageTo As String)

Const AUTHENTICATION_FAIL As String = "credentials are missing"
Const TRANSLATION_ERROR As String = "error"":{""code"":40"
Const TRANSLATION_EXCEED_REQUEST_LIMIT As String = "error"":{""code"":429"
Const TRANSLATION_DETECTION As String = "[{""detectedLanguage"":{""language"":"
Const TRANSLATION_TEXT_DETECT As String = "},""translations"":[{""text"":"""
Const TRANSLATION_DETECTION_ENDOFTEXTFLAG As String = ",""to"":"

'Example of Response Text
'{"error":{"code":400000,"message":"One of the request inputs is not valid."}}
'[{"detectedLanguage":{"language":"it","score":1.0},"translations":[{"text":"TEXT TRANSLATED AS EXAMPLE","to":"en"}]}]

'---Declare Variables---
Dim sHostUrl As String
Dim sRegion As String

Static sAuthenticationKey As String
Dim sNewInputKey As String

Dim sTextforPOST As String

Dim startpl, endpl As Integer
Dim sRawTranslatedsText As String
Dim nLengthofRawTranslatedsText As Long

Dim sDetectedLanguageCode As String

If sLanguageTo = "" Then sLanguageTo = "en"

'---Sub Procedure Start---

If Len(sText) > 0 Then 'if blank do nothing return the blank value

    If InStr(1, sText, """") > 0 Or InStr(1, sText, "\") > 0 Or InStr(1, sText, "/") > 0 Then
        sText = Replace(Expression:=sText, Find:="""", Replace:="'")
        sText = Replace(Expression:=sText, Find:="\", Replace:="_")
        sText = Replace(Expression:=sText, Find:="/", Replace:="_")
        'sText = Replace(Expression:=sText, Find:="ReservedCharacter", Replace:="_")
    End If
    
    sRegion = "eastus"
    sAuthenticationKey = "7d53c79bfbf044339712e46c579fae96" 'authentication Key from subscription, used on 09Dec, @outlook.com
    
    If Len(sAuthenticationKey) < 10 Then
    sAuthenticationKey = InputBox("Please input Authentication Code", "Input Authentication Code")
    End If
    
    sHostUrl = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0" 'required link for authentication
    sHostUrl = sHostUrl & "&from=" & sLanguageFrom & "&to=" & sLanguageTo 'determine languagefrom and langauge to

    sTextforPOST = "[{""text"":" & """" & sText & """}]" 'JSON format spcific requirement [{"text":"value"}] max5000 characters
    'JSON format spcific requirement [{"text":"value"}] max5000 characters

RetryRequestForLimitExceedOrAuthenticationFail:

    Dim objTextforCognitive As Object
    Set objTextforCognitive = CreateObject("WinHttp.WinHttpRequest.5.1") 'need to add reference libary "Microsft WinHTTP Service,Version 5.1"
    
    With objTextforCognitive
        .Open "POST", sHostUrl, False 'open connection to "Translator Text API" POST command required
        
        .setRequestHeader "Ocp-Apim-Subscription-Key", sAuthenticationKey 'Authentication Required
        .setRequestHeader "Ocp-Apim-Subscription-Region", sRegion 'Subscription-Region for Multi-Region Service

        .setRequestHeader "Content-type", "Application/json" 'Content-type Required
        .setRequestHeader "Accept", "application/json" 'Accept not Required for Translator

        .send sTextforPOST 'format = [{"text":"Hello, Please help to Translate this text"}]
'        Debug.Print sTextforPOST

        .waitForResponse 'theresponse takes 1+ seconds needs wait or delay command or results will failas response has not returned data yet
        'Debug.Print .GetAllResponseHeaders
        DoEvents
        
        sRawTranslatedsText = .responseText
'        Debug.Print sRawTranslatedsText
    End With

    If InStr(1, sRawTranslatedsText, AUTHENTICATION_FAIL) > 0 Then
    '{"error":{"code":401000,"message":"The request is not authorized because credentials are missing or invalid."}}"

        sNewInputKey = InputBox("Authentication Fail.." & vbCrLf & "Please input Authentication Code", "Authentication Fail")

            If Len(sNewInputKey) < 1 Then
                Call MsgBox("Wrong Authentication Code, Translation Abort, Please Retry", vbOKOnly, "Authentication Fail")
                MSTranslate = "Authentication Fail, ERROR CODE:" & vbCrLf & sRawTranslatedsText
                
                Exit Function
            End If

        sAuthenticationKey = sNewInputKey
        GoTo RetryRequestForLimitExceedOrAuthenticationFail

    End If


    If InStr(1, sRawTranslatedsText, TRANSLATION_ERROR) > 0 Then

        MSTranslate = "TRANSLATION FAIL, ERROR CODE:" & vbCrLf & sRawTranslatedsText

    '-----------------------------------------------------
    'For Error: 42900#
    '{"error":{"code":429001,"message":"The server rejected the request because the client has exceeded request limits."}}"
    'Wait for 8 seconds and retry request

    ElseIf InStr(1, sRawTranslatedsText, TRANSLATION_EXCEED_REQUEST_LIMIT) > 0 Then

        Sleep 8000
        
        Application.StatusBar = "Wait for 8 Seconds for Error 42900#: The server rejected the request for request limits exceeded"

        GoTo RetryRequestForLimitExceedOrAuthenticationFail
    
    ElseIf InStr(1, sRawTranslatedsText, TRANSLATION_DETECTION) > 0 Then

        'if Language Code is empty, them the translator will detect language and response text as below
        startpl = InStr(1, sRawTranslatedsText, TRANSLATION_TEXT_DETECT) + 27
        endpl = InStr(startpl, sRawTranslatedsText, TRANSLATION_DETECTION_ENDOFTEXTFLAG)
        '[{"translations":[{"text":"Hellouser","to":"en"}]}]
        'MSTranslate = VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated text
              
        MSTranslate = sRawTranslatedsText 'Parse out translated text
        sDetectedLanguageCode = VBA.Mid(sRawTranslatedsText, 35, 2)
        
    
        'startpl = 77 'if Language Code is empty, them the translator will detect language and response text as below, use start position from 77
         '[{"detectedLanguage":{"language":"it","score":1.0},"translations":[{"text":"TEXT TRANSLATED","to":"en"}]}]
        'endpl = InStr(startpl, sRawTranslatedsText, """") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        sDetectedLanguageCode = VBA.Mid(sRawTranslatedsText, 35, 2)

        MSTranslate = "Languge Code Not Defined, Language Detected:" & sDetectedLanguageCode & vbCrLf & VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated tex


    Else 'The translation was successfully done without error.
    
        startpl = 28 'Successful Translation response with Text body from characters count 28
        endpl = InStr(startpl, sRawTranslatedsText, ",""to"":") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        MSTranslate = VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated tex
        'MSTranslate = sRawTranslatedsText

    End If
        objTextforCognitive.abort
        Set objTextforCognitive = Nothing

    nLengthofRawTranslatedsText = Len(sRawTranslatedsText)

    If nLengthofRawTranslatedsText < 200 Then nLengthofRawTranslatedsText = 200

    Sleep (nLengthofRawTranslatedsText / 2)
    'wait for Milliseconds / Characters, e.g. 4000 ms for 8000 Characters, minimus 200ms.

Else

    MSTranslate = sText 'if blank do nothing return the blankvalue

    
End If

Application.StatusBar = False

End Function

