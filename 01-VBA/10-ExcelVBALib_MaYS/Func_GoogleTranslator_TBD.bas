Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Function GoogleTranslate(sText As String, sLanguageFrom As String)

'Attribute VB_Name = "Func_GoogleTranslator_TBD"
'GoogleTranslator must call Python modules, TBD..

Const TRANSLATION_ERROR As String = "error"":{""code"
Const TRANSLATION_DETECTION As String = "[{""detectedLanguage"":{""language"":"
Const TRANSLATION_TEXT_DETECT As String = "},""translations"":[{""text"":"""

'Example of Response Text
'{"error":{"code":400000,"message":"One of the request inputs is not valid."}}
'[{"detectedLanguage":{"language":"it","score":1.0},"translations":[{"text":"TEXT TRANSLATED","to":"en"}]}]

'---Declare Variables---
Dim sHostUrl As String
Dim sRegion As String

Static sAuthenticationKey As String
Dim sNewInputKey As String

Dim sTextforPOST As String

Dim startpl, endpl As Integer
Dim sRawTranslatedsText

'---Sub Procedure Start---

If Len(sText) > 0 Then 'if blank do nothing return the blank value

    If InStr(1, sText, """") > 0 Or InStr(1, sText, "\") > 0 Or InStr(1, sText, "/") > 0 Then
        sText = Replace(Expression:=sText, Find:="""", Replace:="'")
        sText = Replace(Expression:=sText, Find:="\", Replace:="_")
        sText = Replace(Expression:=sText, Find:="/", Replace:="_")
        'sText = Replace(Expression:=sText, Find:="ReservedCharacter", Replace:="_")
    End If
    
    
    'sRegion = "eastus"
    'sAuthenticationKey = "7d53c79bfbf044339712e46c579fae96" 'authentication Key from subscription, used on 09Dec, @outlook.com
    
    'If Len(sAuthenticationKey) < 10 Then
    'sAuthenticationKey = InputBox("Please input Authentication Code", "Input Authentication Code")
    'End If
    
    sHostUrl = "https://translate.google.cn/_/TranslateWebserverUi/data/batchexecute" 'required link for authentication
    'sHostUrl = sHostUrl & "&from=" & sLanguageFrom & "&to=en" 'determine languagefrom and langauge to

    'sTextforPOST = "[{""text"":" & """" & sText & """}]" 'JSON format spcific requirement [{"text":"value"}] max5000 characters
    sTextforPOST = "%5B%5B%5B%22MkEWBc%22%2C%22%5B%5B%5C%22Vielen%20vielen%20Dank%5C%22%2C%5C%22auto%5C%22%2C%5C%22en%5C%22%2Ctrue%5D%2C%5B1%5D%5D%22%2Cnull%2C%22generic%22%5D%5D%5D&"
    'JSON format spcific requirement [{"text":"value"}] max5000 characters

InputAuthenticationCode:
    Dim objTextforCognitive As Object
    Set objTextforCognitive = CreateObject("WinHttp.WinHttpRequest.5.1") 'need to add reference libary "Microsft WinHTTP Service,Version 5.1"
    
    With objTextforCognitive
        .Open "POST", sHostUrl, False 'open connection to "Translator Text API" POST command required
        
        .setRequestHeader "Referer", "http://translate.google.cn" 'Authentication Required
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.106 Safari/537.36"
        .setRequestHeader "Content-type", "application/x-www-form-urlencoded;charset=utf-8" 'Content-type Required
        '.setRequestHeader "Accept", "application/json" 'Accept not Required for Translator

        .send sTextforPOST 'format = [{"text":"Hello, Please help to Translate this text"}]
        Debug.Print sTextforPOST

        .waitForResponse 'theresponse takes 1+ seconds needs wait or delay command or results will failas response has not returned data yet
        Debug.Print .GetAllResponseHeaders
        DoEvents
        Sleep 200
        
        sRawTranslatedsText = .responseText
        'Debug.Print sRawTranslatedsText
    End With

    If InStr(1, sRawTranslatedsText, "credentials are missing") > 0 Then
    '{"error":{"code":401000,"message":"The request is not authorized because credentials are missing or invalid."}}"

        sNewInputKey = InputBox("Authentication Fail.." & vbCrLf & "Please input Authentication Code", "Authentication Fail")

            If Len(sNewInputKey) < 1 Then
                Call MsgBox("Wrong Authentication Code, Translation Abort, Please Retry", vbOKOnly, "Authentication Fail")
                GoogleTranslate = "Authentication Fail, ERROR CODE:" & vbCrLf & sRawTranslatedsText
                
                Exit Function
            End If

        sAuthenticationKey = sNewInputKey
        GoTo InputAuthenticationCode

    End If

        
    If InStr(1, sRawTranslatedsText, TRANSLATION_ERROR) > 0 Then
        GoogleTranslate = "TRANSLATION FAIL, ERROR CODE:" & vbCrLf & sRawTranslatedsText

    
    ElseIf InStr(1, sRawTranslatedsText, TRANSLATION_DETECTION) > 0 Then

        'if Language Code is empty, them the translator will detect language and response text as below
        startpl = InStr(1, sRawTranslatedsText, TRANSLATION_TEXT_DETECT) + 27
        endpl = InStr(startpl, sRawTranslatedsText, ",""to"":") '[{"translations":[{"text":"Hellouser","to":"en"}]}]
        'GoogleTranslate = VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated text
              
        GoogleTranslate = sRawTranslatedsText 'Parse out translated text
        sDetectedLanguageCode = VBA.Mid(sRawTranslatedsText, 35, 2)
        
    
        'startpl = 77 'if Language Code is empty, them the translator will detect language and response text as below, use start position from 77
         '[{"detectedLanguage":{"language":"it","score":1.0},"translations":[{"text":"TEXT TRANSLATED","to":"en"}]}]
        'endpl = InStr(startpl, sRawTranslatedsText, """") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        sDetectedLanguageCode = VBA.Mid(sRawTranslatedsText, 35, 2)

        GoogleTranslate = "Languge Code Not Defined, Language Detected:" & sDetectedLanguageCode & vbCrLf & VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated tex


    Else 'The translation was successfully done without error.
    
        startpl = 28 'Successful Translation response with Text body from characters count 28
        endpl = InStr(startpl, sRawTranslatedsText, ",""to"":") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        GoogleTranslate = VBA.Mid(sRawTranslatedsText, startpl, endpl - startpl - 1) 'Parse out translated tex
        'GoogleTranslate = sRawTranslatedsText

    End If
    
        objTextforCognitive.abort
Else

    GoogleTranslate = sText 'if blank do nothing return the blankvalue
    
End If

End Function




    ����          �<$���fv&�s��̍O��           !   0   ?   Q   c   u   �   �   �   2.23.140.1.2.2 1.3.6.1.5.5.7.3.1 2.23.140.1.2.1 2.23.140.1.2.2 1.3.6.1.5.5.7.3.1 1.3.6.1