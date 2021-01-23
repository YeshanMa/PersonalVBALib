'Attribute VB_Name = "Func_MSKeyPhraseExtrac"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Function MSKeyPhraseExtract(sText As String)

Dim sHostUrl As String
Dim sRegion As String

Static sAuthenticationKey As String
Dim sNewInputKey As String

Dim sTextforPOST As String

Dim DetectedLanguageCode As String
     
Dim startpl, endpl As Integer
Dim sRawTextMSKeyPhraseExtract

If Len(sText) > 0 Then 'if blank do nothing return the blank value

    sRegion = "eastus"
    sAuthenticationKey = "05c1c6fe90d7465c96f188397705068a" 'authentication Key from subscription, used on 09Dec, @outlook.com
    
    If Len(sAuthenticationKey) < 10 Then
    sAuthenticationKey = InputBox("Please input Authentication Code", "Input Authentication Code")
    End If
    
    sHostUrl = "https://eastus.api.cognitive.microsoft.com/text/analytics/v2.1/keyPhrases?model-version={string}&showStats={boolean}" 'required link for authentication
    'sHostUrl = sHostUrl & "&from=" & sLanguageFrom & "&to=en" 'determine languagefrom and langauge to

    sTextforPOST = "{""documents"":[{""language"":""en"",""id"":""1"",""text"": " & """" & sText & """" & "}]}"
    'JSON format spcific requirement

InputAuthenticationCode:
    Dim objTextforCognitive As Object
    Set objTextforCognitive = CreateObject("WinHttp.WinHttpRequest.5.1") 'need to add reference libary "Microsft WinHTTP Service,Version 5.1"
    
    With objTextforCognitive
        .Open "POST", sHostUrl, False 'open connection to "Translator Text API" POST command required
        
        .setRequestHeader "Ocp-Apim-Subscription-Key", sAuthenticationKey 'Authentication Required
        '.setRequestHeader "Ocp-Apim-Subscription-Region", sRegion 'Subscription-Region for Multi-Region Service

        .setRequestHeader "Content-type", "application/json" 'Content-type Required
        .setRequestHeader "Accept", "application/json" 'Accept Required
        
        .send sTextforPOST 'format = [{"text":"Hello, Please help to Translate this text"}]
        'Debug.Print sTextforPOST
        
        .waitForResponse 'theresponse takes 1+ seconds needs wait or delay command or results will failas response has not returned data yet
        'Debug.Print .GetAllResponseHeaders
        DoEvents
        Sleep 200
        
        sRawTextMSKeyPhraseExtract = .responseText
    'Debug.Print sRawTextMSKeyPhraseExtract
    End With

    If InStr(1, sRawTextMSKeyPhraseExtract, "invalid subscription key") > 0 Then
    '{"error":{"code":401000,"message":"The request is not authorized because credentials are missing or invalid."}}"

        sNewInputKey = InputBox("Authentication Fail.." & vbCrLf & "Please input Authentication Code", "Authentication Fail")

            If Len(sNewInputKey) < 1 Then
                Call MsgBox("Wrong Authentication Code, Translation Abort, Please Retry", vbOKOnly, "Authentication Fail")
                MSTranslate = "Authentication Fail, ERROR CODE:" & vbCrLf & sRawTranslatedsText
                
                Exit Function
            End If

        sAuthenticationKey = sNewInputKey
        GoTo InputAuthenticationCode

    End If

    If InStr(1, sRawTextMSKeyPhraseExtract, "InvalidRequest") > 0 Then
        MSKeyPhraseExtract = "Key Phase Extract FAIL, ERROR CODE:" & vbCrLf & sRawTextMSKeyPhraseExtract
        '{"error":{"code":"InvalidRequest","message":"Invalid Request.","inner error":{"code":"EmptyRequest","message":"Request body must be present."}}}
    
    ElseIf InStr(1, sRawTextMSKeyPhraseExtract, """keyPhrases"":[") > 0 Then
    
        startpl = 39 'Error code start from character count 39
         '[{"detectedLanguage":{"language":"it","score":1.0},"translations":[{"text":"TEXT TRANSLATED","to":"en"}]}]
        endpl = InStr(startpl, sRawTextMSKeyPhraseExtract, "]}],") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        sRawTextMSKeyPhraseExtract = Mid(sRawTextMSKeyPhraseExtract, startpl, endpl - startpl) 'Parse out translated tex
        sRawTextMSKeyPhraseExtract = Replace(Expression:=sRawTextMSKeyPhraseExtract, Find:="""", Replace:="")
        
        MSKeyPhraseExtract = Replace(Expression:=sRawTextMSKeyPhraseExtract, Find:=",", Replace:=vbCrLf)

    Else
    
        startpl = 1
        'if you use auto languae detect you will need toadjust this number to "69" or greater
        endpl = InStr(startpl, sRawTextMSKeyPhraseExtract, """") '[{"translations":[{"text":"Hellouser","to":"en"}]}]

        sRawTextMSKeyPhraseExtract = Mid(sRawTextMSKeyPhraseExtract, startpl, endpl - startpl) 'Parse out translated tex
        sRawTextMSKeyPhraseExtract = Replace(Expression:=sRawTextMSKeyPhraseExtract, Find:="""", Replace:="")

        MSKeyPhraseExtract = Replace(Expression:=sRawTextMSKeyPhraseExtract, Find:=",", Replace:=vbCrLf)
        
    End If
    
    objTextforCognitive.abort
Else

    MSKeyPhraseExtract = sText 'if blank do nothing return the blankvalue
    
End If

End Function