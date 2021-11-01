
'Ver 1.0.0, 31-Oct-2021
'Latest Update:
        '1. New Created and Copy Source Code from 
        '2. https://analysistabs.com/excel-vba/interacting-with-other-applications-using-vba/

'https://stackoverflow.com/questions/9535959/calling-a-word-vba-sub-with-arguments-from-excel-vba
'https://social.msdn.microsoft.com/forums/en-US/b86c8cd9-26b0-442b-9fa2-406f2d8471e3/how-to-manipulate-word-userform-from-excel-vba-office-2010
Sub CallwdSubSample()

Dim wdApp As Object
Dim wdDoc As Object
Dim bStarted As Boolean
Dim strMyTemplate As String
strMyTemplate = "D:\Word 2010 Templates\Template.dot"

On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
    If Err Then
        Set wdApp = CreateObject("Word.Application")
        bStarted = True
    End If
Set wdDoc = wdApp.Documents.Add(strMyTemplate)

wdApp.Visible = True
wdApp.Activate
' do whatever else you want in Word here -e.g. wdApp.Run YourWordMacro etc.
wdApp.Run ("ShowMyForm")
' and when you're finished
    If bStarted = True Then
        wdApp.Quit
    End If
Set wdDoc = Nothing
Set wdApp = Nothing

End Sub

Sub sbWord_CreatingAndFormatingWordDoc()

Dim oWApp As Word.Application
Dim oWDoc As Word.Document
Dim sText As String
Dim iCntr As Long
'
Set oWApp = New Word.Application
Set oWDoc = oWApp.Documents.Add() '(&amp;quot;C:\Documents\Doc1.dot&amp;quot;) 'You can specify your template here
'
'Adding new Paragraph
'
Dim para As Paragraph
Set para = oWDoc.Paragraphs.Add
'
para.Range.Text = &amp;quot;Paragraph 1 - My Heading: ANALYSISTABS.COM&amp;quot;
para.Format.Alignment = wdAlignParagraphCenter
para.Range.Font.Size = 18
para.Range.Font.Name = &amp;quot;Cambria&amp;quot;
'
For i = 0 To 2
Set para = oWDoc.Paragraphs.Add
para.Space2
Next
'
Set para = oWDoc.Paragraphs.Add
With para
.Range.Text = &amp;quot;Paragraph 2 - Example Paragraph, you can format it as per yor requirement&amp;quot;
.Alignment = wdAlignParagraphLeft
.Format.Space15
.Range.Font.Size = 14
.Range.Font.Bold = True
End With
'
oWDoc.Paragraphs.Add
'
Set para = oWDoc.Paragraphs.Add
With para
.Range.Text = &amp;quot;Paragraph 3 - Another Paragraph, you can create number of paragraphs like this and format it&amp;quot;
.Alignment = wdAlignParagraphLeft
.Format.Space15
.Range.Font.Size = 12
.Range.Font.Bold = False
End With
'
oWApp.Visible = True
End Sub

Sub sbWord_CreatingAndFormatingWordDocLateBinding()

Dim oWApp As Object
Dim oWDoc As Object
Dim sText As String
Dim iCntr As Long
'
Set oWApp = New Word.Application
Set oWDoc = oWApp.Documents.Add() '(&amp;quot;C:\Documents\Doc1.dot&amp;quot;) 'You can specify your template here
'
'Adding new Paragraph
'
Dim para As Paragraph
Set para = oWDoc.Paragraphs.Add
'
para.Range.Text = &amp;quot;Paragraph 1 - My Heading: ANALYSISTABS.COM&amp;quot;
para.Format.Alignment = wdAlignParagraphCenter
para.Range.Font.Size = 18
para.Range.Font.Name = &amp;quot;Cambria&amp;quot;
'
For i = 0 To 2
Set para = oWDoc.Paragraphs.Add
para.Space2
Next
'
Set para = oWDoc.Paragraphs.Add
With para
.Range.Text = &amp;quot;Paragraph 2 - Example Paragraph, you can format it as per yor requirement&amp;quot;
.Alignment = wdAlignParagraphLeft
.Format.Space15
.Range.Font.Size = 14
.Range.Font.Bold = True
End With
'
oWDoc.Paragraphs.Add
'
Set para = oWDoc.Paragraphs.Add
With para
.Range.Text = &amp;quot;Paragraph 3 - Another Paragraph, you can create number of paragraphs like this and format it&amp;quot;
.Alignment = wdAlignParagraphLeft
.Format.Space15
.Range.Font.Size = 12
.Range.Font.Bold = False
End With
'
oWApp.Visible = True
End Sub
