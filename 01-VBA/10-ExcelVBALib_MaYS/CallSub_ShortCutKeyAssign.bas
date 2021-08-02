Sub ShortCutKeysAssignment()

Application.OnKey "^+{C}", "ConvertToComment"
Application.OnKey "^+{R}", "ResetCommentsPosition"
Application.OnKey "^+{T}", "MSTranslatorScript"
Application.OnKey "^+{G}", "GoogleTransAutoScriptPy"
Application.OnKey "^+{V}", "BOMStructureView"
Application.OnKey "^+{I}", "IndexAllSheets"

'Application.OnKey "{RIGHT}", ""
'Application.OnKey "{RIGHT}", ""
'Application.OnKey "{RIGHT}", ""
'Application.OnKey "{RIGHT}", ""
'Application.OnKey "{RIGHT}", ""

End Sub
