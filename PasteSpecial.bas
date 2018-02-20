Attribute VB_Name = "PasteSpecial"
Sub PasteSpecialValues()
Attribute PasteSpecialValues.VB_ProcData.VB_Invoke_Func = "V\n14"
'Keyboard Shortcut: Ctrl+Shift+V

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
End Sub
Sub PasteSpecialFormulas()
Attribute PasteSpecialFormulas.VB_ProcData.VB_Invoke_Func = "F\n14"
'Keyboard Shortcut: Ctrl+Shift+F
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub
