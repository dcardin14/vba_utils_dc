Attribute VB_Name = "Module4"
Sub turngray()
Attribute turngray.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' turngray Macro
'
' Keyboard Shortcut: Ctrl+Shift+G
' Check if more than 20 cells are selected
    If Selection.Count > 20 Then
        MsgBox "Please select 20 or fewer cells before running this macro.", vbExclamation
        Exit Sub
    End If
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
End Sub
Sub turnyellow()
Attribute turnyellow.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' turnyellow Macro
'
' Keyboard Shortcut: Ctrl+Shift+Y
' Check if more than 20 cells are selected
    If Selection.Count > 20 Then
        MsgBox "Please select 20 or fewer cells before running this macro.", vbExclamation
        Exit Sub
    End If
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
