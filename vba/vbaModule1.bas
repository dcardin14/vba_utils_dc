Attribute VB_Name = "turnblack" 'Turns the font black
Sub turnblack()
Attribute turnblack.VB_Description = "Turn font black"
Attribute turnblack.VB_ProcData.VB_Invoke_Func = "B\n14"
'
' turnblack Macro
' Turn font black
'
' Keyboard Shortcut: Ctrl+Shift+B
'
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub
