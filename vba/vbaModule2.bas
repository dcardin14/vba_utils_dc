Attribute VB_Name = "Module2"
Sub turn_red()
Attribute turn_red.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' turn_red Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    With Selection.Font
        .Color = -16777024
        .TintAndShade = 0
    End With
End Sub
Sub mwshssp()
Attribute mwshssp.VB_ProcData.VB_Invoke_Func = "Z\n14"
'
' mwshssp Macro
'
' Keyboard Shortcut: Ctrl+Shift+Z
'
Set selectedCell = ActiveCell
    
    ' Check if a cell is selected
    If Not selectedCell Is Nothing Then
        selectedCell.Value = selectedCell.Value & " a married woman dealing in her sole and separate property"
    Else
        MsgBox "Please select a cell to update."
    End If
End Sub

Sub mmssp()
Attribute mmssp.VB_ProcData.VB_Invoke_Func = "X\n14"
'
' mwshssp Macro
'
' Keyboard Shortcut: Ctrl+Shift+X
'
Set selectedCell = ActiveCell
    
    ' Check if a cell is selected
    If Not selectedCell Is Nothing Then
        selectedCell.Value = selectedCell.Value & " a married man dealing in his sole and separate property"
    Else
        MsgBox "Please select a cell to update."
    End If
End Sub
