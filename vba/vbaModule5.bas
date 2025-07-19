Attribute VB_Name = "Module5"
Sub shade()
Attribute shade.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' shade Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
    Dim selectedRange As Range
    Dim cell As Range

    ' Check if more than 20 cells are selected
    If Selection.Count > 20 Then
        MsgBox "Please select 20 or fewer cells before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Check if there is a selection
    If Selection Is Nothing Then
        MsgBox "No cells are selected."
        Exit Sub
    End If
    
    ' Store the current selection
    Set selectedRange = Selection

    ' Loop through each cell in the selection
    For Each cell In selectedRange
        ' Expand the selection to the entire row of the current cell
        cell.EntireRow.Select
    Next cell
    
    ' Apply shading
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    ' Restore the original selection
    selectedRange.Select

End Sub
Sub green()
Attribute green.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' green Macro
'
' Keyboard Shortcut: Ctrl+Shift+D
    ' Check if more than 20 cells are selected
    If Selection.Count > 20 Then
        MsgBox "Please select 20 or fewer cells before running this macro.", vbExclamation
        Exit Sub
    End If
    
Dim selectedRange As Range
    Dim cell As Range
    
    ' Check if there is a selection
    If Selection Is Nothing Then
        MsgBox "No cells are selected."
        Exit Sub
    End If
    
    ' Store the current selection
    Set selectedRange = Selection
    
    ' Loop through each cell in the selection
    For Each cell In selectedRange
        ' Expand the selection to the entire row of the current cell
        cell.EntireRow.Select
    Next cell
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub PasswordBreaker()

'Breaks worksheet password protection.

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer
On Error Resume Next
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126

ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

If ActiveSheet.ProtectContents = False Then
MsgBox "One usable password is " & Chr(i) & Chr(j) & _
Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
Exit Sub
End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub

