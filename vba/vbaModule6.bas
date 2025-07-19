Attribute VB_Name = "Module6"
Dim originalData As Variant

Sub shift_timeline_right()
Attribute shift_timeline_right.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim col As Integer
    Dim startColLetter As String
    Dim endColLetter As String
    Dim selectRange As String
    Dim pasteRange As String
    Dim originalColumnRange As String

    ' Get the column number of the currently selected cell
    col = ActiveCell.Column
    
    ' Convert column number to letter (e.g., 1 to "A")
    startColLetter = Split(Cells(1, col).Address, "$")(1)
    
    ' Define the end column letter (DA)
    endColLetter = "DA"
    
    ' Define the range to select (from row 1 to 300 in the identified column to column DA)
    selectRange = startColLetter & "1:" & endColLetter & "300"
    
    ' Store the original data
    originalData = Range(selectRange).Value
    
    ' Select the defined range
    Range(selectRange).Select
    
    ' Copy the selected range to the clipboard
    Selection.Copy
    
    ' Define the range to paste (one column to the right)
    pasteRange = Cells(1, col + 1).Address(False, False) & ":" & Cells(300, col + 27).Address(False, False) ' DA is 26 columns from the start
    
    ' Paste the copied range to the new location
    Range(pasteRange).PasteSpecial Paste:=xlPasteAll
    
    ' Define the range to clear (the first column of the original range)
    originalColumnRange = Cells(1, col).Address(False, False) & ":" & Cells(300, col).Address(False, False)
    
    ' Clear the contents of the first column in the original range
    Range(originalColumnRange).ClearContents
    
    ' Deselect to remove the marching ants (optional)
    Application.CutCopyMode = False
End Sub

Sub undo_shift_timeline_right()
Attribute undo_shift_timeline_right.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' undo_shift_timeline_right Macro
'
' Keyboard Shortcut: Ctrl+z
'
    Dim col As Integer
    Dim startColLetter As String
    Dim endColLetter As String
    Dim selectRange As String
    
    ' Get the column number of the original cell
    col = ActiveCell.Column - 1
    
    ' Convert column number to letter (e.g., 1 to "A")
    startColLetter = Split(Cells(1, col).Address, "$")(1)
    
    ' Define the end column letter (DA)
    endColLetter = "DA"
    
    ' Define the range to select (from row 1 to 300 in the identified column to column DA)
    selectRange = startColLetter & "1:" & endColLetter & "300"
    
    ' Restore the original data
    Range(selectRange).Value = originalData
End Sub
