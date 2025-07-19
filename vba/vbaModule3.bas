Attribute VB_Name = "Module3"
Sub MoveRightAndClear()
Attribute MoveRightAndClear.VB_ProcData.VB_Invoke_Func = "m\n14"
    ' Get the current column index
    Dim CurrentCol As Integer
    CurrentCol = ActiveCell.Column
    
    ' Define the range to copy (from current column to column DO)
    Dim CopyRange As Range
    Set CopyRange = Range(Cells(1, CurrentCol), Cells(306, Range("DO1").Column))
    
    ' Copy the data
    CopyRange.Copy
    
    ' Move one column to the right
    ActiveCell.Offset(0, 1).Select
    
    ' Paste the copied data
    ActiveSheet.Paste
    
    ' Clear only the original column's contents up to row 306
    Range(Cells(1, CurrentCol), Cells(306, CurrentCol)).ClearContents
End Sub
