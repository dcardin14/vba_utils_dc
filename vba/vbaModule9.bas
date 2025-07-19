Attribute VB_Name = "Module9"
Sub clearAllFilters()
Attribute clearAllFilters.VB_Description = "So that I'm not searching in a filtered range accidentally."
Attribute clearAllFilters.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' clearAllFilters Macro
' So that I'm not searching in a filtered range accidentally.
'
' Keyboard Shortcut: Ctrl+Shift+A
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Loop through all tables (ListObjects) in the worksheet
    For Each tbl In ws.ListObjects
        If tbl.ShowAutoFilter Then
            tbl.AutoFilter.ShowAllData
        End If
    Next tbl
    
    ' Check if there are any AutoFilters in the worksheet that are not part of a table
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
End Sub
