Attribute VB_Name = "Module12"
Sub hideHistory()
Attribute hideHistory.VB_Description = "Hides all the prior conveyances and gives me a recap I can work from."
Attribute hideHistory.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' hideHistory Macro
' Hides all the prior conveyances and gives me a recap I can work from.
'
' Keyboard Shortcut: Ctrl+Shift+H
'
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim targetCol As Long
    Dim r As Long

    Set ws = ActiveSheet

    ' Step 1: Find rightmost non-empty cell in range B1:DM1
    For col = 118 To 2 Step -1 ' Columns DM (118) to B (2)
        If Not IsEmpty(ws.Cells(1, col).Value) Then
            lastCol = col
            Exit For
        End If
    Next col

    ' Step 2: Hide columns from column 2 to column n (inclusive)
    If lastCol >= 2 Then
        ws.Range(ws.Columns(2), ws.Columns(lastCol)).EntireColumn.Hidden = True
    End If

    ' Step 3: Set headers in row 1 and 2 of the target column
    targetCol = lastCol + 1
    ws.Cells(1, targetCol).Value = "RECAP"
    ws.Cells(2, targetCol).Value = "(autogen)"

    ' Step 4: Copy date from row 3 of lastCol to row 3 of targetCol
    ws.Cells(3, targetCol).Value = ws.Cells(3, lastCol).Value

    ' Step 5: Copy values from column 124 (DT) to column n+1 for rows 4 to 305
    For r = 4 To 305
        ws.Cells(r, targetCol).Value = ws.Cells(r, 124).Value
    Next r
End Sub
Sub exportAllMacros()
Attribute exportAllMacros.VB_ProcData.VB_Invoke_Func = " \n14"
'
' exportAllMacros Macro
'

'
End Sub
Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String

    exportPath = "C:\code\vba_utils_dc\vba"

    If Dir(exportPath, vbDirectory) = "" Then
        MsgBox "Export folder does not exist: " & exportPath, vbCritical
        Exit Sub
    End If

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' Module, Class Module, UserForm
                vbComp.Export exportPath & vbComp.Name & GetExtension(vbComp.Type)
        End Select
    Next vbComp
    MsgBox "Export complete!"
End Sub

Function GetExtension(vbType As Long) As String
    Select Case vbType
        Case 1: GetExtension = ".bas"
        Case 2: GetExtension = ".cls"
        Case 3: GetExtension = ".frm"
        Case Else: GetExtension = ""
    End Select
End Function

