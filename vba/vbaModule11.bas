Attribute VB_Name = "Module11"
Sub ShowCurrentOwnership()
Attribute ShowCurrentOwnership.VB_Description = "shows the current owners and their interests in my chain sheet (from Lisa Morriss)"
Attribute ShowCurrentOwnership.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' ShowCurrentOwnership Macro
' shows the current owners and their interests in my chain sheet (from Lisa Morriss)
'
' Keyboard Shortcut: Ctrl+p

    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long, outputRow As Long
    Dim ownerName As String
    Dim interestVal As Double
    Dim tblRange As Range
    Dim tbl As ListObject

    ' Use the active sheet as the source
    Set wsSource = ActiveSheet

    ' Delete existing 'CurrentOwners' sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("CurrentOwners").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new output sheet
    Set wsOutput = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsOutput.Name = "CurrentOwners"

    ' Determine last used row in column A
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' Write headers
    wsOutput.Range("A1").Value = "Owner"
    wsOutput.Range("B1").Value = "Interest"
    outputRow = 2

    ' Loop through and collect nonzero interests
    For i = 2 To lastRow
        ownerName = wsSource.Cells(i, 1).Value
        interestVal = Val(wsSource.Cells(i, 121).Value) ' Column DQ

        If interestVal > 0 Then
            wsOutput.Cells(outputRow, 1).Value = ownerName
            wsOutput.Cells(outputRow, 2).Value = interestVal
            outputRow = outputRow + 1
        End If
    Next i

    ' Sort descending by interest
    With wsOutput.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsOutput.Range("B2:B" & outputRow - 1), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange wsOutput.Range("A1:B" & outputRow - 1)
        .Header = xlYes
        .Apply
    End With

    ' Format as a table
    Set tblRange = wsOutput.Range("A1:B" & outputRow - 1)
    Set tbl = wsOutput.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "SortedOwnerTable"
    tbl.TableStyle = "TableStyleMedium15"
    tbl.ShowTotals = True

    ' Set formatting
    wsOutput.Columns("A").ColumnWidth = 36
    wsOutput.Columns("B").NumberFormat = "0.000000%"
    wsOutput.Columns("B").ColumnWidth = 12

   '  MsgBox "Done! Owners with nonzero interest are listed and formatted in 'CurrentOwners'.", vbInformation


'
End Sub
