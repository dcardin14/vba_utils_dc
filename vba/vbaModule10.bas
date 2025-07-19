Attribute VB_Name = "Module10"
Sub RenameTabsAndCreateTOC()
    Dim ws As Worksheet
    Dim originalNames As Collection
    Dim newName As String
    Dim i As Integer
    Dim TOC As Worksheet
    Dim rowIndex As Integer
    
    ' Store the original names
    Set originalNames = New Collection
    
    ' Save the original names and rename sheets
    i = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "TOC" Then
            originalNames.Add ws.Name
            i = i + 1
            newName = Chr(64 + i) ' Convert number to letter (1 -> A, 2 -> B, etc.)
            On Error Resume Next
            ws.Name = newName
            If Err.Number <> 0 Then
                MsgBox "Error renaming sheet: " & ws.Name, vbExclamation, "Error"
                Exit Sub
            End If
            On Error GoTo 0
        End If
    Next ws
    
    ' Create the TOC tab
    On Error Resume Next
    Set TOC = ActiveWorkbook.Worksheets("TOC")
    If TOC Is Nothing Then
        Set TOC = ActiveWorkbook.Worksheets.Add
        TOC.Name = "TOC"
    Else
        TOC.Cells.Clear ' Clear existing content
    End If
    On Error GoTo 0
    TOC.Move Before:=ActiveWorkbook.Worksheets(1)
    ' Populate the TOC tab
    TOC.Cells(1, 1).Value = "Original Name"
    TOC.Cells(1, 2).Value = "New Name"
    rowIndex = 2
    i = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "TOC" Then
            i = i + 1
            newName = Chr(64 + i)
            TOC.Cells(rowIndex, 1).Value = originalNames(i)
            TOC.Cells(rowIndex, 2).Value = newName
            rowIndex = rowIndex + 1
        End If
    Next ws
    
    ' Format the TOC
    With TOC.Columns("A:B")
        .AutoFit
    End With
    
    MsgBox "All tabs renamed and TOC created!", vbInformation, "Success"
End Sub

