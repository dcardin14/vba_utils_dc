Attribute VB_Name = "Module8"
Sub shorthandToAliquots()
    Dim selectedRange As Range
    Dim cell As Range
    Dim dict As Object
    Dim textArray() As String
    Dim result As String
    Dim word As Variant
    
    ' Check if there is any selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range containing text."
        Exit Sub
    End If
    
    ' Initialize selected range
    Set selectedRange = Selection
    
    ' Define the dictionary mapping shorthand to verbose
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Define mappings for quarters and halves (adjust or expand as needed)
    dict.Add "NE", "NE¼"
    dict.Add "SE", "SE¼"
    dict.Add "NW", "NW¼"
    dict.Add "SW", "SW¼"
    
    dict.Add "N2", "N½"
    dict.Add "S2", "S½"
    dict.Add "E2", "E½"
    dict.Add "W2", "W½"
    
    ' Iterate through each cell in the selected range
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) Then
            ' Split cell value using SplitSpecial function
            textArray = SplitSpecial(cell.Value)
            
            ' Initialize result string
            result = ""
            
            ' Iterate through each word in the array
            For Each word In textArray
                If dict.Exists(word) Then
                    result = result & dict(word) & " "
                Else
                    result = result & word & " "
                End If
            Next word
            
            ' Replace cell value with the converted text
            cell.Value = Trim(result)
        End If
    Next cell
    
    MsgBox "Conversion complete."
End Sub

Function SplitSpecial(text As String) As Variant
    Dim firstDelimiterPos As Long
    Dim firstSubstring As String
    Dim textArray() As String
    
    ' Find the position of the colon and space ": "
    firstDelimiterPos = InStr(1, text, ": ")
    
    ' If the delimiter is found
    If firstDelimiterPos > 0 Then
        ' Extract the first substring starting after the colon and spaces
        firstSubstring = Mid(text, firstDelimiterPos + 2)
        
        ' Split the first substring based on the delimiter ", "
        textArray = Split(firstSubstring, ", ")
    Else
        ' No delimiter found, the entire text is treated as a single substring
        textArray = Array(text)
    End If
    
    ' Return the text array
    SplitSpecial = textArray
End Function

