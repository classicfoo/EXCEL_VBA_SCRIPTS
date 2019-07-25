Sub transpose()

    ' d = delimiter
    Dim d As String
    d = ","
    
    'r = result
    Dim r As String
    r = ""
    
    Dim rng As Range, cell As Range
    Set rng = Selection
    
    For Each cell In rng
        r = cell.Value + ", " + r
    Next cell
    
    
    r = Left(r, Len(r) - 1)
    r = Left(r, Len(r) - 1)
    
    InputBox "Copy the below text", "Transpose", r

End Sub
