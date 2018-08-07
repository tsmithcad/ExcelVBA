Public Sub PrintSelectedCellsRangeName()
'Date: 8/7/18
'Author: Thomas Divine Smith
'Project: Excel VBA
'Purpose: Print selected cell range name
'Requirements: Selected cells must have individual range name

On Error Resume Next

    Dim c As Range
    
    'Get named ranges of each selected cell
    For Each c In Selection
        
        'Remove sheet name referance
        Dim remove_string As String
        remove_string = ActiveSheet.Name & "!"
        
        'Get cell named range
        Dim cell_name As String
        cell_name = c.Name.Name
        
        'Remove sheet name referance in cell name
        If InStr(cell_name, remove_string) Then
            cell_name = Replace(cell_name, remove_string, "")
        End If
        
        'Get named range with quotations
        Dim named_range As String
        named_range = (Chr(34) & cell_name & Chr(34))
        
        'Remove sheet name referance in named range
        If InStr(named_range, remove_string) Then
            named_range = Replace(named_range, remove_string, "")
        End If
        
        Debug.Print (named_range)
        
    Next

End Sub
