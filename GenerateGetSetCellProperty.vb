Public Sub GetOrSetSelectedCellsNamedRanges()
'Date: 8/7/18
'Author: Thomas Divine Smith
'Project: Excel VBA
'Purpose: Creates a Get/Set Property to access selected cells in Visual Studio/Vb.NET
'Instructions: Select your cells & run this macro.

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
        
        'Create property that returns value of cell
        
        'Name propety
        Dim prop_title As String
        prop_title = "Public Property " & cell_name
        
        'Remove sheet name referance in property title
        If InStr(prop_title, remove_string) Then
            prop_title = Replace(prop_title, remove_string, "")
        End If
        
        'Return value
        Dim prop_get As String
        prop_get = "Return xl" & ActiveSheet.Name & ".Range(" & named_range & ").value"
        
        Dim prop_set As String
        prop_set = "xl" & ActiveSheet.Name & ".Range(" & named_range & ").value"
        
        Debug.Print (prop_title)
        Debug.Print ("Get")
        Debug.Print (prop_get)
        Debug.Print ("End Get")
        Debug.Print ("Set(ByVal Value)")
        Debug.Print (prop_set & " = Value")
        Debug.Print ("End Set")
        Debug.Print ("End Property")
        
'APPEARS IN IMMEDIATES WINDOW
'RESULT EXAMPLE
'        Public Shared Property tBlk_LineItemNumber
'            Get
'                Return xlTitleBlock.Range("tBlk_LineItemNumber").Value
'            End Get
'            Set(ByVal Value)
'                xlTitleBlock.Range("tBlk_LineItemNumber").Value = Value
'            End Set
'        End Property
    Next

End Sub
