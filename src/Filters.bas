
'will use the title / header to find the id
Function GetColumnIndex(ByVal ws As Worksheet, ByVal headerName As String) As Integer 'Byval passes a copy
    Dim range As range
    Set range = ws.Rows(1).Find(headerName, LookAt:=xlWhole) 'find header in first row : Returns Nothing object or Range object if found

    If Not range Is Nothing Then
        GetColumnIndex = range.Column ' Return column number
    Else
        GetColumnIndex = -1 ' Return -1 if not found
    End If
End Function





Sub FilterByDropdown(ByVal sheetName As String, ByVal headerName As String, ByVal criteria As Variant)

    Dim ws As Worksheet
    Dim colIndex As Integer
    '
    Set ws = ThisWorkbook.Sheets(sheetName)
    colIndex = GetColumnIndex(ws, headerName)
    '
    ' Check if the column index was found
    Debug.Print TypeName(colIndex)
    If colIndex = -1 Then
        MsgBox "Header '" & headerName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    ' Apply AutoFilter based on the column index and criteria -> Autofilters are cumulative
    If IsArray(criteria) Then
        ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria, Operator:=xlFilterValues
        '
    Else
        ' Single criteria
        ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria
    End If
    
    'ws.AutoFilterMode = False     can clear an autofilter
End Sub

Sub FilterByDropdownRange(ByVal sheetName As String, ByVal headerName As String, ByVal criteria As Variant)

    Dim ws As Worksheet
    Dim colIndex As Integer
    '
    Set ws = ThisWorkbook.Sheets(sheetName)
    colIndex = GetColumnIndex(ws, headerName)
    '
    ' Check if the column index was found
    Debug.Print TypeName(colIndex)
    If colIndex = -1 Then
        MsgBox "Header '" & headerName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    ' Apply AutoFilter based on the column index and criteria -> Autofilters are cumulative
    If IsArray(criteria) Then
        
        If UBound(criteria) - LBound(criteria) + 1 <= 1 Then
            ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria, Operator:=xlFilterValues
        Else
            Debug.Print "TRACE: " & criteria(1)
            ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria(0), Operator:=xlAnd, Criteria2:=criteria(1)
        End If
        '
    Else
        ' Single criteria
        ws.range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria
    End If
    
    'ws.AutoFilterMode = False     can clear an autofilter
End Sub

Sub ResetAutoFilter(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Check if AutoFilter is enabled
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False ' Turns off AutoFilter, removing all filters
    End If
End Sub


Sub ResetAtoms()
    ResetAutoFilter "Atoms"
End Sub





Sub FindCarbons()

    FindElement "C"
    
End Sub

Sub FindNitrogens()
    Dim atom As Variant
    atom = "N"
    FindElement atom
    
End Sub

Sub FindMetals()
    '
    Dim metallicElements As Variant
    metallicElements = Array("Fe", "Zn", "Cu", "Mg", "Mn", "Co", "Ni", "Ca", "Mo", "W", "V")
    '
    FindElement metallicElements
    
End Sub


Sub BFactorRange()
    '
    Dim range As Variant
    range = Array(">5.0", "<10.0")
    '
    FindBFactor range
    
End Sub

Sub BLowerBound()
    '
    Dim range As Variant
    range = Array(">5.0")
    '
    FindBFactor range
    
End Sub
Sub FindBFactor(element As Variant)

    FilterByDropdownRange "Atoms", "B-Factor", element
    
End Sub


Sub FindElement(element As Variant)

    FilterByDropdown "Atoms", "Element", element
    
End Sub
