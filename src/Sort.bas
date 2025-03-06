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


Sub SortAsc(ByVal sheetName As String, ByVal headerName As String)
    
    Dim ws As Worksheet
    Dim colIndex As Integer
    '
    Set ws = ThisWorkbook.Sheets(sheetName)
    colIndex = GetColumnIndex(ws, headerName)
    '
    ' Check if the column index was found
    If colIndex = -1 Then
        MsgBox "Header '" & headerName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If
    '
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    '
    Dim rng As range 'SortFields.Add  key must take a Range object
    Set rng = ws.range(ws.Cells(1, colIndex), ws.Cells(lastRow, colIndex))
    '
    With ws.Sort
    .SortFields.Clear 'clearing previous sorts
    .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    .SetRange rng
    .Header = xlYes
    .Apply
    End With


End Sub


Sub SortBFactors()
    
    SortAsc "Atoms", "B-Factor"
    
End Sub


