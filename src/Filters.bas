
'will use the title / header to find the id
Function GetColumnIndex(ByVal ws As Worksheet, ByVal headerName As String) As Integer 'Byval passes a copy
    Dim range As Range
    Set range = ws.Rows(1).Find(headerName, LookAt:=xlWhole) 'find header in first row : Returns Nothing object or Range object if found

    If Not range Is Nothing Then
        GetColumnIndex = range.Column ' Return column number
    Else
        GetColumnIndex = -1 ' Return -1 if not found
    End If
End Function



'
Sub FilterByDropdown()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterCriteria As String

    Set ws = ThisWorkbook.Sheets("Atoms")
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    filterCriteria = ws.Range("B1").Value ' Read selected filter from dropdown
    
    ' Apply AutoFilter
    ws.Range("A1:H" & lastRow).AutoFilter Field:=4, Criteria1:=filterCriteria
End Sub


Sub FilterByDropdown(sheetName as string, headerName As String, criteria As String)
    Dim ws As Worksheet
    Dim colIndex As Integer
    '
    Set ws = ThisWorkbook.Sheets(sheetName)
    colIndex = GetColumnIndex(headerName)
    '
    ' Check if the column index was found
    If colIndex = -1 Then
        MsgBox "Header '" & headerName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Apply AutoFilter based on the column index and criteria -> Autofilters are cumulative 
    ws.Range("A1").CurrentRegion.AutoFilter Field:=colIndex, Criteria1:=criteria  'autofilter hides rows on the current region (connected squares to A1) based on the criteria specified on the column
    'ws.AutoFilterMode = False     can clear an autofilter
End Sub
