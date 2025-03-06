


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
