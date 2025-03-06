Option Explicit

' Private attributes
Private pPath As String
Private pData As New Collection
Private pFilteredData As New Collection
Private pFileNumber As Integer
Private pWs As Worksheet
Private pSheetNumber As Integer
'Private pWorkSheet As Worksheet







' Constructor
Private Sub Class_Initialize() 'Will run when instantiated, but cannot take parameters to set the attributes
    pPath = ""
End Sub







'Specify constructor attributes
Public Sub Init(ByVal path As String) 'Will be called by the user to set the parameters
    pPath = path
    pFileNumber = FreeFile
End Sub






' Property Get/Set for Name
Public Property Get path() As String
    path = pPath
End Property
Public Property Let path(path As String)
    pPath = path
    pFileNumber = FreeFile
End Property





'Load file content
Public Sub LoadPDB()
    ' Check if path exists
    If Len(Dir(pPath)) = 0 Then ' If pPath is empty or path is invalid, Dir(pPath) will return "" such that Len("")==0
        MsgBox "File does not exist."
        Exit Sub
    End If
    
    Dim fileContent As String
    Dim lines() As String
    On Error GoTo ErrorHandler ' Redirect errors to ErrorHandler
    
    ' Read the entire file content as a single string
    Open pPath For Input As #pFileNumber
    fileContent = Input$(LOF(pFileNumber), pFileNumber) ' Read all content in one go
    Close #pFileNumber
    
    ' Split content by LF (\n) line break
    lines = Split(fileContent, vbLf) ' vbLf handles \n (Unix newline)
    
    ' Add each line to pData
    Dim line As Variant
    For Each line In lines
        'Debug.Print "Read line: " & line ' Print the line being read to Immediate Window
        pData.Add line
    Next line
    
    Exit Sub

' Internal error "subroutine"
ErrorHandler:
    HandleError "LoadPDB" ' Call the external error handler
End Sub






'External error subroutine
Private Sub HandleError(ByVal source As String) 'source is a string parameter that holds the name of the procedure where the error occurred, ByVal means the value is passed by copy
    MsgBox "Error in " & source & ": " & Err.Description, vbCritical 'vbCritical is a constant that makes the message box show a red x error icon
    Err.Clear ' Reset the error state
End Sub






Public Sub FindAtoms()

    pWs.Cells(1,1).Value = "PDB Category"
    pWs.Cells(1,2).Value = "Atom ID"
    pWs.Cells(1,3).Value = "Atom Name"
    pWs.Cells(1,4).Value = "Residue"
    pWs.Cells(1,5).Value = "Prot. Chain"
    pWs.Cells(1,6).Value = "Residue Number"
    pWs.Cells(1,7).Value = "X"
    pWs.Cells(1,8).Value = "Y"
    pWs.Cells(1,9).Value = "Z"
    pWs.Cells(1,10).Value = "Occupancy"
    pWs.Cells(1,11).Value = "B-Factor"
    pWs.Cells(1,12).Value = "Element"
    'update attributes
    Set pFilteredData = New Collection 'clears old data
    pSheetNumber = 1
    Sheets(pSheetNumber).Name = "Atoms"
    'Variables
    Dim item As Variant

    For Each item In pData
        If InStr(1, item, "ATOM", vbTextCompare) > 0 Then
            
            If InStr(1, item, "REVDAT", vbTextCompare) = 0 And InStr(1, item, "CAVEAT", vbTextCompare) = 0 And InStr(1, item, "REMARK", vbTextCompare) = 0 Then
            
                pFilteredData.Add item
            End If
        End If
    Next item
End Sub




'Regex split from
'https://github.com/ReneNyffenegger/about-VBA/blob/master/regular-expressions/split.bas
Private Function regexpSplit(text As String, pattern As String) As String()
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim text_0 As String
    
    
    re.pattern = pattern
    re.Global = True
    re.MultiLine = True
    
    text_0 = re.Replace(text, vbNullChar)
    
    regexpSplit = Strings.Split(text_0, vbNullChar)
End Function




'Manipulate Workbook
Public Sub WriteDataToSheet()
    'helper variables
    Set pWs = ThisWorkbook.Sheets(pSheetNumber)
    Dim values As Variant
    Dim item As Variant
    
    ' Loop through each line and split it
    Dim i As Integer, j As Integer
    i = 2
    
    For Each item In pFilteredData
    
        Debug.Print "Item: " & item
        '
        values = regexpSplit(CStr(item), "\s+")
        '
        ' Loop through each value and place it in the Excel sheet
        For j = LBound(values) To UBound(values)
            pWs.Cells(i + 1, j + 1).Value = values(j) 'x,y coordinates
        Next j
        
        i = i + 1
    Next item
End Sub
