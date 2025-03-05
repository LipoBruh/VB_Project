Option Explicit

' Private attributes
Private pPath As String
Private pData As New Collection
Private pFilteredData As New Collection
Private pFileNumber as Integer
Private pWs As Worksheet
Private pSheetNumber as Integer
'Private pWorkSheet As Worksheet



' Constructor
Private Sub Class_Initialize() 'Will run when instantiated, but cannot take parameters to set the attributes
    pPath = ""
End Sub



'Specify constructor attributes 
Public Sub Init(ByVal path as String) 'Will be called by the user to set the parameters
    pPath = path
    pFileNumber = FreeFile
End Sub


' Property Get/Set for Name
Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(path As String)
    pName = path
    pFileNumber = FreeFile
End Property



'Load file content
Public Sub LoadPDB()
    '
    Dim lineText As String
    '
    Open pPath For Input As #pFileNumber
    '
    Do While Not EOF(pFileNumber)
        Line Input #pFileNumber, lineText
        pData.Add lineText  ' Add each line to the collection without a key
    Loop
    '
    Close #pFileNumber
End Sub



Public Sub FindAtoms()
    Set pFilteredData = New Collection
    pSheetNumber = 1
    For Each item In pData
        If InStr(1, item, "ATOM", vbTextCompare) > 0 Then
            pFilteredData.Add item
        End If
    Next item
End Sub

'Manipulate Workbook
Public Sub WriteDataToSheet()
    'helper variables
    Set pWs = ThisWorkbook.Sheets(pSheetNumber)
    Dim i As Integer, j As Integer
    Dim values As Variant
    'regex
    Dim regEx as Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\s+" ' Match multiple spaces/tabs
    regEx.Global = True
    '
    ' Loop through each line and split it
    For i = LBound(pFilteredData) To UBound(pFilteredData) ' LBound = lower bound of the array, UBound = upper bound, useful for a For To loop
        '
        values = regEx.Split(pFilteredData(i))
        '
        ' Loop through each value and place it in the Excel sheet
        For j = LBound(values) To UBound(values)
            pWs.Cells(i + 1, j + 1).Value = values(j) 'x,y coordinates 
        Next j
    Next i
End Sub