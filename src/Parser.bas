Option Explicit

' Private attributes
Private pPath As String


' Constructor
Private Sub Class_Initialize() 'Will run when instantiated, but cannot take parameters to set the attributes
    pPath = ""
End Sub

Public Sub Init(ByVal path as String) 'Will be called by the user to set the parameters
    pPath = path
End Sub


' Property Get/Set for Name
Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(path As String)
    pName = path
End Property

