VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuitableXtraData2720"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public xtraRqms As Collection


Private Sub Class_Initialize()
    Set xtraRqms = New Collection
End Sub

Private Sub Class_Terminate()
    Set xtraRqms = Nothing
End Sub
