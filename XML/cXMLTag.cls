VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXMLTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Value As String

Private Sub Class_Initialize()

End Sub

Public Function ToString() As String
    ToString = Name & "=" & Chr(34) & Value & Chr(34)
End Function
