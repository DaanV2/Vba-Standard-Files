VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Registery As Collection

Private Sub Class_Initialize()
    Set Registery = New Collection
End Sub

Public Sub Clear()
    Set Registery = New Collection
End Sub

Public Function Registrate(Item As String) As Boolean

    If Not Me.Contains(Item) Then
        Registery.Add "", Item
        Registrate = True
    Else
        Registrate = False
    End If

End Function

Public Function Size() As Long

    Size = Registery.Count

End Function

Public Function Contains(Item As String) As Boolean

    On Error GoTo Nope:
    Contains = False
    Dim O As Variant
    O = Registery(Item)
    Contains = True
Nope:

End Function
