Attribute VB_Name = "mCollection"

'Check if a Collection contains an item that is of type: string, int, double, float
Public Function CollectionContains(ByRef Key As String, ByRef Col As Collection) As Boolean

    On Error GoTo Nope
    Dim O As Variant
    O = Col(Key)
    CollectionContains = True
Nope:

End Function

'Check if a Collection contains an item that is defined as a class / object
Public Function CollectionContains2(ByRef Key As String, ByRef Col As Collection) As Boolean

    On Error GoTo Nope
    Dim O As Variant
    Set O = Col(Key)
    CollectionContains2 = True
    Exit Function
Nope:

End Function

