VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXMLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Tags As Collection
Public SubNodes As Collection

Public InnerValue As String
Public Name As String

Private Sub Class_Initialize()
    Set Tags = New Collection
    Set SubNodes = New Collection
End Sub

Public Sub Parse(Text As String)
    ParseXmlNode Text, Me
End Sub

Public Function GetXMLNode(XmlPath As String) As cXMLNode

    On Error GoTo ErrLabel
    Dim Item As Variant

    If Contains(XmlPath, "\") Then
    
        Dim GetSubNode As String
        Dim RemaingPath As String
        
        GetSubNode = Left(XmlPath, Find(XmlPath, "\", 1) - 1)
        RemaingPath = Right(XmlPath, Len(XmlPath) - Len(GetSubNode) - 1)
    
        For Each Item In SubNodes
            If Item.Name = GetSubNode Then
                Set GetXMLNode = Item.GetXMLNode(RemaingPath)
                Exit Function
            End If
        Next
    Else
        For Each Item In SubNodes
            If Item.Name = XmlPath Then
                Set GetXMLNode = Item
                Exit Function
            End If
        Next
    End If
    
    Exit Function
ErrLabel:
    Debug.Print Err.Description
    Debug.Print Err.Source
    Debug.Print Err.HelpContext

End Function
