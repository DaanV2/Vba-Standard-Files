Attribute VB_Name = "mXML_Parse_Functions"
Public Sub ParseXmlFile(Filepath As String, ByRef Output As Collection)

    On Error GoTo ErrLabel
    ParseXmlText ReadAllText(Filepath), Output
    
    Exit Sub
ErrLabel:
    Debug.Print Err.Description
    Debug.Print Err.Source
    Debug.Print Err.HelpContext
    
End Sub

Public Sub ParseXmlText(Text As String, ByRef Output As Collection)

    Debug.Print "Parse Xml Text"

    Dim CurrentNode As String
    
    On Error GoTo ErrLabel
    CurrentNode = ExtractFirstXMLNode(Text)
    
    While CurrentNode <> ""
        Dim NewNode As New cXMLNode
        NewNode.Parse CurrentNode
        Output.Add NewNode
        
        Set NewNode = Nothing
        CurrentNode = ExtractFirstXMLNode(Text, Find(Text, CurrentNode) + Len(CurrentNode))
    Wend
    
    Exit Sub
ErrLabel:
    Debug.Print Err.Description
    Debug.Print Err.Source
    Debug.Print Err.HelpContext

End Sub

Public Sub ParseXmlNode(Text As String, ByRef Node As cXMLNode)

    Debug.Print "Parse Xml Node"

    Dim Container As String
    Dim EndVar As Long
    
    On Error GoTo ErrLabel
    Container = ExtractNodeContainer(Text)
    Node.Name = ExtractNodeName(Container): 'Debug.Print "Node: " & Node.Name
        
    ParseXmlNodeTags Container, Node
    Node.InnerValue = ExtractInnerValue(Text)
    
    If Node.InnerValue <> "" Then
        If Contains(Node.InnerValue, "<") And Contains(Node.InnerValue, "/") And Contains(Node.InnerValue, ">") Then
            ParseXmlText Node.InnerValue, Node.SubNodes
        End If
    End If
    
    Exit Sub
ErrLabel:
    Debug.Print Err.Description
    Debug.Print Err.Source
    Debug.Print Err.HelpContext
    
End Sub

Public Sub ParseXmlNodeTags(Container As String, ByRef Node As cXMLNode)
    
    Debug.Print "Parse Xml Node Tags"
    
    If Not Contains(Container, "=") Then Exit Sub

    Dim T As String
    Dim StartTag As Long
    Dim EndTag As Long
    Dim I As Long
    Dim InText As Boolean

    On Error GoTo ErrLabel
    For I = 1 To Len(Container)
        T = Mid(Container, I, 1)
        
        If T = Chr(34) Then InText = Not InText
        
        If (Not InText) And T = "=" Then
            Dim NewTag As New cXMLTag
        
            For StartTag = I To 1 Step -1
                If Mid(Container, StartTag, 1) = " " Then Exit For
            Next
            For EndTag = I + 2 To Len(Container)
                If Mid(Container, EndTag, 1) = Chr(34) Then Exit For
            Next
            
            NewTag.Name = Mid(Container, StartTag + 1, I - StartTag - 1)
            NewTag.Value = Mid(Container, I + 2, EndTag - (I + 2))
            
            'Debug.Print "Tag:   " & NewTag.ToString()
            
            Node.Tags.Add NewTag
            
            Set NewTag = Nothing
        End If
    Next
    
    Exit Sub
ErrLabel:
    Debug.Print Err.Description
    Debug.Print Err.Source
    Debug.Print Err.HelpContext

End Sub




