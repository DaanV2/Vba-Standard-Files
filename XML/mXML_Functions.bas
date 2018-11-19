Attribute VB_Name = "mXML_Functions"
Attribute VB_Description = "Custom XML functions"
Public Function ExtractXMLNode(ByRef NodeName As String, ByRef Text As String, ByRef StartPosition As Long) As String
Attribute ExtractXMLNode.VB_Description = "Extract one node from a XML text"
        Dim StartVar As Long
        Dim EndVar As Long
        Dim Balance As Long
        Dim StartText As String
    
        StartVar = Find(Text, "<" & NodeName, StartPosition)
        
Restart:
        If StartVar = 0 Then
            ExtractXMLNode = ""
        Else
            StartText = Mid(Text, StartVar, Len(NodeName) + 2)
            
            If Not (StartText = "<" & NodeName & " " Or StartText = "<" & NodeName & "/" Or StartText = "<" & NodeName & ">") Then
                StartVar = Find(Text, "<" & NodeName, StartVar + 1)
                GoTo Restart
            Else
                'Debug.Print StartText
            End If

            EndVar = FindEndNode(NodeName, Text, StartVar)
            ExtractXMLNode = Mid(Text, StartVar, EndVar - StartVar)
        End If

End Function

Public Function ExtractFirstXMLNode(ByRef Text As String, Optional ByRef StartPosition As Long = 1) As String
        Dim StartVar As Long
        Dim EndVar As Long
        Dim I As Long
        Dim T As String
        Dim NodeName As String
           
        NodeName = ExtractNodeName(Text, StartPosition)
        
        If NodeName = "" Then Exit Function
        If Contains(NodeName, "/") Then
            Debug.Print NodeName
        End If
        
        ExtractFirstXMLNode = ExtractXMLNode(NodeName, Text, StartPosition)
    
End Function

Public Sub ExtractXMLCollection(ByRef ItemTag As String, ByRef CollectionText As String, ByRef StartPosition As Long, ByRef Output As Collection)
Attribute ExtractXMLCollection.VB_Description = "Extract a collection of nodes from a XML text"

    Dim Item As String
    Dim I As Long
    Dim Max As Long
    
    If CollectionText = "" Then Exit Sub
        
    I = 1
    Max = Len(CollectionText)
    'Get First Sub Item
    Item = ExtractXMLNode(ItemTag, CollectionText, I)
    If Item = "" Then Exit Sub
    
    While I < Max
        
        'Debug.Print "   " & I & "/" & Max
        Output.Add Item
    
        I = Find(CollectionText, Item, I) + 1
        Item = ExtractXMLNode(ItemTag, CollectionText, I)
        
        If Item = "" Then I = Max
    Wend

End Sub

Public Function ExtractXMLTag(ByRef Text As String, ByRef XMLTag As String, Optional StartPosition As Long = 1) As String
Attribute ExtractXMLTag.VB_Description = "Extract a tag within a node"
        Dim StartVar As Long
        Dim EndVar As Long
    
        StartVar = Find(Text, XMLTag + "=" + Chr(34), StartPosition) + Len(XMLTag) + 2
        EndVar = Find(Text, Chr(34), StartVar)
        
        ExtractXMLTag = Mid(Text, StartVar, EndVar - StartVar)
End Function

Public Function FindEndNode(ByRef NodeName As String, ByRef Text As String, Optional StartOfNode As Long = 1) As Long
Attribute FindEndNode.VB_Description = "Finds the end of a XML node "

    Dim EndVar As Long
    Dim Instring As Boolean
    Dim OutsideNode As Boolean

    For EndVar = StartOfNode + Len(NodeName) + 1 To Len(Text)
        Dim Temp As String
        Dim Ending As String
        
        Temp = Mid(Text, EndVar, 12)
               
        If Instring Then
            If Mid(Text, EndVar, 1) = Chr(34) Then Instring = False
        Else
            If OutsideNode Then
                If Mid(Text, EndVar, Len(NodeName) + 3) = "</" & NodeName & ">" Then
                    EndVar = EndVar + Len(NodeName) + 3
                    GoTo Done
                End If
            Else
                If Mid(Text, EndVar, 3) = " />" Then
                    EndVar = EndVar + 3
                    GoTo Done
                ElseIf Mid(Text, EndVar, 2) = "?>" Then
                    EndVar = EndVar + 2
                    GoTo Done
                End If
                
                If Mid(Text, EndVar, 1) = ">" Then OutsideNode = True
            End If
            
            If Mid(Text, EndVar, 1) = Chr(34) Then Instring = True
        End If
    Next
Done:
    
    FindEndNode = EndVar

End Function

Public Function ExtractNodeName(ByRef Text As String, Optional StartPosition As Long = 1)
    Dim StartVar As Long
    Dim EndVar As Long

    For StartVar = StartPosition To Len(Text)
        If Mid(Text, StartVar, 1) = "<" And Mid(Text, StartVar + 1, 1) <> "/" Then
            For I = StartVar To Len(Text)
                T = Mid(Text, I, 1)
                
                If T = " " Or T = ">" Then
                    ExtractNodeName = Replace(Mid(Text, StartVar + 1, I - StartVar - 1), " /", "")
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Public Function ExtractNodeContainer(ByRef Text As String, Optional StartPosition As Long = 1) As String
    Dim StartVar As Long
    Dim EndVar As Long
    Dim InText As Boolean
    Dim T As String
    
    For StartVar = StartPosition To Len(Text)
        If Mid(Text, StartVar, 1) = "<" Then
            For EndVar = StartVar To Len(Text)
                T = Mid(Text, EndVar, 1)
            
                If T = Chr(34) Then InText = Not InText
                
                If Not InText And (T = ">") Then
                    ExtractNodeContainer = Mid(Text, StartVar, EndVar - StartVar + 1)
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Public Function ExtractInnerValue(ByRef NodeText As String) As String
    Dim StartVar As Long
    Dim EndVar As Long
    Dim Container As String
    Dim NodeName As String
    
    Container = ExtractNodeContainer(NodeText, 1)
    
    If Container = NodeText Then Exit Function
    
    NodeName = ExtractNodeName(Container, 1)
    StartVar = Find(NodeText, Container) + Len(Container)
    EndVar = FindEndNode(NodeName, NodeText, Find(NodeText, Container) + 1) - (Len(NodeName) + 3)
    ExtractInnerValue = Mid(NodeText, StartVar, EndVar - StartVar)
    
End Function
