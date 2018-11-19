Attribute VB_Name = "mXML"
Attribute VB_Description = "Custom XML functions"
Public Function ExtractXMLNode(Tag As String, Text As String, StartPosition As Long) As String
Attribute ExtractXMLNode.VB_Description = "Extract one node from a XML text"
        Dim StartVar As Long
        Dim EndVar As Long
        Dim Balance As Long
        Dim StartText As String
    
        StartVar = Find(Text, "<" & Tag, StartPosition)
        
Restart:
        If StartVar = 0 Then
            ExtractXMLNode = ""
        Else
            StartText = Mid(Text, StartVar, Len(Tag) + 2)
            
            If Not (StartText = "<" & Tag & " " Or StartText = "<" & Tag & "/" Or StartText = "<" & Tag & ">") Then
                StartVar = Find(Text, "<" & Tag, StartVar + 1)
                GoTo Restart
            Else
                'Debug.Print StartText
            End If

            EndVar = FindEndTag(Tag, Text, StartVar)
            ExtractXMLNode = Mid(Text, StartVar, EndVar - StartVar)
        End If

End Function

Public Sub ExtractXMLCollection(ItemTag As String, CollectionText As String, StartPosition As Long, ByRef Output As Collection)
Attribute ExtractXMLCollection.VB_Description = "Extract a collection of nodes from a XML text"

    Dim Item As String
    Dim I As Long
    
    If CollectionText = "" Then Exit Sub
        
    I = 1
    'Get First Sub Item
    Item = ExtractXMLNode(ItemTag, CollectionText, I)
    If Item = "" Then Exit Sub
    
    While I < Len(CollectionText)
    
        Output.Add Item
    
        I = Find(CollectionText, Item, 1) + Len(Item)
        Item = ExtractXMLNode(ItemTag, CollectionText, I)
        
        If Item = "" Then I = Len(CollectionText)
    Wend

End Sub

Public Function ExtractXMLTag(Text As String, XMLTag As String, Optional StartPosition As Long = 1) As String
Attribute ExtractXMLTag.VB_Description = "Extract a tag within a node"
        Dim StartVar As Long
        Dim EndVar As Long
    
        StartVar = Find(Text, XMLTag + "=" + Chr(34), StartPosition) + Len(XMLTag) + 2
        EndVar = Find(Text, Chr(34), StartVar)
        
        ExtractXMLTag = Mid(Text, StartVar, EndVar - StartVar)
End Function

Public Function FindEndTag(Tag As String, Text As String, StartOfTag As Long) As Long
Attribute FindEndTag.VB_Description = "Finds the end of a XML node "

    Dim EndVar As Long
    Dim Instring As Boolean
    Dim OutsideNode As Boolean

    For EndVar = StartOfTag + Len(Tag) + 1 To Len(Text)
        Dim Temp As String
        Dim Ending As String
        
        Temp = Mid(Text, EndVar, 12)
               
        If Instring Then
            If Mid(Text, EndVar, 1) = Chr(34) Then Instring = False
        Else
            If OutsideNode Then
                If Mid(Text, EndVar, Len(Tag) + 3) = "</" & Tag & ">" Then
                    EndVar = EndVar + Len(Tag) + 3
                    GoTo Done
                End If
            Else
                If Mid(Text, EndVar, 3) = " />" Then
                    EndVar = EndVar + 3
                    GoTo Done
                End If
                
                If Mid(Text, EndVar, 1) = ">" Then OutsideNode = True
            End If
            
            If Mid(Text, EndVar, 1) = Chr(34) Then Instring = True
        End If
    Next
Done:
    
    FindEndTag = EndVar

End Function
