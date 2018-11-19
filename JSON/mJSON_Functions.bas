Attribute VB_Name = "mJSON_Functions"
Public Function HasName(Text As String, ByRef OutFoundAt As Long) As Boolean
    
    Dim InString As Boolean
    Dim Index As Long
    Dim C As String
    Dim TextLength As Long
    
    TextLength = Len(Text)
    
    For Index = 1 To TextLength
        C = Mid(Text, Index, 1)
        
        If C = Chr(34) Then
            If Index > 1 Then
                If Mid(Text, Index - 1, 1) <> "\" Then InString = Not InString
            Else
                InString = Not InString
            End If
        End If
        
        If Not InString Then
            If C = "{" Or C = "[" Then
                Balance = Balance + 1
            ElseIf C = "}" Or C = "]" Then
                Balance = Balance - 1
            End If
            
            If Balance = 0 Then
                If C = ":" Then
                    OutFoundAt = Index
                    HasName = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    OutFoundAt = -1
    HasName = False
    
End Function

Public Function ConvertToJSONAttribute(Name As String, Value As String) As cJSONNode
    Set ConvertToJSONAttribute = New cJSONNode
    
    ConvertToJSONAttribute.Name = Name
    ConvertToJSONAttribute.Attributes.Add Value
End Function
