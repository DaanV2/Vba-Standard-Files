Attribute VB_Name = "mJSON_Parsing"
Public Sub GetItems(Text As String, ByRef Output As Collection)

    Dim StartVar As Long
    Dim Index As Long
    Dim TextLength As Long
    Dim Balance As Long
    Dim InString As Boolean
    Dim C As String
    
    StartVar = 1
    TextLength = Len(Text)
    
    For Index = 1 To TextLength
        C = Mid(Text, Index, 1)
        
        If C = Chr(34) Then
            If Index > 1 Then
                If Mid(Text, Index - 1, 1) <> "\" Then InString = Not InString
            Else
                InString = Not InString
            End If
        Else
            If Not InString Then
                If C = "{" Or C = "[" Then
                    Balance = Balance + 1
                ElseIf C = "}" Or C = "]" Then
                    Balance = Balance - 1
                End If
                
                If Balance = 0 Then
                    If C = "," Then
                        Dim Item As String
                        Item = Mid(Text, StartVar, Index - StartVar)
                        StartVar = Index + 1
                        
                        Output.Add Item
                    End If
                End If
            End If
        End If
    Next
    
    Item = Mid(Text, StartVar, Len(Text) - StartVar)
    Output.Add Item
    
End Sub

Public Sub ParseItems(Text As String, ByRef Output As Collection)
    Dim SubItems As Collection
    Dim Index As Long
    Dim Item As String
    Dim J As Variant
    
    On Error GoTo OnError
    Text = TrimAwayEmpty(Text)

    Set SubItems = New Collection
    GetItems Text, SubItems
    
    For Each J In SubItems
        Item = TrimAwayEmpty(CStr(J))
        
        If StartWith(Item, "{") Or StartWith(Item, "[") Or mJSON_Functions.HasName(Text, Index) Then
            Dim NewNode As New cJSONNode
            NewNode.Parse CStr(Item)
            Output.Add NewNode
            
            Set NewNode = Nothing
        Else
            If StartWith(Item, Chr(34)) Then
                If Len(Item) > 2 Then
                    Item = UnSlashSpecialCharacters(UnQuoteText(TrimToQuote(Item)))
                Else
                    Item = ""
                End If
                
                Output.Add CStr(Item)
            Else
                Output.Add CStr(Item)
            End If
        End If
    Next
    
    Exit Sub
OnError:
    Debug.Print Err.Description
End Sub
