VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJSONNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Attributes As Collection
Public Name As String
Public IsArray As Boolean

Private Sub Class_Initialize()
    Set Attributes = New Collection
    IsArray = False
    Name = "null"
End Sub

Public Sub Parse(Text As String)

    On Error GoTo OnError

    Dim SubItems As Collection
    Dim Item As Variant
    Dim Index As Long
    Dim Value As String
        
    If HasName(Text, Index) Then
        Name = Mid(Text, 1, Index - 1)
        Name = UnQuoteText(TrimToQuote(Name))
        Value = Mid(Text, Index + 1, Len(Text) - Index)
    Else
        Value = Text
    End If
    
    'If Name = "servlet-mapping" Then
    '    Debug.Print "#"
    'End If
    
    Debug.Print Name & ":"
    Value = TrimAwayEmpty(Value)
    
    If StartWith(Value, "{") Then
        Debug.Print "   {"
        Value = TrimToCurlyBrackets(Value)
        mJSON_Parsing.ParseItems Mid(Value, 2, Len(Value) - 1), Attributes
        IsArray = False
        Debug.Print "   }"
        
    ElseIf StartWith(Value, "[") Then
        Debug.Print "   ["
        Value = TrimToSquareBrackets(Value)
        mJSON_Parsing.ParseItems Mid(Value, 2, Len(Value) - 1), Attributes
        IsArray = True
        Debug.Print "   ]"
        
    ElseIf StartWith(Value, Chr(34)) Then
        Debug.Print "   " & Value
        
        If Len(Value) > 2 Then
            Value = UnSlashSpecialCharacters(UnQuoteText(TrimToQuote(Value)))
        Else
            Value = ""
        End If
        
        Attributes.Add Value
        
    Else
        Attributes.Add Value
        Debug.Print "   " & Value
        
    End If
    
    Exit Sub
OnError:
    Debug.Print Err.Description
End Sub

Public Function ToString(Optional Depth As Long = 0)

    Dim EachLine As String
    Dim Ending As String
    Dim I As Long
    EachLine = String(Depth * 2, " ")
    
    ToString = EachLine
    If Name <> "" And Name <> "null" Then ToString = ToString & QuoteText(Name) & ": "
    
    If IsArray Then
        ToString = ToString & "[" & vbNewLine
        Ending = "]"
    Else
        If Attributes.Count > 1 Then
            ToString = ToString & "{" & vbNewLine
            Ending = "}"
        ElseIf TypeOf Attributes(1) Is cJSONNode Then
            ToString = ToString & "{" & vbNewLine
            Ending = "}"
        Else
            ToString = ToString & QuoteText(SlashSpecialCharacters(Attributes(1)))
            Exit Function
        End If
    End If
    
    For I = 1 To Attributes.Count
        If I > 1 And I <= Attributes.Count Then
            ToString = ToString & "," & vbNewLine
        End If
            
        If TypeOf Attributes(I) Is cJSONNode Then
            ToString = ToString & Attributes(I).ToString(Depth + 1)
        Else
            ToString = ToString & EachLine & "  " & QuoteText(Attributes(I))
        End If
    Next
    
    If Ending <> "" Then ToString = ToString & vbNewLine & EachLine & Ending

End Function
