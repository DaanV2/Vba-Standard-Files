Attribute VB_Name = "mJSON_Functions_string"
Public Function SlashSpecialCharacters(Text As String) As String
    SlashSpecialCharacters = Replace(Text, Chr(34), "\" & Chr(34))
    SlashSpecialCharacters = Replace(SlashSpecialCharacters, "\", "\\")

End Function

Public Function UnSlashSpecialCharacters(Text As String) As String
    UnSlashSpecialCharacters = Replace(Text, "\" & Chr(34), Chr(34))
    UnSlashSpecialCharacters = Replace(UnSlashSpecialCharacters, "\\", "\")

End Function

Public Function QuoteText(Text As String) As String
    QuoteText = Chr(34) & Text & Chr(34)
    
End Function

Public Function UnQuoteText(Text As String) As String
    UnQuoteText = Mid(Text, 2, Len(Text) - 2)
    
End Function

Public Function TrimToQuote(Text As String) As String
    Text = TrimStartTo(Text, Chr(34))
    TrimToQuote = TrimEndTo(Text, Chr(34))

End Function

Public Function TrimToCurlyBrackets(Text As String) As String
    Text = TrimStartTo(Text, "{")
    TrimToCurlyBrackets = TrimEndTo(Text, "}")

End Function

Public Function TrimToSquareBrackets(Text As String) As String

    Text = TrimStartTo(Text, "[")
    TrimToSquareBrackets = TrimEndTo(Text, "]")

End Function

Public Function TrimStartTo(Text As String, StartText As String) As String
    TrimStartTo = Text
    
    If StartWith(Text, StartText) Then Exit Function
    
    If Contains(Text, StartText) Then
        Dim I As Long
        I = Find(Text, StartText, 1)
        TrimStartTo = Mid(Text, I, Len(Text) - I + 1)
    End If

End Function

Public Function TrimEndTo(Text As String, EndText As String) As String
    TrimEndTo = Text
    
    If EndWith(Text, EndText) Then Exit Function
    
    If Contains(Text, EndText) Then

        While Not EndWith(Text, EndText)
            Text = Left(Text, Len(Text) - 1)
        Wend
    End If

End Function

Public Function TrimAwayEmpty(Text As String)

    While StartWith(Text, " ") Or StartWith(Text, vbTab) Or StartWith(Text, vbCr) Or StartWith(Text, vbLf) Or StartWith(Text, vbNewLine)
        Text = Right(Text, Len(Text) - 1)
    Wend
    
    While EndWith(Text, " ") Or EndWith(Text, vbTab) Or EndWith(Text, vbCr) Or EndWith(Text, vbLf) Or EndWith(Text, vbNewLine)
        Text = Left(Text, Len(Text) - 1)
    Wend

    TrimAwayEmpty = Text

End Function

