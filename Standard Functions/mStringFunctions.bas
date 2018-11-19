Attribute VB_Name = "mStringFunctions"
Attribute VB_Description = "The module contains custom string functions based on .NET string functions"
Public Function Contains(ByRef InText As String, ByRef FindText As String) As Boolean
Attribute Contains.VB_Description = "Checks whenever the specified text can be found in the first specified string"
    Contains = (InStr(1, InText, FindText, vbTextCompare) > 0)
End Function

Public Function Find(ByRef InText As String, ByRef FindText As String, Optional StartPos As Long = 1) As Long
Attribute Find.VB_Description = "find the staring position of a text within another text"
    Find = InStr(StartPos, InText, FindText, vbTextCompare)
End Function

Public Function StartWith(ByRef Text As String, ByRef StartWithText As String) As Boolean
Attribute StartWith.VB_Description = "Check if a string starts with another string"
    StartWith = Left(Text, Len(StartWithText)) = StartWithText
End Function

Public Function EndWith(ByRef Text As String, ByRef EndWithText As String) As Boolean
Attribute EndWith.VB_Description = "Checks if a string end with another string"
    EndWith = Right(Text, Len(EndWithText)) = EndWithText
End Function

Public Function Remove(ByRef Text As String, ByRef Length As Long)
Attribute Remove.VB_Description = "Removes a certain length of characters from the start point of the string"
    Remove = Right(Text, Len(Text) - Length)
End Function

Public Function TrimText(Text As String) As String
Attribute TrimText.VB_Description = "trim texts from empty starting characters"

    If StartWith(Text, vbCrLf) Then Text = Remove(Text, 2)
    If StartWith(Text, vbCr) Then Text = Remove(Text, 1)
    If StartWith(Text, vbLf) Then Text = Remove(Text, 1)
    If StartWith(Text, vbNewLine) Then Text = Remove(Text, 1)

    TrimText = Text

End Function

