Attribute VB_Name = "mFile"
Public Function FileExists(Filename As String) As Boolean

    FileExists = (Dir(Filename) <> "")

End Function

Public Function FindDrawing(ByVal Filename As String) As String

    FindDrawing = Left(Filename, Len(Filename) - 7) & ".slddrw"

End Function
