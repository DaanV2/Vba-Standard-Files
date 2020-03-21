Attribute VB_Name = "mFile"
Public Function FileExists(Filename As String) As Boolean

    FileExists = (Dir(Filename) <> "")

End Function
