Attribute VB_Name = "mProcess"
Function IsProcessRunning(ByRef process As String)
    Dim objList As Object

    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If

End Function

