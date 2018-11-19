Attribute VB_Name = "mExcel"
Public Sub Writeline(ByRef Text As String)

    Application.StatusBar = Text
    Debug.Print Text
    DoEvents
    
End Sub
