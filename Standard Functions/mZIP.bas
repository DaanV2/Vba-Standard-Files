Attribute VB_Name = "mZIP"
Attribute VB_Description = "Custom Zip Function"
Public Sub Unzip(ByRef Filepath As Variant, Output As Variant)
    Dim fso As Object
    Dim oApp As Object

    On Error GoTo ErrorL
    'Extract the files into the newly created folder
    Set oApp = CreateObject("Shell.Application")
    Set fso = CreateObject("scripting.filesystemobject")

    oApp.Namespace((Output)).CopyHere oApp.Namespace((Filepath)).Items
    DoEvents
        
    fso.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    
    Exit Sub
ErrorL:
    Debug.Print Err.Description
End Sub
