Private Sub OpenNotepad()
    Execute(Base64Decode("Tm90ZXBhZC5leGU="))
End Sub

Private Sub Execute(ByVal command As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run command, 1, False
End Sub

Private Function Base64Decode(ByVal base64 As String) As String
    Dim base64Dec As Object
    Set base64Dec = CreateObject("MSXML2.DOMDocument").createElement("b64")
    base64Dec.DataType = "bin.base64"
    base64Dec.Text = base64
    Base64Decode = base64Dec.nodeTypedValue
End Function
