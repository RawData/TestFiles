Copy code
Private Sub ObfuscatedCode()
    Call Execute(Base64Decode("JG1haW4gPSAiSEki"))
End Sub

Private Sub Execute(ByVal code As String)
    Dim bytes() As Byte
    bytes = DecodeBase64Bytes(code)
    Dim str As String
    str = StrConv(bytes, vbUnicode)
    ExecuteVBA str
End Sub

Private Sub ExecuteVBA(ByVal code As String)
    Dim objScript As Object
    Set objScript = CreateObject("ScriptControl")
    objScript.Language = "VBScript"
    objScript.AddCode code
End Sub

Private Function DecodeBase64Bytes(ByVal base64 As String) As Byte()
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = base64
    DecodeBase64Bytes = objNode.nodeTypedValue
End Function

Private Function Base64Decode(ByVal base64 As String) As String
    Dim objXML As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Base64Decode = objXML.createElement("tmp").nodeTypedValue
    With CreateObject("Microsoft.XMLDOM")
        .Async = False: .ValidateOnParse = False
        .LoadXML "<root><t>" & base64 & "</t></root>"
        Base64Decode = .getElementsByTagName("t")(0).nodeTypedValue
    End With
End Function
