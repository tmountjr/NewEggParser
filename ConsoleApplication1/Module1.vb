Module Module1

    Sub Main()
        Dim nep As New NeweggParser.NeweggParser
        'nep.Load("c:\users\inet\desktop\foo.json")
        nep.LoadNeweggNumber("N82E16813131976")
        Stop
        Dim keys As String() = nep.GetKeys
        Dim foo As Object = nep.GetProperty("imageGallery")
        Dim foo2 As String() = nep.GetKeys(foo)
        Dim price As String = nep.GetProperty("FinalPrice")
        Stop
    End Sub

End Module
