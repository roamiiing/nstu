Sub lab7_2()
    Dim word As String, i As Integer
    i = 1
    Do While Range("A" & i).Text <> ""
        word = Range("A" & i).Text
        Range("B" & i) = ReplaceSymbols(word)
        i = i + 1
    Loop
End Sub
Function ReplaceSymbols(word As String)
    ReplaceSymbols = Replace(word, "x", "ks")
End Function
