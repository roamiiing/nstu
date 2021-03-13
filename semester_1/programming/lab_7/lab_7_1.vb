Sub lab7_1()
    Dim word As String, i As Integer
    i = 1
    Do While Range("A" & i).Text <> ""
        word = Range("A" & i).Text
        word = Replace(word, "x", "ks")
        Range("B" & i) = word
        i = i + 1
    Loop
End Sub
