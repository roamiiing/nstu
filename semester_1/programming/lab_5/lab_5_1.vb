Sub lab5_1()
Dim i%, w%, p!, q!, s!
Const dw% = 2, w_max% = 100
Range("A1").Value = "w, рад/с"
Range("B1").Value = "АЧХ"
i = 1
For w = 0 To w_max Step dw
    p = 10 - 4 * (w ^ 2)
    q = -0.25 * (w ^ 3) + 12 * w
    s = (1 - 0.05 * w ^ 2) ^ 2 + 0.49 * w ^ 2
    Cells(2 + i, 1).Value = w
    Cells(2 + i, 2).Value = Sqr(p ^ 2 + q ^ 2) / s
    i = i + 1
Next
End Sub
