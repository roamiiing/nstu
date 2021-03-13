Sub Main()
    Dim a As Single, b As Single, pi As Single
    pi = WorksheetFunction.pi
    a = 2
    b = 19.03
    Dim frac As Single, den As Single, num As Single, lnb As Single, result As Single
    frac = a / b
    num = 4.3 * Sin(pi * (frac + 1))
    den = 1 - Cos(pi * (frac - 1))
    lnb = Log(b)
    result = num / den / frac + lnb
    MsgBox ("f(a, b)=" & result)
End Sub
