Dim a#, b#, c#, d#

Sub Main()
    Call calcABCD
    Dim result#
    Dim ab#, ad#, bc#, bd#
    ab = fIntegral(a, b, 1)
    ad = fIntegral(a, d, 2)
    bc = fIntegral(b, c, 3)
    bd = fIntegral(b, d, 4)
    result = (ab - ad) / (bc + bd)
    Range("A2").value = a
    Range("B2").value = b
    Range("C2").value = c
    Range("D2").value = d
    Range("A5").value = ab
    Range("B5").value = ad
    Range("C5").value = bc
    Range("D5").value = bd
    Range("E5").value = result
End Sub

Function f(t As Double) As Double
    f = Sqr(Abs(t + 1))
End Function

Function myCos(x As Double) As Double
    Const eps = 0.01
    Dim sum As Double, add As Double, i As Integer
    i = 1
    sum = 0
    add = eps
    Do While Abs(add) >= eps
        add = ((-1) ^ (i - 1)) * (x ^ (2 * i - 2)) / myFactorial(2 * i - 2)
        sum = sum + add
        i = i + 1
    Loop
    myCos = sum
End Function

Function myFactorial(x As Integer) As LongLong
   If ((x = 0) Or (x = 1)) Then
      myFactorial = 1
   Else
      myFactorial = myFactorial(x - 1) * x
   End If
End Function

Sub calcABCD()
    Dim pi#, x1#, x2#, y1#, y2#, x#, y#
    pi = WorksheetFunction.pi
    x1 = 2
    x2 = 2
    y1 = 1
    y2 = 1
    x = pi / 3
    y = pi / 3
    a = (phi(x1, 0) ^ 2) / (2 * phi(x1, y1) + Sqr(phi(x2, y2)))
    b = 18 / (phi(y1, y2) + Sqr(phi(x1, x2)))
    c = 10 / Sqr(myCos(x))
    d = 6 * (1 - myCos(pi + y))
End Sub

Function phi(p As Double, q As Double) As Double
    phi = Exp(p - q)
End Function

Function fIntegral(a As Double, b As Double, column As Integer) As Double
    Const n = 101
    Dim dt As Double, sum As Double, i As Integer
    sum = 0
    dt = (b - a) / (n - 1)
    For i = 0 To n
        Dim x#, value#
        x = a + i * dt
        value = f(x)
        Cells(i + 10, column * 2 - 1).value = x
        Cells(i + 10, column * 2).value = value
        If ((i = 0) Or (i = n)) Then
            sum = sum + 0.5 * value
        Else
            sum = sum + value
        End If
    Next
    fIntegral = sum * dt
End Function
