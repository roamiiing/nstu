Function Factorial(n As Long) As Currency
    If n <= 1 Then
        Factorial = 1
        Exit Function
    End If
    Factorial = Factorial(n - 1) * n
End Function

Function MyExp(x As Double, e As Double) As Double
    Dim s As Double, val As Double, i As Long
    s = 0
    i = 1
    Do
        val = (x ^ (i - 1)) / Factorial(i - 1)
        s = s + val
        i = i + 1
    Loop While (val > e)
    MyExp = s
End Function

Sub Main()
    Dim ePower As Double, result As Double, x As Double, epsylon As Double
    epsylon = 0.01
    x = 4
    ePower = MyExp(x, epsylon)
    result = Abs(ePower / (ePower + 1 / ePower))
    MsgBox ("f(x)=" & result)
End Sub
