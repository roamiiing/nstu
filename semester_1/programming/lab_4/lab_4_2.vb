Sub Main()
    Dim a As Single, b As Single, n As Single
    a = 0
    b = 5
    n = 61
    Dim dt As Single, val As Single
    dt = (b - a) / (n - 1)
    For i = 0 To n - 1 Step 1
        val = a + i * dt
        If (val <= 3) Then
            Cells(i + 1, 2) = val ^ 2
        ElseIf (val <= 4) Then
            Cells(i + 1, 2) = 1
        Else: Cells(i + 1, 2) = val ^ 3
        End If
        Cells(i + 1, 1) = val
    Next
End Sub
