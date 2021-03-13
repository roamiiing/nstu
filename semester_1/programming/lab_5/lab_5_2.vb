Sub lab5_2()
  Dim matr%(1 To 5, 1 To 5), var As Variant
  Dim i%, j%, max%
  var = Array("a", "b", "c", "d", "e")
  For i = 1 To 5
  For j = 1 To 5
  matr(i, j) = Range(var(i - 1) & j + 1).Value
  Next: Next
  ' Для начала, присвоим max сумму первой строки матрицы
  max = 0
  For i = 1 To 5
      max = max + matr(i, 1)
  Next
  ' Затем будем находить суммы остальных строк, и если
  ' какая-то из них окажется больше max, присвоим max эту сумму
  For i = 2 To 5
      Dim sum%
      sum = 0
      For j = 1 To 5
          sum = sum + matr(j, i)
      Next
      max = WorksheetFunction.max(sum, max)
  Next
  For i = 1 To 5
  For j = 1 To 5
  Cells(i + 1, j + 6) = matr(j, i) / max
  Next: Next
End Sub
