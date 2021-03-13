Private mywords$(), i%

Private Sub GetData()
    Dim file1%
    file1 = FreeFile
    Open "/Users/vitl/Desktop/lab8/inputs.txt" For Input As file1
    i = 1
    Do Until EOF(file1)
        ReDim Preserve mywords(1 To i)
        Line Input #file1, mywords(i)
        Range("a" & i) = mywords(i)
        i = i + 1
    Loop
    Лист1.Buttons("GetDataButton").Caption = "Файл данных успешно прочитан"
Close file1: End Sub

Private Sub SwapSymbols()
    Dim i As Integer, j As Integer
    i = 1
    Do While Range("A" & i).Text <> ""
        Range("B" & i) = Replace(Range("A" & i).Text, "x", "ks")
        i = i + 1
    Loop
    Лист1.Buttons("SwapSymbolsButton").Caption = "Символы успешно заменены"
End Sub

Private Sub SaveData()
    Dim LineOut$, FileOut%, j As Integer
    FileOut = FreeFile
    Open "/Users/vitl/Desktop/lab8/outputs.txt" For Output As FileOut
    j = 1
    Do While Range("B" & j).Text <> ""
        Print #FileOut, Range("B" & j).Text
        j = j + 1
    Loop
    Close FileOut
    Лист1.Buttons("SaveDataButton").Caption = "Данные успешно сохранены"
End Sub

Private Sub ClearList()
    Лист1.Range("A1:B" & i).Clear
    Лист1.Buttons("GetDataButton").Caption = "Прочитать файл данных"
    Лист1.Buttons("SwapSymbolsButton").Caption = "Заменить символы"
    Лист1.Buttons("SaveDataButton").Caption = "Сохранить данные в файл"
    Лист1.Buttons("ClearListButton").Caption = "Очистить лист"
End Sub
