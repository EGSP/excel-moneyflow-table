Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim dCol As ListColumn
    Dim changedRow As Range
    Dim dateCell As Range

    ' Сообщение о входе в обработчик события
    Debug.Print "Worksheet_Change вызван для ячейки/диапазона: " & Target.Address

    ' Попытка получить таблицу "Таблица1"
    On Error Resume Next
    Set tbl = Me.ListObjects("Таблица1")
    On Error GoTo 0
    If tbl Is Nothing Then
        Debug.Print "Ошибка: Таблица 'Таблица1' не найдена!"
        Exit Sub
    Else
        Debug.Print "Таблица 'Таблица1' найдена."
    End If

    ' Попытка получить столбец "Дата"
    On Error Resume Next
    Set dCol = tbl.ListColumns("Дата")
    On Error GoTo 0
    If dCol Is Nothing Then
        Debug.Print "Ошибка: Столбец 'Дата' не найден!"
        Exit Sub
    Else
        Debug.Print "Столбец 'Дата' найден."
    End If

    ' Проверка, что изменение произошло в диапазоне данных таблицы
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        Debug.Print "Изменение произошло в пределах данных таблицы."

        ' Проходим по всем затронутым строкам
        Dim cell As Range
        Dim rowNumber As Long

        For Each cell In Target
            ' Определяем номер строки в таблице
            rowNumber = cell.Row
            Set dateCell = tbl.DataBodyRange.Cells(rowNumber - tbl.DataBodyRange.Row + 1, dCol.Index)

            ' Отладка: проверка адреса и текущего значения
            Debug.Print "Обрабатываем строку " & rowNumber & _
                        ". Ячейка 'Дата': " & dateCell.Address & _
                        " (текущее значение: '" & dateCell.Value & "')"

            ' Если ячейка пуста, заполняем её текущей датой
            If Trim(dateCell.Value) = "" Then
                dateCell.Value = Date
                Debug.Print "В ячейку " & dateCell.Address & " установлена дата: " & Date
            Else
                Debug.Print "Ячейка " & dateCell.Address & " уже содержит значение: " & dateCell.Value
            End If
        Next cell

        Application.EnableEvents = True
    Else
        Debug.Print "Изменение не произошло в диапазоне данных таблицы."
    End If
End Sub

