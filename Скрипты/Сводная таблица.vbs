Private Sub Auto_Update_PivotTable(ByVal Target As Range)
    Dim ws As Worksheet, pt As PivotTable

    On Error GoTo exitHandler
    Application.EnableEvents = False

    Debug.Print " Auto_Update_PivotTable вызвано в " & Now
    Debug.Print "Изменяемая область: " & Target.Address

    ' Проверяем, затронута ли таблица "Таблица1"
    If Not Intersect(Target, Me.Range("Таблица1[#All]")) Is Nothing Then
        Debug.Print "Изменение обнаружено в области Таблица1"
        
        ' Обновляем все сводные таблицы во всех листах
        For Each ws In ThisWorkbook.Worksheets
            For Each pt In ws.PivotTables
                Debug.Print "Обновление сводной таблицы: " & pt.Name & " на листе " & ws.Name
                pt.RefreshTable
            Next pt
        Next ws
    Else
        Debug.Print "Изменение не связано с Таблица1 - никаких действий не выполнено."
    End If

exitHandler:
    Application.EnableEvents = True
    Debug.Print " Auto_Update_PivotTable завершён в " & Now & " "
End Sub
