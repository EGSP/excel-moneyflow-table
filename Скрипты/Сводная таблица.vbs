Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet, pt As PivotTable

    On Error GoTo exitHandler
    Application.EnableEvents = False

    Debug.Print "=== Worksheet_Change triggered at " & Now & " ==="
    Debug.Print "Changed range: " & Target.Address

    ' ���������, ��������� �� ������� "�������1"
    If Not Intersect(Target, Me.Range("�������1[#All]")) Is Nothing Then
        Debug.Print "��������� ���������� � ������� �������1"
        
        ' ��������� ��� ������� ������� �� ���� ������
        For Each ws In ThisWorkbook.Worksheets
            For Each pt In ws.PivotTables
                Debug.Print "���������� ������� �������: " & pt.Name & " �� ����� " & ws.Name
                pt.RefreshTable
            Next pt
        Next ws
    Else
        Debug.Print "��������� �� ������� � �������1 - ������� �������� �� ���������."
    End If

exitHandler:
    Application.EnableEvents = True
    Debug.Print "=== Worksheet_Change �������� � " & Now & " ==="
End Sub
