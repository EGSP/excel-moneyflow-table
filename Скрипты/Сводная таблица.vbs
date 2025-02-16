Private Sub Auto_Update_PivotTable(ByVal Target As Range)
    Dim ws As Worksheet, pt As PivotTable

    On Error GoTo exitHandler
    Application.EnableEvents = False

    Debug.Print " Auto_Update_PivotTable ������� � " & Now
    Debug.Print "���������� �������: " & Target.Address

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
    Debug.Print " Auto_Update_PivotTable �������� � " & Now & " "
End Sub
