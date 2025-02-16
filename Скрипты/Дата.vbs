Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim dCol As ListColumn
    Dim changedRow As Range
    Dim dateCell As Range

    ' ��������� � ����� � ���������� �������
    Debug.Print "Worksheet_Change ������ ��� ������/���������: " & Target.Address

    ' ������� �������� ������� "�������1"
    On Error Resume Next
    Set tbl = Me.ListObjects("�������1")
    On Error GoTo 0
    If tbl Is Nothing Then
        Debug.Print "������: ������� '�������1' �� �������!"
        Exit Sub
    Else
        Debug.Print "������� '�������1' �������."
    End If

    ' ������� �������� ������� "����"
    On Error Resume Next
    Set dCol = tbl.ListColumns("����")
    On Error GoTo 0
    If dCol Is Nothing Then
        Debug.Print "������: ������� '����' �� ������!"
        Exit Sub
    Else
        Debug.Print "������� '����' ������."
    End If

    ' ��������, ��� ��������� ��������� � ��������� ������ �������
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        Debug.Print "��������� ��������� � �������� ������ �������."

        ' �������� �� ���� ���������� �������
        Dim cell As Range
        Dim rowNumber As Long

        For Each cell In Target
            ' ���������� ����� ������ � �������
            rowNumber = cell.Row
            Set dateCell = tbl.DataBodyRange.Cells(rowNumber - tbl.DataBodyRange.Row + 1, dCol.Index)

            ' �������: �������� ������ � �������� ��������
            Debug.Print "������������ ������ " & rowNumber & _
                        ". ������ '����': " & dateCell.Address & _
                        " (������� ��������: '" & dateCell.Value & "')"

            ' ���� ������ �����, ��������� � ������� �����
            If Trim(dateCell.Value) = "" Then
                dateCell.Value = Date
                Debug.Print "� ������ " & dateCell.Address & " ����������� ����: " & Date
            Else
                Debug.Print "������ " & dateCell.Address & " ��� �������� ��������: " & dateCell.Value
            End If
        Next cell

        Application.EnableEvents = True
    Else
        Debug.Print "��������� �� ��������� � ��������� ������ �������."
    End If
End Sub

