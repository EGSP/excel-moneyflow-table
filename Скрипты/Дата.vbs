Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim dCol As ListColumn
    Dim changedRow As Range
    Dim dateCell As Range

    ' ��������� � ����� � ���������� �������
    'MsgBox "Worksheet_Change ������ ��� ������/���������: " & Target.Address

    ' ������� �������� ������� "�������1"
    On Error Resume Next
    Set tbl = Me.ListObjects("�������1")
    On Error GoTo 0
    If tbl Is Nothing Then
        'MsgBox "������: ������� '�������1' �� �������!"
        Exit Sub
    Else
        'MsgBox "������� '�������1' �������."
    End If

    ' ������� �������� ������� "����"
    On Error Resume Next
    Set dCol = tbl.ListColumns("����")
    On Error GoTo 0
    If dCol Is Nothing Then
        'MsgBox "������: ������� '����' �� ������!"
        Exit Sub
    Else
        'MsgBox "������� '����' ������."
    End If

    ' ��������, ��� ��������� ��������� � ��������� ������ �������
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        'MsgBox "��������� ��������� � �������� ������ �������."

        ' �������� �� ���� ���������� �������
        Dim cell As Range
        Dim rowNumber As Long

        For Each cell In Target
            ' ���������� ����� ������ � �������
            rowNumber = cell.Row
            Set dateCell = tbl.DataBodyRange.Cells(rowNumber - tbl.DataBodyRange.Row + 1, dCol.Index)

            ' �������: �������� ������ � �������� ��������
            'MsgBox "������������ ������ " & rowNumber & _
                   ". ������ '����': " & dateCell.Address & _
                   " (������� ��������: '" & dateCell.Value & "')"

            ' ���� ������ �����, ��������� � ������� �����
            If Trim(dateCell.Value) = "" Then
                dateCell.Value = Date
                'MsgBox "� ������ " & dateCell.Address & " ����������� ����: " & Date
            Else
                'MsgBox "������ " & dateCell.Address & " ��� �������� ��������: " & dateCell.Value
            End If
        Next cell

        Application.EnableEvents = True
    Else
        'MsgBox "��������� �� ��������� � ��������� ������ �������."
    End If
End Sub
