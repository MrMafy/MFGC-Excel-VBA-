VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���� ������ ������� �������"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3792
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim arr() As Variant
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    
    ' �����������, ��� ������ ���������� � F2
    Set rng = ThisWorkbook.Sheets("����_������").Range("F2:F" & ThisWorkbook.Sheets("����_������").Cells(Rows.Count, "F").End(xlUp).Row)
    
    For Each cell In rng
        If Not dict.Exists(cell.Value) And IsNumeric(cell.Value) Then
            dict.Add cell.Value, Nothing
        ElseIf Not dict.Exists(cell.Value) And IsError(cell.Value) = False Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    
    ComboBox1.Clear
    ReDim arr(0 To dict.Count - 1)
    i = 0
    For Each Key In dict.Keys
        If Key <> "" Then
            ComboBox1.AddItem Key
        End If
    Next Key
    
    'ComboBox1.List = Application.Transpose(arr)
    'ComboBox1.List = Application.WorksheetFunction.Sort(ComboBox1.List)

    ComboBox1.Text = ComboBox1.List(0)
End Sub

Private Sub ComboBox1_Change()
    If ComboBox1.ListIndex >= 0 Then
        ComboBox1.Text = ComboBox1.List(ComboBox1.ListIndex)
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim wsData As Worksheet
    Dim wsRecord As Worksheet
    Dim wsOutput As Worksheet
    Dim rngData As Range
    Dim rngRecord As Range
    Dim lastRow As Long
    Dim rowCount As Long
    Dim i As Long
    Dim rngBlockJ As Range
    Dim rngBlockK As Range
    
    Set wsData = ThisWorkbook.Sheets("����_������")
    Set wsRecord = ThisWorkbook.Sheets("������")
    Set wsOutput = ThisWorkbook.Sheets("�����")
    
    ' �������� ���������� �������
    wsData.AutoFilterMode = False
    
    ' �������������� �������� ComboBox1 � ������ (���� �� �������� ������)
    Dim filterCriteria As Variant
    
    If IsNumeric(ComboBox1.Value) Then
        filterCriteria = CDbl(ComboBox1.Value)
    Else
        filterCriteria = ComboBox1.Value
    End If
    
    ' ������������� ������� F �� ������ ������ ComboBox1
    'wsData.Range("F1").AutoFilter Field:=6, Criteria1:=filterCriteria
    wsData.Range("F1").AutoFilter Field:=6, Criteria1:="=" & filterCriteria & ""
    
    ' ����� ��������� ������ � ������� F
    lastRow = wsData.Cells(wsData.Rows.Count, "F").End(xlUp).Row
    
    ' ������� �������� ComboBox1 � ���� Immediate
    Debug.Print "�������� ComboBox1: " & ComboBox1.Value
    ' ������� �������� filterCriteria � ���� Immediate
    Debug.Print "�������� filterCriteria: " & filterCriteria
    ' ������� �������� Criteria1 � ���� Immediate
    Debug.Print "�������� Criteria1: " & wsData.AutoFilter.Filters(6).Criteria1

    
    ' ����������� ���������� ����� ��� ����������� �� ������ ���������� ��������
    If ComboBox1.Value = 0.75 Or ComboBox1.Value = 1 Or ComboBox1.Value = 1.5 Then
        rowCount = 10
    ElseIf ComboBox1.Value = 2.5 Then
        rowCount = 8
    ElseIf ComboBox1.Value = 4 Or ComboBox1.Value = 6 Or ComboBox1.Value = 10 Or ComboBox1.Value = 16 Then
        rowCount = 4
    ElseIf ComboBox1.Value >= 20 And ComboBox1.Value <= 100 Then
        rowCount = 6
    Else
        rowCount = 0 ' �������� �� ���������
    End If
    
    ' ������ �������� ��� �����������
    Set rngData = wsData.Range("J2:K" & lastRow)
    Set rngRecord = wsRecord.Range("E2:F" & lastRow)
    
    ' �������� ������� E � F �� ����� "������"
    'wsRecord.Range("E2:F" & lastRow).ClearContents
    wsRecord.Range("E:E").ClearContents
    wsRecord.Range("F:F").ClearContents
    
    ' ���������� ������� ������ �� �������� J � K � ������� E � F �� ����� "������"
    rngData.SpecialCells(xlCellTypeVisible).Copy wsRecord.Range("E2")
    
    ' �������� ������� A �� ����� "�����"
    wsOutput.Range("A:A").ClearContents
    
    
    ' ����������� ������� ������ �� �������� E � F �� ����� "������" � ������� A �� ����� "�����"
    If Application.WorksheetFunction.Subtotal(103, rngRecord) > 1 Then
        ' ����������� ������ � ������� A �� ����� '�����' � ������������ � ��������� ���������
        For i = 1 To lastRow Step rowCount
            ' ������� ���� ������� ����� ��� ����������� �� ������� E
            On Error Resume Next
            Set rngBlockJ = rngRecord.Columns(1).SpecialCells(xlCellTypeVisible).Cells(i).Resize(rowCount, 1)
            On Error GoTo 0
            If Not rngBlockJ Is Nothing Then
                ' ����������� ������� ������ �� ������� E � ������� A �� ����� '�����'
                wsOutput.Range("A" & wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1).Resize(rowCount, 1).Value = rngBlockJ.Value
            End If

            ' ������� ���� ������� ����� ��� ����������� �� ������� F
            On Error Resume Next
            Set rngBlockK = rngRecord.Columns(2).SpecialCells(xlCellTypeVisible).Cells(i).Resize(rowCount, 1)
            On Error GoTo 0
            If Not rngBlockK Is Nothing Then
                ' ����������� ������� ������ �� ������� F � ������� A �� ����� '�����'
                wsOutput.Range("A" & wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1).Resize(rowCount, 1).Value = rngBlockK.Value
            End If
        Next i
    Else
        MsgBox "� ������� F ��� ������� �����, ��������������� ���������� �������� � ComboBox1."
    End If
    
    ' �������� �������
    wsData.AutoFilterMode = False
    
    ' ��� ������ ��������� ������� ����� (UserForm1)
    Unload Me
End Sub

