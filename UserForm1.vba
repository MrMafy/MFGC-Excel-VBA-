Private Sub UserForm_Initialize()
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim arr() As Variant
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    
    ' Предполагая, что данные начинаются с F2
    Set rng = ThisWorkbook.Sheets("Ввод_данных").Range("F2:F" & ThisWorkbook.Sheets("Ввод_данных").Cells(Rows.Count, "F").End(xlUp).Row)
    
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
    
    Set wsData = ThisWorkbook.Sheets("Ввод_данных")
    Set wsRecord = ThisWorkbook.Sheets("Запись")
    Set wsOutput = ThisWorkbook.Sheets("Вывод")
    
    ' Очистить предыдущие фильтры
    wsData.AutoFilterMode = False
    
    ' Преобразование значения ComboBox1 в строку (если не является числом)
    Dim filterCriteria As Variant
    
    If IsNumeric(ComboBox1.Value) Then
        filterCriteria = CDbl(ComboBox1.Value)
    Else
        filterCriteria = ComboBox1.Value
    End If
    
    ' Отфильтровать столбец F на основе выбора ComboBox1
    'wsData.Range("F1").AutoFilter Field:=6, Criteria1:=filterCriteria
    wsData.Range("F1").AutoFilter Field:=6, Criteria1:="=" & filterCriteria & ""
    
    ' Найти последнюю строку в столбце F
    lastRow = wsData.Cells(wsData.Rows.Count, "F").End(xlUp).Row
    
    ' Вывести значение ComboBox1 в окно Immediate
    Debug.Print "Значение ComboBox1: " & ComboBox1.Value
    ' Вывести значение filterCriteria в окно Immediate
    Debug.Print "Значение filterCriteria: " & filterCriteria
    ' Вывести значение Criteria1 в окно Immediate
    Debug.Print "Значение Criteria1: " & wsData.AutoFilter.Filters(6).Criteria1

    
    ' Определение количества строк для копирования на основе выбранного значения
    If ComboBox1.Value = 0.75 Or ComboBox1.Value = 1 Or ComboBox1.Value = 1.5 Then
        rowCount = 10
    ElseIf ComboBox1.Value = 2.5 Then
        rowCount = 8
    ElseIf ComboBox1.Value = 4 Or ComboBox1.Value = 6 Or ComboBox1.Value = 10 Or ComboBox1.Value = 16 Then
        rowCount = 4
    ElseIf ComboBox1.Value >= 20 And ComboBox1.Value <= 100 Then
        rowCount = 6
    Else
        rowCount = 0 ' Значение по умолчанию
    End If
    
    ' Задать диапазон для копирования
    Set rngData = wsData.Range("J2:K" & lastRow)
    Set rngRecord = wsRecord.Range("E2:F" & lastRow)
    
    ' Очистить столбцы E и F на листе "Запись"
    'wsRecord.Range("E2:F" & lastRow).ClearContents
    wsRecord.Range("E:E").ClearContents
    wsRecord.Range("F:F").ClearContents
    
    ' Скопируйте видимые ячейки из столбцов J и K в столбцы E и F на листе "Запись"
    rngData.SpecialCells(xlCellTypeVisible).Copy wsRecord.Range("E2")
    
    ' Очистить столбец A на листе "Вывод"
    wsOutput.Range("A:A").ClearContents
    
    
    ' Скопировать видимые ячейки из столбцов E и F на листе "Запись" в столбец A на листе "Вывод"
    If Application.WorksheetFunction.Subtotal(103, rngRecord) > 1 Then
        ' Скопировать данные в столбец A на листе 'Вывод' в соответствии с выбранным значением
        For i = 1 To lastRow Step rowCount
            ' Выбрать блок видимых строк для копирования из столбца E
            On Error Resume Next
            Set rngBlockJ = rngRecord.Columns(1).SpecialCells(xlCellTypeVisible).Cells(i).Resize(rowCount, 1)
            On Error GoTo 0
            If Not rngBlockJ Is Nothing Then
                ' Скопировать видимые ячейки из столбца E в столбец A на листе 'Вывод'
                wsOutput.Range("A" & wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1).Resize(rowCount, 1).Value = rngBlockJ.Value
            End If

            ' Выбрать блок видимых строк для копирования из столбца F
            On Error Resume Next
            Set rngBlockK = rngRecord.Columns(2).SpecialCells(xlCellTypeVisible).Cells(i).Resize(rowCount, 1)
            On Error GoTo 0
            If Not rngBlockK Is Nothing Then
                ' Скопировать видимые ячейки из столбца F в столбец A на листе 'Вывод'
                wsOutput.Range("A" & wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1).Resize(rowCount, 1).Value = rngBlockK.Value
            End If
        Next i
    Else
        MsgBox "В столбце F нет видимых ячеек, удовлетворяющих выбранному значению в ComboBox1."
    End If
    
    ' Очистить фильтры
    wsData.AutoFilterMode = False
    
    ' Эта строка закрывает текущую форму (UserForm1)
    Unload Me
End Sub

