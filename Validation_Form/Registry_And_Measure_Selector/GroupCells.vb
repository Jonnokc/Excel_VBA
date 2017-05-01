Private Sub GroupCells()

    Dim myRange As Range
    Dim rowCount As Integer
    Dim currentRow As Integer
    Dim firstBlankRow As Integer
    Dim lastBlankRow As Integer
    Dim currentRowValue As String
    Dim neighborColumnValue As String

    'select range based on given named range
    Set myRange = Range("rngList")
    rowCount = Cells(Rows.Count, myRange.Column).End(xlUp).Row

    firstBlankRow = 0
    lastBlankRow = 0
    'for every row in the range
    For currentRow = 1 To rowCount
        currentRowValue = Cells(currentRow, myRange.Column).Value

        If (IsEmpty(currentRowValue) Or currentRowValue = "") Then
            'if cell is blank and firstBlankRow hasn't been assigned yet
            If firstBlankRow = 0 Then
                firstBlankRow = currentRow
            End If
        ElseIf Not (IsEmpty(currentRowValue) Or currentRowValue = "") Then
            'if the cell is not blank,
            'and firstBlankRow hasn't been assigned, then this is the firstBlankRow
            'to consider for grouping
            If firstBlankRow = 0 Then
                firstBlankRow = currentRow
            ElseIf firstBlankRow <> 0 And currentRowValue <> "" Then
                'if firstBlankRow is assigned and this row has a value then the cell one row above this one is to be considered
                'the lastBlankRow to include in the grouping
                lastBlankRow = currentRow - 1
            End If
        End If

        'if first AND last blank rows have been assigned, then create a group
        'then reset the first/lastBlankRow values to 0 and begin searching for next
        'grouping
        If firstBlankRow <> 0 And lastBlankRow <> 0 Then
            Range(Cells(firstBlankRow, myRange.Column), Cells(lastBlankRow, myRange.Column)).EntireRow.Select
            Selection.Group
            firstBlankRow = 0
            lastBlankRow = 0
        End If
    Next
End Sub
