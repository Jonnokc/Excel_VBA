Sub Unmapped_Codes_Sheet_Setup()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Application.ScreenUpdating = False

    Sheets("Unmapped Codes").Select

    Set sht = Worksheets("Unmapped Codes")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Unmapped Codes").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Unmapped_Table"
    tbl.TableStyle = "TableStyleLight9"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    Range("A2").Select

End Sub
