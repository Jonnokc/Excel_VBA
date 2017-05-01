
Sub Code_System()

Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range

With ThisWorkbook
  .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Source_Code_Systems"
End With

Sheets("Unmapped Codes").Select


    Sheets("Unmapped Codes").Select
    ActiveSheet.ListObjects("Unmapped_Table").range.AutoFilter Field:=5, _
        Criteria1:="<>", Operator:=xlAnd

    Range("L1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("Source_Code_Systems").Select
    Range("A1").Select
    ActiveSheet.Paste


    Set sht = Worksheets("Source_Code_Systems")
    Set StartCell = Range("A1")

    'Refresh UsedRange
      Worksheets("Source_Code_Systems").UsedRange

    'Find Last Row and Column
      LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
      sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

      Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
      tbl.Name = "Code_ID_Table"
      tbl.TableStyle = "TableStyleLight9"

      Range("Code_ID_Table[[#Headers],[Code Short Name]]").Select
  Application.CutCopyMode = False
  ActiveSheet.Range("Code_ID_Table[#All]").RemoveDuplicates Columns:=1, Header:= _
      xlYes

      Range("A2").Select
      Range(Selection, Selection.End(xlDown)).Select
      Selection.Name = "Code_ID_List"

      Range("A1").Select
      Selection = "Unmapped Short Name"


End Sub
