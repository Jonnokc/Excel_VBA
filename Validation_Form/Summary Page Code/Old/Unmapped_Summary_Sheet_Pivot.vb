Private Sub Unmapped_Summary_Pivot()
'
' Creates unmapped pivot table which is used for our summary color sheet
'
Dim lastrow As Long
Dim tbl As ListObject
Dim sht As Worksheet
Dim LastColumn As Long
Dim StartCell As Range
Dim rList As Range

Sheets("Unmapped Codes").Select

If ActiveSheet.ListObjects.Count > 0 Then

	With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
	End With

	With rList
			.Interior.ColorIndex = xlColorIndexNone
			.Font.ColorIndex = xlColorIndexAutomatic
			.Borders.LineStyle = xlLineStyleNone
	End With

	End If

	Set sht = Worksheets("Unmapped Codes")
	Set StartCell = Range("A2")

	'Refresh UsedRange
	  Worksheets("Unmapped Codes").UsedRange

	'Find Last Row and Column
	  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

	'Select Range
	  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

	Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
	tbl.Name = "Unmapped_Table"
	tbl.TableStyle = "TableStyleLight12"

	'changes font color of header row to white
	Rows("1:1").Select
	With Selection.Font
		.ThemeColor = xlThemeColorDark1
		.TintAndShade = 0
	End With

'Creates a new sheet which will house the unmapped codes pivot table
	With ThisWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped_Summary_Pivot"
	End With

'Selects the data from the unmapped codes sheet which will be used in the pivot table
	Sheets("Unmapped Codes").Select
	Range("Unmapped_Table[[#Headers],[Registry]]").Select
	Range(Selection, Selection.End(xlToRight)).Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
	"Unmapped_Table", Version:=6).CreatePivotTable TableDestination:= _
	"Unmapped_Summary_Pivot!R1C1", TableName:="Unmapped_Pivot", DefaultVersion:=6
	Sheets("Unmapped_Summary_Pivot").Select
	Cells(1, 1).Select

	ActiveWorkbook.ShowPivotTableFieldList = True

'Adds Registry Fild to pivot table
	With ActiveSheet.PivotTables("Unmapped_Pivot").PivotFields("Registry")
		.Orientation = xlRowField
		.Position = 1
	End With

'Adds Count of Measure to values fields
	ActiveSheet.PivotTables("Unmapped_Pivot").AddDataField ActiveSheet.PivotTables( _
	"Unmapped_Pivot").PivotFields("Source"), "Count of Source", xlCount

'Adds measure field to row at position 2
	With ActiveSheet.PivotTables("Unmapped_Pivot").PivotFields("Measure")
		.Orientation = xlRowField
		.Position = 2
	End With

	ActiveSheet.PivotTables("Unmapped_Pivot").PivotSelect "Registry[All]", _
	xlLabelOnly + xlFirstRow, True

'Sets layout to outline
	ActiveSheet.PivotTables("Unmapped_Pivot").RowAxisLayout xlOutlineRow

'sets repeat rows to TRUE
	ActiveSheet.PivotTables("Unmapped_Pivot").RepeatAllLabels xlRepeatLabels


'Sets empty values to 0 which helps in a couple places! but also allows the below autofill to have a range reference'
	ActiveSheet.PivotTables("Unmapped_Pivot").NullString = "0"

	Range("D1").Select

	lastrow = ActiveSheet.Range("C2").End(xlDown).Row

	Sheets("Unmapped_Summary_Pivot").Select
	Range("D2").Select
	ActiveCell.Formula = "=IF(B3 <>"""",CONCATENATE(A3,""|"",B3),"""")"

	With ActiveSheet.Range("D2")
		.AutoFill Destination:=Range("D2:D" & lastrow&)
	End With

End Sub
