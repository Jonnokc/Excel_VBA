Private Sub Validated_Summary_Pivot()
'
' Creates Validated pivot table which is used for our summary color sheet
'
Dim lastrow As Long
Dim tbl As ListObject
Dim sht As Worksheet
Dim LastColumn As Long
Dim StartCell As Range
Dim rList As Range

Sheets("Potential Mapping Issues").Select

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

	Set sht = Worksheets("Potential Mapping Issues")
	Set StartCell = Range("A2")

	'Refresh UsedRange
	  Worksheets("Potential Mapping Issues").UsedRange

	'Find Last Row and Column
	  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

	'Select Range
	  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Validated_Mappings_Table"
		tbl.TableStyle = "TableStyleLight12"

		'changes font color of header row to white
		Rows("1:1").Select
		With Selection.Font
			.ThemeColor = xlThemeColorDark1
			.TintAndShade = 0
		End With

'Creates a new sheet which will house the validated codes pivot table
	With ThisWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Validated_Summary_Pivot"
	End With

	Sheets("Potential Mapping Issues").Select
	Range("Validated_Mappings_Table[[#Headers],[Registry]:[Concept]]").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
	"Validated_Mappings_Table", Version:=6).CreatePivotTable TableDestination:= _
	"Validated_Summary_Pivot!R1C1", TableName:="Validated_Pivot", DefaultVersion:=6
	Sheets("Validated_Summary_Pivot").Select
	Cells(1, 1).Select


	ActiveSheet.PivotTables("Validated_Pivot").AddDataField ActiveSheet.PivotTables( _
	"Validated_Pivot").PivotFields("Source"), "Count of Source", xlCount

	With ActiveSheet.PivotTables("Validated_Pivot").PivotFields("Registry")
		.Orientation = xlRowField
		.Position = 1
	End With


	With ActiveSheet.PivotTables("Validated_Pivot").PivotFields("Measure")
		.Orientation = xlRowField
		.Position = 2
	End With

'Sets pivot table layout to OUTLINE
	ActiveSheet.PivotTables("Validated_Pivot").RowAxisLayout xlOutlineRow


'Turns on repeat blank lines
	ActiveSheet.PivotTables("Validated_Pivot").RepeatAllLabels xlRepeatLabels

'Sets empty values to 0 which helps in a couple places! but also allows the below autofill to have a range reference'
	ActiveSheet.PivotTables("Validated_Pivot").NullString = "0"


	Range("D1").Select

	lastrow = ActiveSheet.Range("C2").End(xlDown).Row

	Sheets("Validated_Summary_Pivot").Select
	Range("D2").Select
	ActiveCell.Formula = "=IF(B3 <>"""",CONCATENATE(A3,""|"",B3),"""")"


	With ActiveSheet.Range("D2")
		.AutoFill Destination:=Range("D2:D" & lastrow&)
	End With

End Sub
