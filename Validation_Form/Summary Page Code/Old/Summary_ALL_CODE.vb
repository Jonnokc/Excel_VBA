Private Sub Delete_Extra_Sheets()
'Deletes sheets needed for this program. This is important if this needs to be run again.

Dim sheet As Worksheet

Application.DisplayAlerts = False

  For Each sheet In Worksheets

    If sheet.name = "Unmapped_Summary_Pivot" _
     Or sheet.name = "Validated_Summary_Pivot" _
     Or sheet.name = "Clinical_Summary_Pivot" _
     Or sheet.name = "Combined Registry Measures" _
     Then
     sheet.Delete
    End If
  Next sheet

Application.DisplayAlerts = True

End Sub


Private Sub Summary_Cleanup()

Dim sheet As Worksheet
Dim Col_Header_Rng As Variant
Dim HeaderRange As Variant 'Declare array variable
Dim HeaderArray As Variant
Dim i As Long 'The row variable
Dim Icol As Integer 'The column variable if you need to loop through multiple columns
Dim CurrentHeader As Variant 'Variable used to store column value

	'Clears values already in the table incase this needs to be rerun.
	Sheets("Summary View").Select
	Range("B1:L1").Select
	Selection.Name = "Summary_Headers"


	HeaderArray = Array("Registry", "Measure","Validated Mappings", "Unmapped Codes", "Clinical Documentation", "Health Maintenance")

  'HeaderRange = range("Summary_Headers").Value 'writes the named data range to the array variable

  For each cell in Range("Summary_Headers")'Loops through all rows within the range.

			CurrentHeader = cell 'Assigns current value to a variable
			IsInHeaderNameArray = Not IsError(Application.Match(CurrentHeader, HeaderArray, 0))

			'If Column is within the HeaderArray, Then clear the values in that column.
      If IsInHeaderNameArray = True Then
         cell.Offset(1,0).Select
				 Range(Selection, Selection.End(xlDown)).Select
				 Selection.Clear
      End If
  Next Cell

End Sub


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


Private Sub Clinical_Summary_Sheet_Pivot()
'
' Creates Clinical Documentation pivot table which is used for our summary color sheet
'

	Dim lastrow As Long
	Dim tbl As ListObject
	Dim sht As Worksheet
	Dim LastColumn As Long
	Dim StartCell As Range
	Dim rList As Range

	Sheets("Clinical Documentation").Select

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

	Set sht = Worksheets("Clinical Documentation")
	Set StartCell = Range("A2")

'Refresh UsedRange
	Worksheets("Clinical Documentation").UsedRange

'Find Last Row and Column
	LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
	sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

	Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
	tbl.Name = "Clinical_Table"
	tbl.TableStyle = "TableStyleLight12"

	'changes font color of header row to white
	Rows("1:1").Select
	With Selection.Font
		.ThemeColor = xlThemeColorDark1
		.TintAndShade = 0
	End With


'Creates a new sheet which will house the Clinical Documentation pivot table
	With ThisWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical_Summary_Pivot"
	End With


'Selects the data from the Clinical Documentation sheet which will be used in the pivot table
	Sheets("Clinical Documentation").Select
	ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
	"Clinical_Table", Version:=6).CreatePivotTable TableDestination:= _
	"Clinical_Summary_Pivot!R1C1", TableName:="Clinical_Summary_Pivot", DefaultVersion:=6
	Sheets("Clinical_Summary_Pivot").Select
	Cells(1, 1).Select

	ActiveWorkbook.ShowPivotTableFieldList = True

'Adds Registry Fild to pivot table
	With ActiveSheet.PivotTables("Clinical_Summary_Pivot").PivotFields("Registry")
		.Orientation = xlRowField
		.Position = 1
	End With

'Adds Count of Source to values fields
	ActiveSheet.PivotTables("Clinical_Summary_Pivot").AddDataField ActiveSheet.PivotTables( _
	"Clinical_Summary_Pivot").PivotFields("Source"), "Count of Source", xlCount

'Adds measure field to row at position 2
	With ActiveSheet.PivotTables("Clinical_Summary_Pivot").PivotFields("Measure")
		.Orientation = xlRowField
		.Position = 2
	End With

	ActiveSheet.PivotTables("Clinical_Summary_Pivot").PivotSelect "Registry[All]", _
	xlLabelOnly + xlFirstRow, True

'Sets layout to outline
	ActiveSheet.PivotTables("Clinical_Summary_Pivot").RowAxisLayout xlOutlineRow

'sets repeat rows to TRUE
	ActiveSheet.PivotTables("Clinical_Summary_Pivot").RepeatAllLabels xlRepeatLabels

'Sets empty values to 0 which helps in a couple places! but also allows the below autofill to have a range reference'
	ActiveSheet.PivotTables("Clinical_Summary_Pivot").NullString = "0"

	Range("D1").Select

	lastrow = ActiveSheet.Range("C2").End(xlDown).Row

	Sheets("Clinical_Summary_Pivot").Select
	Range("D2").Select
	ActiveCell.Formula = "=IF(B3 <>"""",CONCATENATE(A3,""|"",B3),"""")"

	With ActiveSheet.Range("D2")
		.AutoFill Destination:=Range("D2:D" & lastrow&)
	End With

End Sub



Private Sub Summary_Combined_Lookup_Sheet()
'
' Takes the Registries, Measures and Concepts from the Unmapped and Validated Sheets and combinds them into one sheet.Then creates a CONCATENATE column for lookup.
'

	With ThisWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Combined Registry Measures"
	End With

'Takes all the concept measures and puts them on one sheet
	Sheets("Potential Mapping Issues").Select
	Range("Validated_Mappings_Table[[#Headers],[Registry]:[Concept]]").Select
	Range(Selection, Selection.End(xlDown)).Select
	Selection.Copy
	Sheets("Combined Registry Measures").Select
	Range("A1").Select
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
	:=False, Transpose:=False

	Sheets("Unmapped Codes").Select
	Range("B3:D3").Select
	Range(Selection, Selection.End(xlDown)).Select
	Application.CutCopyMode = False
	Selection.Copy
	Sheets("Combined Registry Measures").Select
	Selection.End(xlDown).Select
	ActiveCell.Offset(1, 0).Range("A1").Select
	ActiveSheet.Paste

	Range("D1").Select
	ActiveCell.Formula = "Concat"

	Range("E1").Select
	ActiveCell.Formula = "Validated_lookup"

	Range("F1").Select
	ActiveCell.Formula = "Unmapped_lookup"

	Range("G1").Select
	ActiveCell.Formula = "Clinical_lookup"

	Range("A1").Select
	Range(Selection, Selection.End(xlDown)).Select
	Range(Selection, Selection.End(xlToRight)).Select
	Application.CutCopyMode = False

	Dim Ws As Worksheet
	Set Ws = ThisWorkbook.Sheets("Combined Registry Measures")

	Sheets("Combined Registry Measures").Select
	Range("A1:G" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

	Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "combined_lookup_range"
	Ws.ListObjects("combined_lookup_range").TableStyle = "TableStyleLight12"

	Sheets("Combined Registry Measures").Select
	ActiveSheet.Range("combined_lookup_range[#All]").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes

	Range("D2").Select
	ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2)"
	Range("D3").Select
	Columns("D:D").EntireColumn.AutoFit

	Range("E2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Validated_Summary_Pivot!C:C,MATCH(D2,Validated_Summary_Pivot!D:D,0)),0)"

	Range("F2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Unmapped_Summary_Pivot!C:C,MATCH(D2,Unmapped_Summary_Pivot!D:D,0)),0)"

	Range("G2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Clinical_Summary_Pivot!C:C,MATCH(D2,Clinical_Summary_Pivot!D:D,0)),0)"

End Sub



Private Sub Summary_Sheet_Initial_Setup()
'
'
'


		Dim tbl As ListObject

		Cells.Select
		Selection.ClearFormats

		Sheets("Combined Registry Measures").Select
		Range("A2:B2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Summary View").Select

		Range("B2").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

		Sheets("Summary View").Select

		Range("L1").Select
		ActiveCell.Formula = "Concat"

		Range("L2").Select
		ActiveCell.Formula = "=CONCATENATE(B2,""|"",C2)"
		Range("L3").Select

		Range("B1").Select
		Range(Selection, Selection.End(xlToRight)).Select
		Range(Selection, Selection.End(xlDown)).Select


		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Summary_Table"
		tbl.TableStyle = "TableStyleLight13"

		Range("L2").Select
		Application.CutCopyMode = False
		Selection.AutoFill Destination:=Range("Summary_Table[Concat]")
		Range("Summary_Table[Concat]").Select


End Sub



Private Sub Remove_Table_Format()

	Dim rList As Range

	Sheets("Potential Mapping Issues").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Sheets("Unmapped Codes").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Sheets("Clinical Documentation").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Range("A2").Select


End Sub


Sub Summary_Pop_Dots()
'
' Copies the temporary values from the lookup columns and pastes the VALUES into the appropriate columns.

	answer = MsgBox("This will populate the street lights on the Summary sheet. Leave computer alone until completed." & vbNextLine & "Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

	If answer = vbYes Then

		Application.ScreenUpdating = False

		Sheets("Summary View").Select
		Columns("E:H").Select
		With Selection
			.HorizontalAlignment = xlCenter
			.VerticalAlignment = xlBottom
			.WrapText = False
			.Orientation = 0
			.AddIndent = False
			.IndentLevel = 0
			.ShrinkToFit = False
			.ReadingOrder = xlContext
			.MergeCells = False
		End With

'refreshes the data on the pivot tables to make sure it is up to date.
		Sheets("Clinical_Summary_Pivot").Select 'Confirms Clinical Summary Pivot is up to date
		ActiveSheet.PivotTables("Clinical_Summary_Pivot").PivotCache.Refresh

		Sheets("Validated_Summary_Pivot").Select 'Confirms Validation Summary Pivot is up to date
		ActiveSheet.PivotTables("Validated_Pivot").PivotCache.Refresh

		Sheets("Unmapped_Summary_Pivot").Select 'Confirms Unmapped Summary Pivot is up to date
		ActiveSheet.PivotTables("Unmapped_Pivot").PivotCache.Refresh


'Copies the values from the combined lookup sheet onto the summary sheet to "hard code"
		Sheets("Combined Registry Measures").Select
		Range("E2:G2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Summary View").Select
		Range("E2").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

'Clears formatting from range
		Range("B2:K2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.ClearFormats

'Applies logic to range to populate the traffic lights
		Range("E2:H2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Application.CutCopyMode = False
		Selection.FormatConditions.AddIconSetCondition
		Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
		With Selection.FormatConditions(1)
			.ReverseOrder = True
			.ShowIconOnly = True
			.IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
		End With

		With Selection.FormatConditions(1).IconCriteria(2)
			.Type = xlConditionValueNumber
			.Value = 1
			.Operator = 7
		End With

		With Selection.FormatConditions(1).IconCriteria(3)
			.Type = xlConditionValueNumber
			.Value = 4
			.Operator = 7
		End With

'Aligns traffic light icons to center to improve appearance.
		With Selection
			.HorizontalAlignment = xlCenter
			.VerticalAlignment = xlBottom
			.WrapText = False
			.Orientation = 0
			.AddIndent = False
			.IndentLevel = 0
			.ShrinkToFit = False
			.ReadingOrder = xlContext
			.MergeCells = False
		End With

	Else

'Do Nothing

	End If

'Adds the hyperlink to the traffic lights to the corresponding sheet.
	Sheets("Summary View").Select
	Range("E2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Potential Mapping Issues'!A1"

	Range("F2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Unmapped Codes'!A1"

	Range("G2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Clinical Documentation'!A1"


'Formats the angle for the header row of Summary Sheet
	Rows("1:1").Select

	With Selection
		.VerticalAlignment = xlBottom
		.WrapText = False
		.Orientation = 45
		.AddIndent = False
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = False
	End With

'Autofit for all cells on screen.
	Cells.Select
	Cells.EntireColumn.AutoFit

	Application.ScreenUpdating = True 'Re-enables screen updating

End Sub



Sub Summary_Sheet_Setup()

	answer = MsgBox("This will launch the summary sheet scripts. Leave computer alone until completed. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

	If answer = vbYes Then

		Application.ScreenUpdating = False

		Call Delete_Extra_Sheets
		Call Summary_Cleanup
		Call Unmapped_Summary_Pivot
		Call Validated_Summary_Pivot
		Call Clinical_Summary_Sheet_Pivot
		Call Summary_Combined_Lookup_Sheet
		Call Summary_Sheet_Initial_Setup
		Call Remove_Table_Format


	Else
'do nothing
	End If

	Application.ScreenUpdating = True

	Sheets("Summary View").Select

End Sub
