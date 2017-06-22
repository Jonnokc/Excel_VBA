Sub Populate_New_Workbook()

	Dim Validation_File_Name As Variant
	Dim wb As Workbook
	Dim lastrow As Long

	Set wb = Workbooks.Add

	User_Name = "JA052464"
	Project_Name = "NBRO"
	Source_Name = "Mellinium"

'Saves the new workbook
	With NewBook
		ChDir "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"
		ActiveWorkbook.SaveAs Filename:= _
		"C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files\" & Source_Name, FileFormat:= _
		xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
	End With


	Validation_File_Name = InputBox("Insert the phsycical name of the current Validation Form File you are using." & vbNextLine & "ex. NBRO_FL Validation Form.xlsm")

	Windows(Source_Name & ".xlsm").Activate

	With ActiveWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Index Sheet"
	End With
	With ActiveWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped Codes"
	End With
	With ActiveWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical Documentation"
	End With
	With ActiveWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Health Maintenance Summary"
	End With

'Selects the Validation form Excel file and copies the data on the unmapped codes sheet to the new workbook.
	Windows(Validation_File_Name).Activate 'Selects the validation excel file
	Sheets("Unmapped Codes").Select
	ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
	Criteria1:="<>"
	Range("Unmapped_Table[[#Headers],[Status]]").Select
	Range(Selection, Selection.End(xlDown)).Select
	Range(Selection, Selection.End(xlToRight)).Select
	Application.CutCopyMode = False
	Selection.Copy

'Selects the newly created excel file and pastes copied cells into unmapped codes sheet
	Windows(Source_Name & ".xlsm").Activate
	Sheets("Unmapped Codes").Select
	Range("A1").Select
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
	:=False, Transpose:=False

	Range("L1").Select
	Selection = "Code Short Name"
	Range("L2").Select
	Selection.Formula = "=IF(ISNUMBER(SEARCH(""urn:cerner:coding"",F2)),TRIM(RIGHT(SUBSTITUTE(TRIM(F2),"":"",REPT("" "",LEN(TRIM(F2)))),LEN(TRIM(F2)))), F2)"

	Set sht = Worksheets("Unmapped Codes")
	Set StartCell = Range("A1")

'Refresh UsedRange
	Worksheets("Unmapped Codes").UsedRange

'Find Last Row and Column
	LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
	sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

	Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
	tbl.Name = "Unmapped_Table"
	tbl.TableStyle = "TableStyleLight9"

	Range("B1").Select

	lastrow = ActiveSheet.Range("B2").End(xlDown).Row

	With ActiveSheet.Range("L2")
		.AutoFill Destination:=Range("L2:L" & lastrow&)
	End With

'Goes back to validation form and copies the sheet to the new excel file
	Windows(Validation_File_Name).Activate
	Sheets("Clinical Documentation").Select
	ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
	Criteria1:="<>"
	Range("Clinical_Table[[#Headers],[Status]]").Select
	Range(Selection, Selection.End(xlDown)).Select
	Range(Selection, Selection.End(xlToRight)).Select
	Application.CutCopyMode = False
	Selection.Copy

'Navigates back to new workbook and pastes the copied rows.
	Windows(Source_Name & ".xlsm").Activate
	Sheets("Clinical Documentation").Select
	Range("A1").Select
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
	:=False, Transpose:=False

	'Adds code Short name column to unmapped codes. This is in place of the Code ID column to provide readability and useability.
	Range("S1").Select
	Selection = "Code Short Name"
	Range("S2").Select
	Selection.Formula = "=IF(ISNUMBER(SEARCH(""urn:cerner:coding"",E2)),TRIM(RIGHT(SUBSTITUTE(TRIM(E2),"":"",REPT("" "",LEN(TRIM(E2)))),LEN(TRIM(E2)))), E2)"

	Set sht = Worksheets("Clinical Documentation")
	Set StartCell = Range("A1")

'Refresh UsedRange
	Worksheets("Unmapped Codes").UsedRange

'Find Last Row and Column
	LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
	sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

	Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
	tbl.Name = "Clinical_Table"
	tbl.TableStyle = "TableStyleLight9"

	Range("E2").Select

	lastrow = ActiveSheet.Range("E2").End(xlDown).Row

	With ActiveSheet.Range("S2")
		.AutoFill Destination:=Range("S2:S" & lastrow&)
	End With






End Sub
