Sub Create_Source_List()

	Dim lastrow As Long
	Dim tbl As ListObject
	Dim sht As Worksheet
	Dim LastColumn As Long
	Dim StartCell As Range
	Dim rList As Range
	Dim Next_Blank_Row  As Long

	For i = 1 To Worksheets.Count

		If Worksheets(i).Name = "Sources List" Then
			exists = True
		End If
		Next i

		If exists <> True Then
			With ThisWorkbook
				.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sources List"
			End With
		End If

		Sheets("Unmapped Codes").Select 'Selects the clinical doc sheet

'If AutoFilters are on turn them off
		If ActiveSheet.AutoFilterMode = True Then
			ActiveSheet.AutoFilterMode = False
		End If

'Checks the current sheet. If it is in table format, convert it to standard format.
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
		lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
		sht.Range(StartCell, sht.Cells(lastrow, LastColumn)).Select

		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Unmapped_Table"
		tbl.TableStyle = "TableStyleLight12"

		Sheets("Unmapped Codes").Select

		ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
		Criteria1:="<>"
		Range("E2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Sources List").Select
		Range("A1").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False


'Finds next blank row to add additional sources
		Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

		Sheets("Clinical Documentation").Select 'Selects the clinical doc sheet

'If AutoFilters are on turn them off
		If ActiveSheet.AutoFilterMode = True Then
			ActiveSheet.AutoFilterMode = False
		End If

'Checks the current sheet. If it is in table format, convert it to standard format.
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
		lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
		sht.Range(StartCell, sht.Cells(lastrow, LastColumn)).Select

		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Clinical_Table"
		tbl.TableStyle = "TableStyleLight12"

		ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
		Criteria1:="<>"
		Range("E2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Sources List").Select
		Range("A" & Next_Blank_Row).Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False
'Deletes the header row added during the copy and paste
		Rows(Next_Blank_Row & ":" & Next_Blank_Row).Select
		Application.CutCopyMode = False
		Selection.Delete Shift:=xlUp

'
    Set sht = Worksheets("Sources List")
    Set StartCell = Range("A1")

    'Refresh UsedRange
      Worksheets("Sources List").UsedRange

    'Find Last Row and Column
      LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
      sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

      Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
      tbl.Name = "Sources_Table"
      tbl.TableStyle = "TableStyleLight12"

      ActiveSheet.Range("Sources_Table[#All]").RemoveDuplicates Columns:=1, Header _
        :=xlYes

		'Create named range of the sources
    Range("A2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Name ="Sources_List"


	End Sub
