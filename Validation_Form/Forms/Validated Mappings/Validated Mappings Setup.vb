Sub Validated_Mappings_Setup()

  Dim Val_Headers As Variant
  Dim Rng As Range
  Dim cell As Range
  Dim tbl As ListObject
  Dim sht As Worksheet
  Dim Sheet As Worksheet
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim StartCell As Range


  'Helps improve performance
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  Val_Headers = Array("Measure", "Concept Alias", "Concepts Where Code is Normalized", "Standard Code System Display", "Raw Code Display", "Raw Code ID", "Raw Code System", "Registry", "Standard Code Display", "Standard Code ID")


  ' SUB - Stores the location of the columns by their header.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  ' Finds Column Header Locations for The Combined Data Export
  Sheet1.Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(Val_Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Val_Headers(i)) = LCase(Header) Then
        Val_Headers(i) = Mid(Header.Address, 2, 1)
        Header_Check = True
        Exit For
      End If
    Next Header
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Val_Headers(i) & "'" & " on the " & Sheet1.Name & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

      'If user hits cancel then close program.
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        Val_Headers(i) = Header_User_Response
      End If
    End If

  Next i


  ' SUB - CONVERTS Numbers stored as text to numbers
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Selects all cells in column
  Range(Val_Headers(5) & "2:" & Val_Headers(5) & Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Number_Check"
  Set Rng = Range("Number_Check")

  'If cell is a number, then convert format to number
  For Each cell In Rng
    If IsNumeric(cell) Then
      cell.Value = Val(cell.Value)
      cell.NumberFormat = "0"
    End If
  Next cell


  ' PRIMARY - CREATES PIVOT TABLE AND FINALIZES Workbook
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  ' SUB - Formats the Combined Sheet as a table
  Sheets("Combined").Select

  If Sheets("Combined").AutoFilterMode = True Then
    Sheets("Combined").AutoFilterMode = False
  End If

  ' Checks the current sheet. If it is in table format, convert it to range.
  If Sheets("Combined").ListObjects.Count > 0 Then
    With Sheets("Combined").ListObjects(1)
      Set rList = .Range
      .Unlist
    End With
  End If


  Set sht = Worksheets("Combined")
  Set StartCell = Range("A1")

  With sht
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
  End With

  With sht.Range("A2:A" & LastRow)
    If WorksheetFunction.CountBlank(.Cells) > 0 Then
      .SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    End If
  End With

  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  ' If User wants to clear all the formatting then clear all, else clear boarders
  If ClearFormatting = True Then
    Selection.ClearFormats
  Else
    Selection.Borders.LineStyle = xlNone
  End If

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Combined_Table"
  tbl.TableStyle = "TableStyleLight9"
  Columns.AutoFit
  Range("A1").Select

  ' Checks to see if the sheet already exists and if it does, deletes it.
  Application.DisplayAlerts = False

  For Each Sheet In Worksheets
    If Sheet.Name = "Validated_Pivot" Or Sheet.Name = "Validated_To_Review" Then
      Sheet.Delete
    End If
  Next Sheet

  Application.DisplayAlerts = True

  ' Re-creates the sheet
  With ThisWorkbook
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Validated_Pivot"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Validated_To_Review"
  End With

  ' Creates the Pivot Table
  ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
  "Combined_Table", Version:=6).CreatePivotTable TableDestination:= _
  "Validated_Pivot!R3C1", TableName:="Validated_Pivot", DefaultVersion:=6

  ' set the Pivot Table "EV_Pilot" to an Object
  Set PvtTbl = Sheets("Validated_Pivot").PivotTables("Validated_Pivot")
  With PvtTbl
    .ColumnGrand = False
    .RowGrand = False
    .RowAxisLayout xlTabularRow

    With .PivotFields("Standard Code Display")
      .Orientation = xlRowField
      .Position = 1
      .RepeatLabels = True
    End With
    With .PivotFields("Raw Code Display")
      .Orientation = xlRowField
      .Position = 2
      .RepeatLabels = True
    End With
    With .PivotFields("Concepts Where Code is Normalized")
      .Orientation = xlRowField
      .Position = 3
    End With
    With .PivotFields("Registry")
      .Orientation = xlRowField
      .Position = 4
    End With
    With .PivotFields("Measure")
      .Orientation = xlRowField
      .Position = 5
    End With
    With .PivotFields("Concept Alias")
      .Orientation = xlRowField
      .Position = 6
    End With
    With .PivotFields("Standard Code ID")
      .Orientation = xlRowField
      .Position = 7
    End With
    With .PivotFields("Raw Code ID")
      .Orientation = xlRowField
      .Position = 8
    End With
    With .PivotFields("Standard Code System Display")
      .Orientation = xlRowField
      .Position = 9
    End With
    With .PivotFields("Raw Code System")
      .Orientation = xlRowField
      .Position = 10
    End With

  End With


  ' Copies Pivot Table Data into a normal table for analysis
  With Sheets("Validated_Pivot")
    .Select
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
  End With

  ' Pastes the copied data into the Validated_To_Review Sheet

  With Sheets("Validated_To_Review")
    .Range("A1").PasteSpecial Paste:=xlPasteValues
  End With

  ' Formats Validated_To_Review Sheet As Table

  Sheets("Validated_To_Review").Select

  Set sht = Worksheets("Validated_To_Review")
  Set StartCell = Range("A1")

  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  'Turn selected Range Into Table
  Set tbl = Sheets("Validated_To_Review").ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Validated_To_Review"
  tbl.TableStyle = "TableStyleLight12"
  Columns.AutoFit
  Range("A1").Select


  ' Deletes the Pivot Table Sheet. It Is Not Needed
  Application.DisplayAlerts = False

  For Each Sheet In Worksheets
    If Sheet.Name = "Validated_Pivot" Then
      Sheet.Delete
    End If
  Next Sheet

  Application.DisplayAlerts = True


  'Re-enables settings
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  ' Exits Program and error handlings
  MsgBox ("Program is completed!")

  Exit Sub

User_Exit:

  'Re-enables settings
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  MsgBox ("Ut Oh! Something Went Wrong! Contact current code owner to find out what!")

End Sub
