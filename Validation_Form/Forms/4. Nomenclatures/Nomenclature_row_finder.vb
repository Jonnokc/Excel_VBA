Private Sub Nomenclature_Row_Finder()

  Dim wb As Workbook
  Dim Table_Obj As ListObject
  Dim Table_ObjIsVisible As Boolean
  Dim Visible_Rows_Count As Integer
  Dim i As Integer
  Dim LR As Long
  Dim Results_Range As Range
  Dim Val_Vis_Row As Range
  Dim StartCell As Range
  Dim WkNames As Variant
  Dim TblNames As Variant
  Dim ValMappingsHeaders As Variant
  Dim ValMappingsNumHeaders As Variant
  Dim ResultsHeaders As Variant
  Dim ResultsNumHeaders As Variant
  Dim ValidationSheetHeaders As Variant
  Dim ValidationSheetNumHeaders As Variant
  Dim NewLineHeaders As Variant
  Dim EventCode As Variant
  Dim NextBlank As Variant
  Dim Header_Check As Boolean


  'DEBUG

  WkNames = Array("Validated Mappings", "Results", "Validation Sheet")
  TblNames = Array("Mappings_Tbl", "Results_Tbl", "Val_Tbl")
  ValMappingsHeaders = Array("CODE ID", "MAPPING STATUS")
  ValMappingsNumHeaders = Array("CODE ID", "MAPPING STATUS")
  ResultsHeaders = Array("DTA_EC", "NOMEN_MNEM", "ALPHA_NOMEN_ID")
  ResultsNumHeaders = Array("DTA_EC", "NOMEN_MNEM", "ALPHA_NOMEN_ID")
  ValidationSheetHeaders = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Event Code Mapped?", "Nomenclature Mapped?")
  ValidationSheetNumHeaders = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Event Code Mapped?", "Nomenclature Mapped?")
  NewLineHeaders = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Event Code Mapped?", "Nomenclature Mapped?")

  ' SUB - Creates the EVent Code Mapped? And Nomenclature Mapped? Columns on the Validation SheetArray
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Sheets(WkNames(2)).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"


  ' SUB - Finds the Event Code Mapped? Column or adds it
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  For Each Header In Range("Header_row")
    Header_Check = False
    If LCase(Header) = LCase("Event Code Mapped?") Then
      columnaddress = Mid(Header.Address, 2, 1)
      Header_Check = True
      Exit For
    End If
  Next Header

  ' Creates the new column
  If Header_Check = False Then
    NewHeader = Mid(Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
    Range(NewHeader & "1") = "Event Code Mapped?"
  End If


  ' SUB - Finds the Nomenclature Mapped? Column or adds it
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For Each Header In Range("Header_row")
    Header_Check = False
    If LCase(Header) = LCase("Nomenclature Mapped?") Then
      columnaddress = Mid(Header.Address, 2, 1)
      Header_Check = True
      Exit For
    End If
  Next Header

  ' Creates the new column
  If Header_Check = False Then
    NewHeader = Mid(Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
    Range(NewHeader & "1") = "Nomenclature Mapped?"
  End If

  ' SUB - Finds column locations for Results sheet
  '''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Sheets(WkNames(1)).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(ResultsHeaders)
    ' Finds columns by header name
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Header) = LCase(ResultsHeaders(i)) Then
        ResultsHeaders(i) = Mid(Header.Address, 2, 1)
        ResultsNumHeaders(i) = Range(ResultsHeaders(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & ResultsHeaders(i) & " on the " & WkNames(1) & "Sheet" & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        ResultsHeaders(i) = UCase(Header_User_Response)
        ResultsNumHeaders(i) = Range(ResultsHeaders(i) & "1").Column
      End If
    End If
  Next i


  ' SUB - Finds column locations for Validated Mappings Headers sheet
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Sheets(WkNames(0)).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(ValMappingsHeaders)
    ' Finds columns by header name
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Header) = LCase(ValMappingsHeaders(i)) Then
        ValMappingsHeaders(i) = Mid(Header.Address, 2, 1)
        ValMappingsNumHeaders(i) = Range(ValMappingsHeaders(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & ValMappingsHeaders(i) & " on the " & WkNames(0) & "Sheet" & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        ValMappingsHeaders(i) = UCase(Header_User_Response)
        ValMappingsNumHeaders(i) = Range(ValMappingsHeaders(i) & "1").Column
      End If
    End If
  Next i


  ' SUB - Finds column locations for Validation Sheet sheet
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Sheets(WkNames(2)).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(ValidationSheetHeaders)
    ' Finds columns by header name
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Header) = LCase(ValidationSheetHeaders(i)) Then
        ValidationSheetHeaders(i) = Mid(Header.Address, 2, 1)
        ValidationSheetNumHeaders(i) = Range(ValidationSheetHeaders(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & ValidationSheetHeaders(i) & " on the " & WkNames(2) & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        ValidationSheetHeaders(i) = UCase(Header_User_Response)
        ValidationSheetNumHeaders(i) = Range(ValidationSheetHeaders(i) & "1").Column
      End If
    End If
  Next i


  ' SUB - Finds column locations for the New Lines Sheet
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Sheets("New Lines").Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(NewLineHeaders)
    ' Finds columns by header name
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Header) = LCase(NewLineHeaders(i)) Then
        NewLineHeaders(i) = Mid(Header.Address, 2, 1)
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & NewLineHeaders(i) & " on the " & WkNames(1) & "Sheet" & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        NewLineHeaders(i) = UCase(Header_User_Response)
      End If
    End If
  Next i


  ' This disables settings to improve macro performance.
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  ' SUB - Converts important sheets to TableStyle
  '''''''''''''''''''''''''''''''''''''''''''''''
  For i = 0 To UBound(WkNames)

    Sheets(WkNames(i)).Select

    If Sheets(WkNames(i)).AutoFilterMode = True Then
      Sheets(WkNames(i)).AutoFilterMode = False
    End If

    ' Checks the current sheet. If it is in table format, convert it to range.
    If Sheets(WkNames(i)).ListObjects.Count > 0 Then
      With Sheets(WkNames(i)).ListObjects(1)
        Set rList = .Range
        .Unlist
      End With
      ' Reverts the color of the range back to standard.
      With rList
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlLineStyleNone
      End With
    End If

    Set sht = Worksheets(WkNames(i))
    Set StartCell = Range("A1")

    ' Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
    Sheet_Name = WkNames(i)

    ' Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    ' Creates the table
    Set tbl = Sheets(WkNames(i)).ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = TblNames(i)
    tbl.TableStyle = "TableStyleLight12"

    Rows("1:1").Select
    With Selection.Font
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = 0
    End With


    ' Results Sheet - Adds formulas
    '''''''''''''''''''''''''''''''

    If WkNames(i) = "Results" Then

      ' Remove Duplicates from the Results Sheet
      Sheets(WkNames(1)).Range("Results_Tbl[#All]").RemoveDuplicates Columns:=Array(ResultsNumHeaders(0), ResultsNumHeaders(1), ResultsNumHeaders(2)), _
      Header:=xlYes

      ' Adds new Mapping note column
      Range("M1").Select
      Selection = "Mapping Note"

      ' Eventcode formula
      EventMapped = "=IFERROR(INDEX('Validated Mappings'!" & ValMappingsHeaders(1) & ":" & ValMappingsHeaders(1) & _
      ",MATCH(" & ResultsHeaders(0) & "2,'Validated Mappings'!" & ValMappingsHeaders(0) & ":" & ValMappingsHeaders(0) & ",0)),0)"

      ' Nomenclature formula
      NomenclatureMapped = "=IFERROR(INDEX('Validated Mappings'!" & ValMappingsHeaders(1) & ":" & ValMappingsHeaders(1) & _
      ",MATCH(" & ResultsHeaders(2) & "2,'Validated Mappings'!" & ValMappingsHeaders(0) & ":" & ValMappingsHeaders(0) & ",0)),0)"

      ' Event Code mapped?
      Range("A2").Select
      Selection.Formula = EventMapped

      ' Noemcnature mapped?
      Range("B2").Select
      Selection.Formula = NomenclatureMapped

      ' Both mapped?
      Range("C2").Select
      Selection.Formula = "=IF(AND(A2 =""Validated"", B2 = ""Validated""),""Both Validated"", 0)"

      ' Hides rows which are validated in both columns
      Sheets(WkNames(1)).ListObjects("Results_Tbl").Range.AutoFilter Field:=3, Criteria1:= _
      "0"

    End If


    ' If the sheet is the validation sheet, then remove duplicates and create range
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If WkNames(i) = "Validation Sheet" Then

      ' Removes Duplicates by Event Code from the Validation Sheet
      Sheets(WkNames(2)).Range("Val_Tbl[#All]").RemoveDuplicates Columns:=ValidationSheetNumHeaders(8), Header:= _
      xlYes

      LR = Range("A" & Rows.Count).End(xlUp).Row
      Range(ValidationSheetHeaders(8) & "2:" & ValidationSheetHeaders(8) & LR).SpecialCells(xlCellTypeVisible).Name = "Event_Codes"

    End If

  Next i


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '       PRIMARY - Loops through all event codes to identify lines that need to be handled
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For Each EventCode In Range("Event_Codes")

    ' Filters sheet by event code
    Sheets(WkNames(1)).Select
    Sheets(WkNames(1)).ListObjects("Results_Tbl").Range.AutoFilter Field:=ResultsNumHeaders(0), Criteria1:= _
    EventCode, Operator:=xlAnd

    Set Results_Range = Range("Results_Tbl")

    ' Error handling. If no codes are found, then skip the code.
    On Error GoTo NoBlanks

    ' Count number of visible rows on the Results sheet
    Visible_Rows_Count = Results_Range.SpecialCells(xlCellTypeVisible).Rows.Count

    ' Filters sheet by the current event code
    Sheets(WkNames(2)).Select
    Sheets(WkNames(2)).ListObjects("Val_Tbl").Range.AutoFilter Field:=ValidationSheetNumHeaders(8), Criteria1:= _
    EventCode, Operator:=xlAnd

    Validation_Visible_Row = Range("A2", Cells(Rows.Count, "A").End(xlUp)).SpecialCells(xlCellTypeVisible).Row

    Rows(Validation_Visible_Row).Select
    Selection.Copy

    Sheets("New Lines").Select

    ' Used to determine sheet location when replacing nomenclature values after new lines have been created.
    Code_Blank_Line = Range("A" & Rows.Count).End(xlUp).Row + 1


    '      SUB - Creates a new line for each "hit" for a specific code.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For i = 1 To Visible_Rows_Count
      ' Used to determine next blank line for copying the new validation line.
      Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

      Range("A" & Next_Blank_Row).Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Next i


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PRIMARY - Copies Columns From the Filtered Results Sheet to the New Lines Worksheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '   SUB - Copies the Alpha_Mon_ID column to the New Lines Sheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Sheets(WkNames(1)).Select

    ' Confirms active cell is within the table
    Range("A2").Select

    ' Selects the first visible cell in column the Alpha_Mon_ID column to the Nomenclature ID column on new lines
    Sheets(WkNames(1)).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, ResultsNumHeaders(2)).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("New Lines").Select

    Range(ValidationSheetHeaders(12) & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False


    '       SUB - Copies the Nomen_Source Column to the Nomenclature Notes column on the New Lines Sheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Switches back to Results Sheet to Copy next column
    Sheets(WkNames(1)).Select

    ' Confirms active cell is within the table
    Range("A2").Select

    ' Selects the first visible cell in the NOMEN_MNEM column
    Sheets(WkNames(1)).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, ResultsNumHeaders(1)).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("New Lines").Select

    Range(ValidationSheetHeaders(13) & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False


    '       SUB - Copies the Event Code Mapped? Column to the New Lines Sheet - Used when populating cmments
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Switches back to Results Sheet to Copy next column
    Sheets(WkNames(1)).Select

    ' Confirms active cell is within the table
    Range("A2").Select

    ' Selects the first visible cell in column Event Code Mapped? Column
    Sheets(WkNames(1)).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("New Lines").Select

    Range(NewLineHeaders(20) & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False


    '        SUB - Copies the Nomenclature Mapped? Column to the New Lines Sheet - Used when populating comments
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Switches back to Results Sheet to Copy next column
    Sheets(WkNames(1)).Select

    ' Confirms active cell is within the table
    Range("A2").Select

    ' Selects the first visible cell in the Nomenclature mapped? column
    Sheets(WkNames(1)).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("New Lines").Select

    Range(NewLineHeaders(21) & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False


    ' Error handling. No codes found, so skipping
NoBlanks:
    ' MsgBox("No Code for " & EventCode)
    Resume ClearError

ClearError:
    ' Clears variables for next loop
    Visible_Rows_Count = 0

  Next EventCode

  ' Re-enables previously disabled settings after all code has run.
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  Sheets("New Lines").Select

  Exit Sub

User_Exit:

  ' Re-enables previously disabled settings after all code has run.
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  MsgBox ("Exiting per user action")
  End


End Sub
