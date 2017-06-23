Sub BORIS_PCST()

'
'
' PURPOSE: To create the PCST formatted files from the Validation Form.
' AUTHOR: Jonathan Adams
'
'
'

    Dim wb As Workbook
    Dim sht As Worksheet
    Dim Sheet As Worksheet
    Dim Table_Obj As ListObject
    Dim tbl As ListObject
    Dim Validation_File_Name As Variant
    Dim cValue As Variant
    Dim StartCell As Range
    Dim rList As Range
    Dim User_Name As String
    Dim Project_Name As String
    Dim Save_Path As String
    Dim Code_Sheet As String
    Dim Current_Value As String
    Dim Sheet_Name As String
    Dim Name_Input_Checker As Integer
    Dim Confirm_Scrubbed As Integer
    Dim Folder_Check As Integer
    Dim Project_Name_Checker As Integer
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim Next_Blank_Row As Long
    Dim Table_ObjIsVisible As Boolean

    'This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Confirm_Scrubbed = MsgBox("*NOTICE* It is highly advised that you review the data on the Unmapped Codes, Clinical Documentation, and the Health Maintenance Summary Sheet before running this program." & vbNewLine & vbNewLine & "You should delete unneeded lines and review concept endings to confirm the data is correct before proceeding. Otherwise errors will multiplied accross all newly created files.", vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Scrubbed = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        Exit Sub
    End If

    'Tells user the program is about to to and to leave their machine alone until the program is completed.
    MsgBox ("Program is about to run. This will take a minute or two to complete. During the process you will be asked for a couple inputs at the beginning and you may receive popups asking for permission to overwrite files if you have already ran this program once before. Otherwise please leave computer alone until completed popup appears.")

    'Names variable current file name
    Validation_File_Name = ActiveWorkbook.Name

    'Checks to confirm the user entered a correct user ID. This is needed for file save path.
    Do
        Name_Input_Checker = 0
        User_Name = InputBox("Please enter your Cerner userID." & vbNewLine & vbNewLine & "ex. BE042983")

        If User_Name = vbNullString Then
            MsgBox ("Canceling program per user action.")
            Exit Sub

        ElseIf Len(User_Name) <> 8 Then
            MsgBox ("That was not the correct format...Lets try this again..." & vbNewLine & "Please enter your user_ID. No spaces" & vbNewLine & vbNewLine & "ex. BE042983")

        Else
            Name_Input_Checker = 1
        End If

    Loop While Name_Input_Checker = 0


    'Checks to confirm user entered correct project name. This is needed for file name.
    Do
        Project_Name_Checker = 0
        Project_Name = InputBox("Please enter the abbreviation for this project." & vbNewLine & vbNewLine & "ex. NBRO")

        If Project_Name = vbNullString Then
            MsgBox ("Canceling program per user action.")

        ElseIf Len(Project_Name) = 4 Or Len(Project_Name) = 7 Then    'If length of user inut incorrect, prompt user to try again.
            Project_Name_Checker = 1
        Else
            MsgBox ("Lets try this again.... Please enter the project name..." & vbNewLine & vbNewLine & "ex. NBRO")
        End If

    Loop While Project_Name_Checker = 0

    'Assigns file save path to variable.
    Save_Path = "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"


    'If the folder already exists then do nothing. Else make it.
    If Len(Dir(Save_Path, vbDirectory)) = 0 Then
        MkDir Save_Path    'Creates the folder

    Else
        Folder_Check = MsgBox("Looks like the folder already exists... Do you want to continue?", vbOKCancel + vbQuestion, "Empty Sheet")    'Folder already exists so continuing on.
    End If

    'If user hits cancel on the folder check then cancel program.
    If Folder_Check = vbCancel Then
        MsgBox ("Canceling program per user action.")
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'CREATES THE SOURCE CODE SHEET AND TABLE FOR LOOP ON THE VALIDATION FORM
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Disables screen alert which would prompt user to confirm sheet deletion.
    Application.DisplayAlerts = False

    'Checks to see if sources list sheet already exists and if so deletes the worksheet so a new one can be created.
    For Each Sheet In Worksheets
        If Sheet.Name = "Sources List" Then
            exists = True
            Sheet.Delete
        End If
    Next Sheet

    're-enables screen alert after handling source code sheet deletion.
    Application.DisplayAlerts = True

    'Creates a new sheet titled 'Sources List'
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sources List"
    End With


    '''''''''''FORMATS THE UNMAPPED CODE SHEET FOR COPY TO THE NEW WORKBOOK'''''''''''''

    Sheets("Unmapped Codes").Select

    'If AutoFilters are on turn them off. ''''Table can not be created if autofilters are enabled within the range.''''
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
    End If

    'Checks the current sheet. If it is in table format, convert it to range.
    If ActiveSheet.ListObjects.Count > 0 Then
        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With
        'Reverts the color of the range back to standard.
        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With
    End If

    'Sets all cells on sheet to a table.
    Set sht = Worksheets("Unmapped Codes")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Unmapped Codes").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Converts range to table.
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Unmapped_Table"    'Names table
    tbl.TableStyle = "TableStyleLight12"    'Sets Table Style

    Sheets("Unmapped Codes").Select

    'Filters to remove blank lines
    ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
            Criteria1:="<>"

    'Selects and copies range
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sources List").Select
    Range("A1").Select

    'Pastes copied values without formatting
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


    'Finds next blank row to add additional sources.
    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

    Sheets("Clinical Documentation").Select    'Selects the clinical doc sheet

    'If AutoFilters are on turn them off
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
    End If

    'Checks the current sheet. If it is in table format, convert it to standard format.
    If ActiveSheet.ListObjects.Count > 0 Then

        'Converts table back to a range.
        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        'Removes color formatting and such from previous table.
        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With
    End If

    'Formats Clinical Documentation sheet as table
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

    'Filters to remove blank lines.
    ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
            Criteria1:="<>"

    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sources List").Select
    Range("A" & Next_Blank_Row).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

    'Selects the next blank row
    Rows(Next_Blank_Row & ":" & Next_Blank_Row).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    'Selects Sources List Sheet converts to table
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
    Selection.Name = "Sources_List"


    '''''''''''FORMATS HEALTHY INTENT WORKSHEET FOR COPY''''''''''''

    Sheets("Health Maintenance Summary").Select
    Worksheets("Health Maintenance Summary").AutoFilterMode = False    'Removes any autofilters on the page

    'Checks the current sheet. If it is in table format, convert it to standard format.
    If ActiveSheet.ListObjects.Count > 0 Then
        'convert the table back to a range
        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With
        'Reverts the colors back to normal.
        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With
    End If

    Range("K5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Health_Maint_Table"
    tbl.TableStyle = "TableStyleLight9"


    '
    ''''''''''''''''''''''''''''''''''''''''
    '         CREATE NEW WORKBOOK
    ''''''''''''''''''''''''''''''''''''''''
    '


    'Loop through the sources
    For Each Source_Name In Range("Sources_List")
        Set wb = Workbooks.Add    'Opens a new workbook

        'Saves the new workbook
        With NewBook
            ChDir "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"
            ActiveWorkbook.SaveAs Filename:= _
                    "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files\" & Source_Name, FileFormat:= _
                    xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        End With

        'Selects new workbook
        Windows(Source_Name & ".xlsm").Activate

        'Populates basic sheets on new workbook
        With ActiveWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped Codes"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical Documentation"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Health Maintenance Summary"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Source_Code_Systems"
        End With


        '''''''COPIES UNMAPPED CODES SHEET''''''''

        Windows(Validation_File_Name).Activate
        Sheets("Unmapped Codes").Select
        ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                Criteria1:="<>"
        Range("Unmapped_Table[[#Headers],[Status]]").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Application.CutCopyMode = False
        Selection.Copy

        'Selects the newly created excel file and pastes copied cells onto unmapped codes sheet
        Windows(Source_Name & ".xlsm").Activate
        Sheets("Unmapped Codes").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False    'Pastes values to prevent formatting errors

        Range("L1").Select
        Selection = "Code Short Name"

        Set sht = Worksheets("Unmapped Codes")
        Set StartCell = Range("A1")

        'Refresh UsedRange
        Worksheets("Unmapped Codes").UsedRange

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        'Selects range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        'Sets selected cells as a table and names the table
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = "Unmapped_Table"
        tbl.TableStyle = "TableStyleLight9"

        'Removes duplicates if any exist by Raw Code and Raw Display
        ActiveSheet.Range("Unmapped_Table[#All]").RemoveDuplicates Columns:=Array(7, 8 _
                ), Header:=xlYes

        ' Copies code id to short code id column
        ' Check to see there are any unmapped codes, if there are, then create code short name

        Set Table_Obj = ActiveSheet.ListObjects(1)

        'Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
        If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
            Table_ObjIsVisible = True
        Else
          Table_ObjIsVisible = False
        End If

        'If Table_ObjIsVisible = True then "X"
        If Table_ObjIsVisible = True Then
          Range("F2").Select
          Range(Selection, Selection.End(xlDown)).Select
          Selection.Copy
          Range("L2").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                  :=False, Transpose:=False

            'If Table_ObjIsVisible = False then "Y"
        Else

        End If


        'Applies sorting by Concept -> Measure -> Registry column
        ActiveWorkbook.Worksheets("Unmapped Codes").ListObjects("Unmapped_Table").Sort. _
                SortFields.Clear
        ActiveWorkbook.Worksheets("Unmapped Codes").ListObjects("Unmapped_Table").Sort. _
                SortFields.Add Key:=Range("Unmapped_Table[Concept]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Unmapped Codes").ListObjects("Unmapped_Table").Sort. _
                SortFields.Add Key:=Range("Unmapped_Table[Measure]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Unmapped Codes").ListObjects("Unmapped_Table").Sort. _
                SortFields.Add Key:=Range("Unmapped_Table[Registry]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        With ActiveWorkbook.Worksheets("Unmapped Codes").ListObjects("Unmapped_Table"). _
                Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With


        ''''''''''''LOOP THROUGH SHORT NAME COLUMN AND ADJUST TO SHORT NAME VERSION''''''''''''''

        'Select and assign range
        Range("L2:L" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select    'Selects all cells not empty in column
        Selection.Name = "Unmapped_Code_Short"    'Names Range

        'Assigns range to variable
        Set Rng = Range("Unmapped_Code_Short")


        ''''''''''Checks if CodeID is a cerner nomenclature code id and shorten name appropriately''''''''''

        For Each cell In Rng
            If InStr(cell, "urn:cerner:coding:codingsystem:nomenclature.source_vocab:") > 0 Then  'If cell contains "x"
                cell.Select
                cValue = cell.Value
                With Selection
                    cPlace = InStr(cell, "vocab")
                    Selection.Value = "nomenclature - " & Right(cValue, Len(cValue) - (cPlace + 5))    'Replace cell with nomenclature and all text after vocab
                End With

                'Checks if code id contains PTCARE and shortens appropriately
            ElseIf InStr(cell, "PTCARE") > 0 Then
                cell.Select
                cValue = cell.Value
                With Selection
                    cPlace = InStr(cValue, "vocab")
                    Selection.Value = Right(cValue, Len(cValue) - (cPlace + 5))
                End With

                'Checks if code id contains healthmaintenance and then shortens appropriately
            ElseIf InStr(cell, "healthmaintenance") > 0 Then
                cell.Select
                cValue = cell.Value
                With Selection
                    cPlace = InStr(cValue, "healthmaintenance")
                    Selection.Value = Right(cValue, Len(cValue) - (cPlace + 16))
                End With

                'Checks if code id is normal cerner code set and then shortens appropriately
            ElseIf InStr(cell, "urn:cerner:coding:codingsystem:codeset:") > 0 Then
                cell.Select
                cValue = cell.Value
                With Selection
                    cPlace = InStr(cValue, "codeset:")
                    Selection.Value = Right(cValue, Len(cValue) - (cPlace + 7))
                End With

                'Checks Catches alternate nomenclature code. This catches the general nomenclature code id's which do not contain the tail descriptor
            ElseIf InStr(cell, "urn:cerner:coding:codingsystem:nomenclature") > 0 Then
                cell.Select
                cValue = cell.Value
                With Selection
                    cPlace = InStr(cValue, "system:")
                    Selection.Value = Right(cValue, Len(cValue) - (cPlace + 6))
                End With
            End If
        Next cell


        ''''''''''''''Converts Nomenclature - PTCARE to Nomenclature - Patient Care to standardize naming convention.''''''''''''''

        For Each cell In Rng
            If cell = "nomenclature - PTCARE" Then
                cell.Select
                With Selection
                    Selection.Value = "Nomenclature - Patient Care"
                End With
            End If
        Next cell

        'Names the code short name range.
        Sheets("Unmapped Codes").Select
        Range("L2:L" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
        Selection.Name = "Code_Short"

        ''''''''''Checks Code Short length and if it is more than >28 characters then shorten the name.'''''''''

        For Each Code_Short In Range("Code_Short")
            Current_Value = Code_Short.Value
            If (Len(Current_Value) > 28) Then
                Code_Short.ClearContents
                Code_Short.Value = Left(Current_Value, 28)
            End If

            'Checks code short name for invalid characters and replaces them with a space
            If InStr(Current_Value, ":") > 0 Then
                Code_Short.ClearContents
                New_Value = Replace(Current_Value, ":", " ")
                Code_Short.Value = New_Value
            End If
        Next Code_Short


        ''''''''''''''POPULATES THE CLINICAL DOCUMENTATION SHEET'''''''''''''

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

        Set sht = Worksheets("Clinical Documentation")
        Set StartCell = Range("A1")

        'Refresh UsedRange
        Worksheets("Clinical Documentation").UsedRange

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        'Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = "Clinical_Table"
        tbl.TableStyle = "TableStyleLight9"

        'Removes dups by source,
        ActiveSheet.Range("Clinical_Table[#All]").RemoveDuplicates Columns:=Array(5, _
                10, 11), Header:=xlYes

        'Removes duplicates by source, nomenclature, and nomenclature ID
        ActiveSheet.Range("Clinical_Table[#All]").RemoveDuplicates Columns:=Array(5, _
                15, 16), Header:=xlYes

        Range("E2").Select

        LastRow = ActiveSheet.Range("E2").End(xlDown).Row

        With ActiveSheet.Range("S2")
            .AutoFill Destination:=Range("S2:S" & LastRow&)
        End With

        'Applies sorting by Concept -> Measure -> Registry
        Range("Clinical_Table[[#Headers],[Source]]").Select
        ActiveWorkbook.Worksheets("Clinical Documentation").ListObjects( _
                "Clinical_Table").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Clinical Documentation").ListObjects( _
                "Clinical_Table").Sort.SortFields.Add Key:=Range("Clinical_Table[Concept]"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Clinical Documentation").ListObjects( _
                "Clinical_Table").Sort.SortFields.Add Key:=Range("Clinical_Table[Measure]"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Clinical Documentation").ListObjects( _
                "Clinical_Table").Sort.SortFields.Add Key:=Range("Clinical_Table[Registry]") _
                , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Clinical Documentation").ListObjects( _
                "Clinical_Table").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With


        ''''''''''POPULATES THE HEALTH MAINTENANCE SUMMARY SHEET''''''''''''''

        Windows(Validation_File_Name).Activate    'Go to Validation Form
        Sheets("Health Maintenance Summary").Select
        Range("K5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.Copy

        Windows(Source_Name & ".xlsm").Activate    'Go to new file
        Sheets("Health Maintenance Summary").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        Set sht = Worksheets("Health Maintenance Summary")
        Set StartCell = Range("A1")

        'Refresh UsedRange
        Worksheets("Health Maintenance Summary").UsedRange

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        'Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = "Health_Maint_Table"
        tbl.TableStyle = "TableStyleLight9"

        'Removes duplicates by source, code and code display
        ActiveSheet.Range("Health_Maint_Table[#All]").RemoveDuplicates Columns:=Array _
                (9, 10, 11), Header:=xlYes


        '''''''''''''CREATES THE SOURCE CODE SHEET'''''''''''''

        Sheets("Unmapped Codes").Select
        ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                Criteria1:=Source_Name, Operator:=xlAnd

        Range("L1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        'Pastes unmapped codes list onto the source code systems sheet
        Sheets("Source_Code_Systems").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        'Formats Source_Code_Systems Sheet
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

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Checks to determine how many unique code ID's there are for this source.
        ' If there is only 1 or no code id's, then set the code id source to just  codeset "72"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'If there are no unmapped codes for this source, then set code source to 72 only.
        Range("A2").Select
        If Selection.Value = "" Then
            Selection.Value = "72"
        End If

        Range("A3").Select
        If Selection.Value = "" Then
            Range("A2").Select
            Selection.Name = "Code_ID_List"

        Else
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Name = "Code_ID_List"
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '             Creates a new sheet and names the sheet with the current source
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Loops through each unique Code ID for this source and creates a sheet with the relavant data.
        For Each code In Range("Code_ID_List")

            With ActiveWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
            End With

            Code_Sheet = code

            'Special instructions for code set 72
            If code = "72" Then

                Sheets(Code_Sheet).Select

                'Populates sheet 72 headers
                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "EventCode"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "EventDisplay"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "TaskAssay"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Notes"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select

                ''''''''''''Copies Clinical Documentation to 72'''''''''''''

                Sheets("Clinical Documentation").Select

                'Filters for only the current source
                ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd

                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks filtered table for visible data
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible, then copy visible data
                If Table_ObjIsVisible = True Then

                    Sheets("Clinical Documentation").Select
                    Columns("B:O").Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Range("A1").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Clinical Documentation").Select
                    Columns("Q:Q").Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Columns("O:O").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Clinical Documentation").Select
                    Columns("R:R").Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Range("P1").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                ''''''''''Copies unmapped codes to 72 sheet''''''''''

                Sheets("Unmapped Codes").Select

                'Applies filters for only this source and code being currently reviewed.
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=12, _
                        Criteria1:=code, Operator:=xlAnd

                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks table for visible data
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible then copy data
                If Table_ObjIsVisible = True Then

                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlToLeft)).Select
                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy

                    Sheets(Code_Sheet).Select

                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
                    'Selects next blank row
                    Range("A" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Unmapped Codes").Select

                    Range("G2:I2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Application.CutCopyMode = False
                    Selection.Copy

                    Sheets(Code_Sheet).Select
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                '''''''''Populates Health Maintenance to CS 72''''''''''''''

                Sheets("Health Maintenance Summary").Select

                ActiveSheet.ListObjects("Health_Maint_Table").Range.AutoFilter Field:=11, _
                        Criteria1:=Source_Name, Operator:=xlAnd

                Set tbl = ActiveSheet.ListObjects(1)

                'Checks table for visible data
                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                'If data is visible then copy visible data
                If Table_ObjIsVisible = True Then

                    'Copies Sources Column
                    Range("K2:K" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Uses The CodeSet 72 column "Sources" to determine next blank row
                    Next_Blank_Row = Range("D" & Rows.Count).End(xlUp).Row + 1

                    'Pastes Sources on new sheet
                    Range("D" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Expect_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("B2:B" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Expect_Meaning on new sheet
                    Range("F" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Satisfier_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("G2:G" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Satisfier_Meaning on new sheet
                    Range("G" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Entry_Type Column
                    Sheets("Health Maintenance Summary").Select
                    Range("C2:C" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Entry_Type on new sheet
                    Range("L" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Event_CD Column
                    Sheets("Health Maintenance Summary").Select
                    Range("I2:I" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Event_CD on new sheet
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Event_CD_DISP Column
                    Sheets("Health Maintenance Summary").Select
                    Range("J2:J" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Event_CD_DISP on new sheet
                    Range("J" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                '''''''''Populates headers for all other sheets''''''''''

            Else
                Sheets(Code_Sheet).Select

                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "Code"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "Display"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "vlookup"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Team"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select

                'Filters unmapped codes table for current source and code within loop.
                Sheets("Unmapped Codes").Select
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=12, _
                        Criteria1:=code, Operator:=xlAnd

                'Sets variable to the table on the active sheet.
                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks filtered table for visible data.
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible then copy data.
                If Table_ObjIsVisible = True Then

                    Sheets("Unmapped Codes").Select
                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy

                    Sheets(Code_Sheet).Select

                    'Finds next blank row
                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

                    'Pastes data on next blank row
                    Range("A" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Unmapped Codes").Select

                    Range("G2:I2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Application.CutCopyMode = False
                    Selection.Copy

                    Sheets(Code_Sheet).Select
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                End If

            End If

        Next code


        '''''''POPULATES NOMENCLATURE - PATIENT CARE SHEET IF NEEDED''''''''


        Sheets("Clinical Documentation").Select

        'Filters table for current source.
        ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
                Criteria1:=Source_Name, Operator:=xlAnd
        ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=15, _
                Criteria1:="<>"

        'Eventually update this to filter out all rows which ARE MAPPED CORRECTLY. To only leave
        'incorrect rows.
        'ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=18, _
         'Criteria1:="<>"

        Set Table_Obj = ActiveSheet.ListObjects(1)

        'Checks filtered table for visible data.
        If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
            Table_ObjIsVisible = True
        Else
            Table_ObjIsVisible = False
        End If

        'If data is visible then copy data.
        If Table_ObjIsVisible = True Then
            'Check to see if Nomenclature - Patient Care sheet already exists
            For Each Sheet In Worksheets
                If Sheet.Name = "Nomenclature - Patient Care" Then
                    exists = True
                    Exit For
                Else
                    exists = False
                End If
            Next Sheet

            'If sheet does NOT exist, then create the sheet
            If exists = False Then
                ActiveWorkbook.Sheets.Add(After:=Worksheets(1)).Name = "Nomenclature - Patient Care"

                'Populate Headers
                Sheets("Nomenclature - Patient Care").Select
                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "Code"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "Display"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "vlookup"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Team"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select
            End If

            'Populates the Nomenclature - Patient Care Sheet with data from Clinical Documentation
            Sheets("Clinical Documentation").Select
            Range("B2:Q2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy

            Sheets("Nomenclature - Patient Care").Select

            'Selects next blank row
            Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
            Range("A" & Next_Blank_Row).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

            Set StartCell = Range("A1")
            Set sht = Worksheets("Nomenclature - Patient Care")

            'Finds last row with text
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row


        Else
            'Do Nothing
        End If


        '''''''''''Loops through sheets and delets unneeded sheets''''''''''''

        'Disables on screen confirm prompt
        Application.DisplayAlerts = False

        For Each Sheet In Worksheets
            If Sheet.Name = "Unmapped Codes" _
                    Or Sheet.Name = "Health Maintenance Summary" _
                    Or Sheet.Name = "Clinical Documentation" _
                    Or Sheet.Name = "Source_Code_Systems" _
                    Or Sheet.Name = "Sheet1" _
                    Then
                Sheet.Delete
            End If
        Next Sheet

        're-enables on screen prompt
        Application.DisplayAlerts = True

        '''''''Creates Index Sheet'''''''''

        'Creates index sheet
        ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "Index Sheet"

        Range("A1").Select
        Selection.Value = "Index Sheet"
        ActiveCell.Offset(1, 0).Select    'Moves down a row

        For Each Sheet In Worksheets
            If Sheet.Name <> "Index Sheet" Then
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheet.Name & "'" & "!A1", TextToDisplay:=Sheet.Name
                ActiveCell.Offset(1, 0).Select    'Moves down a row
            End If
        Next Sheet

        ''''''''Loops through sheets and formats as table'''''''

        For Each Sheet In Worksheets
            'Activates current sheet
            Sheet.Activate

            Set sht = Sheet    'Sets value
            Set StartCell = Range("A1")    'Start cell used to determine where to begin creating the table range

            'Find Last Row and Column
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
            Sheet_Name = Sheet.Name    'Assigns sheet name to a variable as a string

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            'Creates the table
            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = Sheet_Name    'Names the table
            tbl.TableStyle = "TableStyleLight9"    'Sets table color theme
            Columns.AutoFit    'Autofits columns on sheet
            Range("A1").Select    'Selects Cell A1 on sheet. Completely cosmetic.
        Next Sheet

        ''''''''Aligns index sheet'''''''''

        Sheets("INDEX SHEET").Select
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With


        '''''''Saves the new workbook'''''''''

        Workbooks(Source_Name & ".xlsm").Close SaveChanges:=True
        Windows(Validation_File_Name).Activate    'Switches back to old workbook to begin next loop

    Next Source_Name    'Start over with a new source

    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    'Notifies user that the program has completed.
    MsgBox ("Your PCST Files have been created. Folder is loctated within your My Documents.")

End Sub
