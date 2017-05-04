Sub BORIS_PCST()

  '
  '
  ' PURPOSE: Creates all the PCST files formatted properly.
  ' USE:    Open validation form which will be used to make PCST files. Run program and follow on screen prompts.
  ' AUTHOR: Jonathan Adams
  '
  '
  '
  '

  Dim wb As Workbook
  Dim sht As Worksheet
  Dim Sheet As Worksheet
  Dim tbl As ListObject
  Dim Validation_File_Name As Variant
  Dim cValue As Variant
  Dim Val_Wk_Array As Variant
  Dim Val_Tbl_Name_Array As Variant
  Dim CS_72_Header_Num_Array As Variant
  Dim CS_72_Header_Ltr_Array As Variant
  Dim Others_Header_Array As Variant
  Dim Clin_Doc_Col_Num_Array As Variant
  Dim Clin_Doc_Col_Ltr_Array As Variant
  Dim Clin_Doc_Col_Name_Array As Variant
  Dim Unmapped_Col_Ltr_Array As Variant
  Dim Unmapped_Col_Num_Array As Variant
  Dim Unmapped_Col_Name_Array As Variant
  Dim CS_72_Header_Name_Array As Variant
  Dim Gen_Sht_Header_Ltr_Array As Variant
  Dim Gen_Sht_Header_Num_Array As Variant
  Dim Gen_Sht_Header_Name_Array As Variant
  Dim Health_Maint_Num_Array As Variant
  Dim Health_Maint_Ltr_Array As Variant
  Dim Health_Maint_Name_Array As Variant
  Dim Header_User_Response As Variant
  Dim CurrentTable As Variant
  Dim Source_Name As Variant
  Dim NewBook As Variant
  Dim cell As Variant
  Dim start_cell As Variant
  Dim Header As Variant
  Dim Table_Obj As Variant
  Dim Rng As Variant
  Dim cPlace As Variant
  Dim UserNameErr As Variant
  Dim Code_Short As Variant
  Dim New_Value As Variant
  Dim code As Variant
  Dim ClearFormattingCheck As Variant
  Dim StartCell As Range
  Dim rList As Range
  Dim User_Name As String
  Dim Project_Name As String
  Dim Save_Path As String
  Dim Code_Sheet As String
  Dim Current_Value As String
  Dim Sheet_Name As String
  Dim CurrentSheet As String
  Dim Source_Combined As String
  Dim Current_Source As String
  Dim Name_Input_Checker As Integer
  Dim Confirm_Scrubbed As Integer
  Dim Folder_Check As Integer
  Dim Project_Name_Checker As Integer
  Dim Sources_Check As Integer
  Dim Off_Count As Integer
  Dim col_Count As Integer
  Dim ArraySize As Integer
  Dim i As Integer
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim Next_Blank_Row As Long
  Dim Table_ObjIsVisible As Boolean
  Dim Checker_Health_Maint As Boolean
  Dim First_Time As Boolean
  Dim Header_Check As Boolean
  Dim exists As Boolean
  Dim ClearFormatting As Boolean
  Dim LR As Long



  ' This disables settings to improve macro performance.
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False


  ' DEBUG

  ' Error Handling
  ' On Error GoTo ErrHandler

  ' Arrays used for sheet creation.

  Val_Wk_Array = Array("Clinical Documentation", "Unmapped Codes", "Health Maintenance Summary")
  Val_Tbl_Name_Array = Array("Clinical_Table", "Unmapped_Table", "Health_Maint_Table")

  CS_72_Header_Ltr_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssay", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")
  CS_72_Header_Num_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssay", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")
  CS_72_Header_Name_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssay", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")

  Gen_Sht_Header_Ltr_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "Code", "Display", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "vlookup", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")
  Gen_Sht_Header_Num_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "Code", "Display", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "vlookup", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")
  Gen_Sht_Header_Name_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "Code", "Display", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "vlookup", "Team", "Comments", "Standard Code", "Standard Coding System", "Internal Data Source")

  Clin_Doc_Col_Num_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Internal Data Source")
  Clin_Doc_Col_Ltr_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Internal Data Source")
  Clin_Doc_Col_Name_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssayCD", "Nomenclature Notes", "Social History Notes", "Grid Notes", "Freetext Notes", "Team", "Internal Data Source")

  Unmapped_Col_Num_Array = Array("Registry", "Measure", "Concept", "Source", "Code System", "Raw Code", "Raw Display", "Count", "Notes", "Team", "Code Short Name", "Internal Data Source")
  Unmapped_Col_Ltr_Array = Array("Registry", "Measure", "Concept", "Source", "Code System", "Raw Code", "Raw Display", "Count", "Notes", "Team", "Code Short Name", "Internal Data Source")
  Unmapped_Col_Name_Array = Array("Registry", "Measure", "Concept", "Source", "Code System", "Raw Code", "Raw Display", "Count", "Notes", "Team", "Code Short Name", "Internal Data Source")

  Health_Maint_Num_Array = Array("EXPECT_NAME", "EXPECT_MEANING", "ENTRY_TYPE", "EXPECT_SAT_ID", "EXPECT_SAT_NAME", "EXPECT_ID", "SATISFIER_MEANING", "PARENT_VALUE", "EVENT_CD", "EVENT_CD_DISP", "SOURCE", "Internal Data Source")
  Health_Maint_Ltr_Array = Array("EXPECT_NAME", "EXPECT_MEANING", "ENTRY_TYPE", "EXPECT_SAT_ID", "EXPECT_SAT_NAME", "EXPECT_ID", "SATISFIER_MEANING", "PARENT_VALUE", "EVENT_CD", "EVENT_CD_DISP", "SOURCE", "Internal Data Source")
  Health_Maint_Name_Array = Array("EXPECT_NAME", "EXPECT_MEANING", "ENTRY_TYPE", "EXPECT_SAT_ID", "EXPECT_SAT_NAME", "EXPECT_ID", "SATISFIER_MEANING", "PARENT_VALUE", "EVENT_CD", "EVENT_CD_DISP", "SOURCE", "Internal Data Source")

  'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
  Confirm_Scrubbed = MsgBox("*NOTICE* It is highly advised that you review the data on the Unmapped Codes, Clinical Documentation, and the Health Maintenance Summary Sheet before running this program." & vbNewLine & vbNewLine & "You should delete unneeded lines and review concept endings to confirm the data is correct before proceeding. Otherwise errors will be multiplied accross all newly created files." & vbNewLine & vbNewLine & "Once you click OK BORIS will start. Follow on screen prompts otherwise leave your computer alone until BORIS is done.", vbOKCancel + vbQuestion, "BORIS!")

  ' If user hits cancel then close program.
  If Confirm_Scrubbed = vbCancel Then
    GoTo User_Exit
  End If


  ' Names variable current file name
  Validation_File_Name = ActiveWorkbook.Name


  ' Checks to confirm user entered correct project name. This is needed for file name.
  Project_Name_Checker = 0
  Project_Name = InputBox("Please enter the abbreviation for this project." & vbNewLine & vbNewLine & "ex. NBRO")
  Do
    If Project_Name = vbNullString Then
      GoTo User_Exit
    ElseIf Len(Project_Name) = 4 Or Len(Project_Name) = 7 Then    'If length of user inut incorrect, prompt user to try again.
      Project_Name_Checker = 1
    Else
      Project_Name = InputBox("Why must we fight.... Lets try this again.... Please enter the project name..." & vbNewLine & vbNewLine & "ex. NBRO")
    End If

  Loop While Project_Name_Checker = 0

  ' Error handing if computer says path not found, send user back to try again on their username
Retry_UserID:

  ' Checks to confirm the user entered a correct user ID. This is needed for file save path.
  Name_Input_Checker = 0
  User_Name = InputBox("Please enter your Cerner userID." & vbNewLine & vbNewLine & "ex. BE042983", "BORIS will know...")
  Do

    If User_Name = vbNullString Then
      GoTo User_Exit

    ElseIf Len(User_Name) <> 8 Then
      User_Name = InputBox(" /Sigh... That was not the correct format...Lets try this again..." & vbNewLine & "Please enter your user_ID. No spaces" & vbNewLine & vbNewLine & "ex. BE042983", "#BORISForEmployeeOfTheMonth")

    Else
      Name_Input_Checker = 1
    End If

  Loop While Name_Input_Checker = 0


  ' Assigns file save path to variable.
  Save_Path = "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"


  ' If the folder already exists then do nothing. Else make it.
  If Len(Dir(Save_Path, vbDirectory)) = 0 Then
    ' On Error GoTo UserNameErr:
    MkDir Save_Path    'Creates the folder
  Else
    Folder_Check = MsgBox("Looks like the folder already exists... Do you want to continue?", vbOKCancel + vbQuestion, "BORIS!")    'Folder already exists so continuing on.
  End If

  ' If user hits cancel on the folder check then cancel program.
  If Folder_Check = vbCancel Then
    GoTo User_Exit
  End If

  ' Error handling for wrong user ID entered. If computer fails to find path, it is because username was wrong. Send user back to fix.
  If UserNameErr <> 0 Then
UserNameErr:
    MsgBox ("I think you entered your user ID wrong... Computer told me so. Sending you back to try again.")
    Resume Retry_UserID:
  End If

  ClearFormattingCheck = MsgBox("Is it ok to clear all cell formatting or do you want to keep cell coloring?", vbYesNo + vbQuestion, "BORIS!")
    If ClearFormattingCheck = vbYes Then
      ClearFormatting = True
    Else
      ClearFormatting = False
    End If




  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' PRIMARY - Formats worksheets for copying to new workbook
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' SUB - Unhides needed worksheets
  ''''''''''''''''''''''''''''''''''''
  For Each Sheet In Worksheets
    For i = 0 To UBound(Val_Wk_Array)
      If Val_Wk_Array(i) = Sheet.Name Then
        Sheet.Visible = xlSheetVisible
      End If
    Next i
  Next Sheet


  ' Switches Error handling back to normal
  ' On Error GoTo ErrHandler

  For i = 0 To UBound(Val_Wk_Array)


    Sheets(Val_Wk_Array(i)).Select

    ' table can not be created if autofilters are on
    If ActiveSheet.AutoFilterMode = True Then
      ActiveSheet.AutoFilterMode = False
    End If

    ' Checks the current sheet. If it is in table format, convert it to range.
    If ActiveSheet.ListObjects.Count > 0 Then
      With ActiveSheet.ListObjects(1)
        Set rList = .Range
        .Unlist
      End With
      'Reverts the color of the range back to standard.
      ' With rList
      '   .Interior.ColorIndex = xlColorIndexNone
      '   .Font.ColorIndex = xlColorIndexAutomatic
      '   .Borders.LineStyle = xlLineStyleNone
      ' End With
    End If

    ' Sets all cells on sheet to a table.
    Set sht = Worksheets(Val_Wk_Array(i))

    ' Health Maint Sheet data starts on a different row
    If Val_Wk_Array(i) <> "Health Maintenance Summary" Then
      Set StartCell = Range("A2")
      LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

      sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select
    Else
      Range("A5").Select
      Range(Selection, Selection.End(xlDown)).Select
      Range(Selection, Selection.End(xlToRight)).Select

    End If


    'Converts range to table.
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = Val_Tbl_Name_Array(i)
    tbl.TableStyle = "TableStyleLight12"

    'Filters to remove blank lines
    ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
    Criteria1:="<>"

  Next i

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' PRIMARY - Performs Checks on the Validation Form to Confirm Format is Correct Before Proceeding Any Further
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Checks the Health Maintenance Summary If format is incorrect, then tell user and end program
  If LCase(Sheets("Health Maintenance Summary").Range("K5").Value) <> "source" Then

    MsgBox ("BORIS has detected a possible error with the Validation Form layout" & vbNewLine & vbNewLine & _
    "BORIS expected Column K on the Health Maintenance Summary sheet to be 'Source'. He needs the source column for Health Maintenance. Please resolve the issue and then run again.")
    GoTo User_Exit
  End If


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'PRIMARY - CREATES THE SOURCE CODE SHEET AND TABLE FOR LOOP ON THE VALIDATION FORM
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Disables screen alert which would prompt user to confirm sheet deletion.
  Application.DisplayAlerts = False

  ' Checks to see if sources list sheet already exists and if so deletes the worksheet so a new one can be created.
  For Each Sheet In Worksheets
    If Sheet.Name = "Sources List" Then
      exists = True
      Sheet.Delete
    End If
  Next Sheet

  ' re-enables screen alert after handling source code sheet deletion.
  Application.DisplayAlerts = True

  'Creates a new sheet titled 'Sources List'
  With ThisWorkbook
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sources List"
    Range("A1").Value = "Sources"
  End With

  ' Assigns starting value to Next Blank Row
  Next_Blank_Row = Sheets("Sources List").Range("A" & Rows.Count).End(xlUp).Row + 1


  ' SUB - Loops through the sheets to find all the sources and put them on the sources list sheet
  For i = 0 To UBound(Val_Wk_Array)
    CurrentSheet = Val_Wk_Array(i)
    Sheets(CurrentSheet).Select

    ' Finds the sources column for copying
    If CurrentSheet <> Val_Wk_Array(2) Then
      Range("A1").Select
      Range("A2", Selection.End(xlToRight)).Name = "Header_row"

      For Each Header In Range("Header_row")
        Header_Check = False
        If LCase(Header) = LCase(CS_72_Header_Ltr_Array(3)) Then
          SourceCol = Mid(Header.Address, 2, 1)
          Exit For
        End If
      Next Header

      'Copies sources from the sheets to the sources list
      Range(ActiveSheet.Range(SourceCol & "2"), ActiveSheet.Range(SourceCol & "2").End(xlDown)).Copy Sheets("Sources List").Range("A" & Next_Blank_Row)
      Sheets("Sources List").Rows(Next_Blank_Row & ":" & Next_Blank_Row).Delete Shift:=xlUp
    Else

        Range("A5").Select
        Range("A5", Selection.End(xlToRight)).Name = "Header_row"

        For Each Header In Range("Header_row")
          Header_Check = False
          If LCase(Header) = LCase(CS_72_Header_Ltr_Array(10)) Then
            SourceCol = Mid(Header.Address, 2, 1)
            Exit For
          End If
        Next Header
      'Copies sources from the sheets to the sources list
      Range(ActiveSheet.Range(SourceCol & "5"), ActiveSheet.Range(SourceCol & "5").End(xlDown)).Copy Sheets("Sources List").Range("A" & Next_Blank_Row)
      Sheets("Sources List").Rows(Next_Blank_Row & ":" & Next_Blank_Row).Delete Shift:=xlUp
    End If
    Next_Blank_Row = Sheets("Sources List").Range("A" & Rows.Count).End(xlUp).Row + 1
  Next i

  Sheets("Sources List").Select
  Range("A1").Select
  Range(Selection, Selection.End(xlDown)).Select

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Sources_Table"
  tbl.TableStyle = "TableStyleLight12"

  ' Removes Duplicates from the sources table
  ActiveSheet.Range("Sources_Table[#All]").RemoveDuplicates Columns:=1, Header _
  :=xlYes

  LR = Range("A" & Rows.Count).End(xlUp).Row
  Range("A2" & ":A" & LR).SpecialCells(xlCellTypeVisible).Select
  Selection.Name = "Sources_List"


  '
  ' SUB - Shows list of all the sources and has user confirm the sources are correct
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For Each Source_Name In Range("Sources_List")
    Current_Source = Source_Name
    Source_Combined = Source_Combined & Current_Source & vbNewLine

  Next Source_Name

  ' Asks user to confirm the sources are correct before continuing.
  Sources_Check = MsgBox("BORIS found the following source(s). Please confirm that all sources are unique and there are no duplicates. If there are please click Cancel, rename the source(s) and then re-run the program. If the source(s) are good to go click OK to continue." & vbNewLine & vbNewLine & "List:" & vbNewLine & Source_Combined, vbOKCancel + vbQuestion, "There can only be one BORIS!")

  ' If user hits cancel then close program.
  If Sources_Check = vbCancel Then
    GoTo User_Exit
  End If



  ''''''''''''''''''''''''''''''''''''''''
  '   PRIMARY - Create New Workbook
  ''''''''''''''''''''''''''''''''''''''''

  ' Variable used to track column header location loop to only find column header locations once instead of repeating for each source.
  First_Time = True

  ' Loop through the sources and create the file for each source
  For Each Source_Name In Range("Sources_List")
    Set wb = Workbooks.Add

    ' Saves the new workbook
    With NewBook
      ChDir "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"
      ActiveWorkbook.SaveAs Filename:= _
      "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files\" & Source_Name
    End With

    Windows(Source_Name & ".xlsx").Activate

    ' Populates basic sheets on new workbook
    With ActiveWorkbook
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical Documentation"
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped Codes"
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Health Maintenance Summary"
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Source_Code_Systems"
    End With

    '
    ' SUB - Copies the data from the sheets in the validation form to the new workbook sheets
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Finds the start cell to begin range for copying. Different rules for Health Maint Sheet
    For i = 0 To UBound(Val_Wk_Array)
      CurrentSheet = Val_Wk_Array(i)
      CurrentTable = Val_Tbl_Name_Array(i)

      Windows(Validation_File_Name).Activate
      Sheets(CurrentSheet).Select

      ' Finds the location of the Registry Column
      If CurrentSheet <> Val_Wk_Array(2) Then
        Range("A2:U2").Name = "Header_row"

        For Each cell In Range("Header_row")
          If cell = "Registry" Then
            start_cell = cell.Address
            Exit For
          End If
        Next cell

      Else
        ' Changes range for the Health Maintenance Summary Sheet
        Range("A5:K5").Name = "Header_row"

        For Each cell In Range("Header_row")
          If cell = "EXPECT_NAME" Then
            start_cell = cell.Address
            Exit For
          End If
        Next cell

      End If

      ' Copies the data to the new sheet
      Sheets(CurrentSheet).Select
      ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
      Criteria1:="<>"
      Range(start_cell).Select
      Range(Selection, Selection.End(xlDown)).Select
      Range(Selection, Selection.End(xlToRight)).Select
      Selection.Copy

      ' Selects the newly created excel file and pastes copied cells onto unmapped codes sheet
      Windows(Source_Name & ".xlsx").Activate
      Sheets(CurrentSheet).Select
      Range("A1").Select
      ActiveSheet.Paste

      ' Sets header for code short name column
      If CurrentSheet = Val_Wk_Array(1) Then
        Range("K1").Value = "Code Short Name"
      End If
    Next i

    ' Formats the new sheets as tables
    For i = 0 To UBound(Val_Wk_Array)
      ' Confirms Selection of new workbook
      Windows(Source_Name & ".xlsx").Activate
      Sheets(Val_Wk_Array(i)).Select

      ' If table was not created when pasting values, then create the table
      If ActiveSheet.ListObjects.Count = 0 Then

        Set sht = Worksheets(Val_Wk_Array(i))
        Set StartCell = Range("A1")

        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        ' Converts range to table.
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = Val_Tbl_Name_Array(i)
        tbl.TableStyle = "TableStyleLight12"

        ' Filters to remove blank lines
        ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
        Criteria1:="<>"
      End If

    Next i

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PRIMARY - Finds Location of headers for the Main Sheets
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' All worksheets are identical no need to check column locations over and over again....
    If First_Time = True Then
      'Re-enables previously disabled settings after all code has run.
      Application.ScreenUpdating = True
      Application.EnableEvents = True

      ' Finds Column Header Locations for Clinical Documentation
      Sheets(Val_Wk_Array(0)).Select
      Range("A1").Select
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(Clin_Doc_Col_Num_Array)
        Header_Check = False
        For Each Header In Range("Header_row")
          If LCase(Clin_Doc_Col_Name_Array(i)) = LCase(Header) Then
            Clin_Doc_Col_Ltr_Array(i) = Mid(Header.Address, 2, 1)
            Clin_Doc_Col_Num_Array(i) = Range(Clin_Doc_Col_Ltr_Array(i) & "1").Column
            Header_Check = True
            Exit For
          End If
        Next Header
        If Header_Check = False Then
          Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Clin_Doc_Col_Name_Array(i) & "'" & " on the " & Val_Wk_Array(0) & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

          'If user hits cancel then close program.
          If Header_User_Response = vbNullString Then
            GoTo User_Exit
          Else
            Clin_Doc_Col_Ltr_Array(i) = Header_User_Response
            Clin_Doc_Col_Num_Array(i) = Range(Clin_Doc_Col_Ltr_Array(i) & "1").Column
          End If
        End If

      Next i

      ' Finds column Header Locations for Unmapped Columns
      Sheets(Val_Wk_Array(1)).Select
      Range("A1").Select
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(Unmapped_Col_Ltr_Array)
        Header_Check = False
        For Each Header In Range("Header_row")
          If LCase(Unmapped_Col_Name_Array(i)) = LCase(Header) Then
            Unmapped_Col_Ltr_Array(i) = Mid(Header.Address, 2, 1)
            Unmapped_Col_Num_Array(i) = Range(Unmapped_Col_Ltr_Array(i) & "1").Column
            Header_Check = True
            Exit For
          End If
        Next Header
        If Header_Check = False Then
          Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Unmapped_Col_Name_Array(i) & "'" & " on the " & Val_Wk_Array(1) & " Sheet....." & vbNewLine & vbNewLine & "However, However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

          'If user hits cancel then close program.
          If Header_User_Response = vbNullString Then
            GoTo User_Exit
          Else
            Unmapped_Col_Ltr_Array(i) = Header_User_Response
            Unmapped_Col_Num_Array(i) = Range(Unmapped_Col_Ltr_Array(i) & "1").Column
          End If
        End If
      Next i

      ' Finds Column Header Locations for Health Maintenance Summary
      Sheets(Val_Wk_Array(2)).Select
      Range("A1").Select
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(Health_Maint_Name_Array)
        Header_Check = False
        For Each Header In Range("Header_row")
          If LCase(Health_Maint_Name_Array(i)) = LCase(Header) Then
            Health_Maint_Ltr_Array(i) = Mid(Header.Address, 2, 1)
            Health_Maint_Num_Array(i) = Range(Health_Maint_Ltr_Array(i) & "1").Column
            Header_Check = True
            Exit For
          End If
        Next Header
        If Header_Check = False Then
          Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Health_Maint_Name_Array(i) & "'" & " on the " & Val_Wk_Array(2) & " Sheet....." & vbNewLine & vbNewLine & "However, However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

          ' If user hits cancel then close program.
          If Header_User_Response = vbNullString Then
            GoTo User_Exit
          Else
            Health_Maint_Ltr_Array(i) = Header_User_Response
            Health_Maint_Num_Array(i) = Range(Health_Maint_Ltr_Array(i) & "1").Column
          End If
        End If
      Next i

    End If

    ' Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' Sets value to prevent repeat of column header finders
    First_Time = False

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PRIMARY - Unmapped Remove Duplicates and Set Code Short Name
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Unampped codes removes duplicates by EvCode
    Sheets(Val_Wk_Array(1)).Range(Val_Tbl_Name_Array(1) & "[#All]").RemoveDuplicates Columns:=Array(Unmapped_Col_Num_Array(5)), Header:=xlYes

    '
    ' SUB - Code Short Name Creation
    ''''''''''''''''''''''''''''''''
    Set Table_Obj = Sheets(Val_Wk_Array(1)).ListObjects(1)

    ' Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
    Set tbl = Sheets(Val_Wk_Array(1)).ListObjects(1)

    If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
      Table_ObjIsVisible = True
    Else:
    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
  End If

  ' If there are visible rows, copy the Code System Column to the Code Short Name Column
  If Table_ObjIsVisible = True Then
    Sheets(Val_Wk_Array(1)).Select
    Range(Unmapped_Col_Ltr_Array(4) & "2:" & Unmapped_Col_Ltr_Array(4) & Cells.SpecialCells(xlCellTypeLastCell).Row).Copy Range(Unmapped_Col_Ltr_Array(10) & "2")
  End If

  ' Clears any autofilters
  Sheets(Val_Wk_Array(1)).AutoFilter.ShowAllData

  ' filters by concept -> registry for easy reviewing in final product
  With ActiveWorkbook.Sheets(Val_Wk_Array(1)).ListObjects(1).Sort
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Concept]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Measure]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Registry]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Apply
  End With

  '''''''''''''''''''''''''''''''''''''''''''''
  ' PRIMARY - Change Format of Unmapped Code Sheet Code Short Name
  '''''''''''''''''''''''''''''''''''''''''''''

  ' Selects Code Short Name Column and names the range for loop
  Sheets(Val_Wk_Array(1)).Range(Unmapped_Col_Ltr_Array(10) & "2:" & Unmapped_Col_Ltr_Array(10) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Code_Short"

  'Assigns range to variable
  Set Rng = Range("Code_Short")

  For Each cell In Rng
    If InStr(cell, "urn:cerner:coding:codingsystem:nomenclature.source_vocab:") > 0 Then
      cValue = cell.Value
      cPlace = InStr(cell, "vocab")
      cell.Value = "nomenclature - " & Right(cValue, Len(cValue) - (cPlace + 5))

      'Checks if code id contains PTCARE and shortens appropriately
    ElseIf InStr(cell, "PTCARE") > 0 Then
      cValue = cell.Value
      cPlace = InStr(cValue, "vocab")
      cell.Value = Right(cValue, Len(cValue) - (cPlace + 5))

      'Checks if code id contains healthmaintenance and then shortens appropriately
    ElseIf InStr(cell, "healthmaintenance") > 0 Then
      cValue = cell.Value
      cPlace = InStr(cValue, "healthmaintenance")
      cell.Value = Right(cValue, Len(cValue) - (cPlace + 16))

      'Checks if code id is normal cerner code set and then shortens appropriately
    ElseIf InStr(cell, "urn:cerner:coding:codingsystem:codeset:") > 0 Then
      cValue = cell.Value
      cPlace = InStr(cValue, "codeset:")
      cell.Value = Right(cValue, Len(cValue) - (cPlace + 7))

      'Checks Catches alternate nomenclature code. This catches the general nomenclature code id's which do not contain the tail descriptor
    ElseIf InStr(cell, "urn:cerner:coding:codingsystem:nomenclature") > 0 Then
      cValue = cell.Value
      cPlace = InStr(cValue, "system:")
      cell.Value = Right(cValue, Len(cValue) - (cPlace + 6))
    End If

  Next cell


  '   Converts Nomenclature - PTCARE to Nomenclature - Patient Care to standardize naming convention.
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For Each cell In Rng
    If cell = "nomenclature - PTCARE" Then
      cell.Value = "Nomenclature - Patient Care"
    End If
  Next cell

  ' Checks Code Short length and if it is more than >28 characters then shorten the name.
  For Each Code_Short In Range("Code_Short")
    Current_Value = Code_Short.Value
    If (Len(Current_Value) > 28) Then
      Code_Short.ClearContents
      Code_Short.Value = Left(Current_Value, 28)
    End If

    ' Checks code short name for invalid characters and replaces them with a space
    If InStr(Current_Value, ":") > 0 Then
      Code_Short.ClearContents
      New_Value = Replace(Current_Value, ":", " ")
      Code_Short.Value = New_Value
    End If
  Next Code_Short


  ' PRIMARY - Remove Duplicates Clinical Documentation, Filter For duplicates, Then Combines the sheets into one
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  ' filters by concept -> registry for easy reviewing in final product
  With ActiveWorkbook.Sheets(Val_Wk_Array(0)).ListObjects(1).Sort
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(0) & "[Concept]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(0) & "[Measure]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range(Val_Tbl_Name_Array(0) & "[Registry]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Apply
  End With

  ' SUB - Filters to exclude lines were the nomenclature is mapped
  Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(15), _
  Criteria1:= _
  "<>*This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure.*" _
  , Operator:=xlAnd, Criteria2:= _
  "<>*This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure.*"

  ' SUB - Creates the source code worksheet
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Windows(Source_Name & ".xlsx").Activate

  Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=4, _
  Criteria1:=Source_Name, Operator:=xlAnd

  ' Pastes the values from the Code Short Name onto the Source_Code_Systems Sheet
  Sheets(Val_Wk_Array(1)).Range(Unmapped_Col_Ltr_Array(10) & "1:" & Unmapped_Col_Ltr_Array(10) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Copy Sheets("Source_Code_Systems").Range("A1")

  Sheets("Source_Code_Systems").Select

  Set sht = Worksheets("Source_Code_Systems")
  Set StartCell = Range("A1")

  Worksheets("Source_Code_Systems").UsedRange

  ' Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Code_ID_Table"
  tbl.TableStyle = "TableStyleLight9"

  ' Removes Duplicates from the Code ID Lists
  Range("Code_ID_Table[[#Headers],[Code Short Name]]").Select
  Application.CutCopyMode = False
  ActiveSheet.Range("Code_ID_Table[#All]").RemoveDuplicates Columns:=1, Header:= _
  xlYes


  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' PRIMARY - Checks to determine how many unique code ID's there are for this source.
  ' If there is only 1 source "72" then set range to one cell, otherwise select all cells and set the range.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'If there are no unmapped codes for this source, then set code source to 72 only.
  If Range("A2") = "" Then
    Range("A2").Value = "72"
  End If

  ' If A3 is empty and thus only 1 code id, then set the range of Code_ID_List to just cell A2
  If Range("A3") = "" Then
    Range("A2").Name = "Code_ID_List"
  Else
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Name = "Code_ID_List"
  End If

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '      PRIMARY - Creates a new sheet and names the sheet with the current source
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Loops through each unique Code ID for this source and creates a sheet with the relavant data.
  For Each code In Range("Code_ID_List")

    With ActiveWorkbook
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
    End With

    Code_Sheet = code

    ' Special instructions for code set 72
    If code = "72" Then

      ' Populates the headers on the CS72 sheet
      Off_Count = 0
      For i = 0 To UBound(CS_72_Header_Name_Array)
        Sheets(Code_Sheet).Range("A1").Offset(0, Off_Count).Value = CS_72_Header_Name_Array(i)
        Off_Count = Off_Count + 1
      Next i

      ' Records the Addresses of the CS 72 headers
      Sheets(Code_Sheet).Select
      Range("A1").Select
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(CS_72_Header_Num_Array)
        col_Count = 0
        For Each Header In Range("Header_row")
          col_Count = col_Count + 1
          If LCase(CS_72_Header_Num_Array(i)) = LCase(Header) Then
            CS_72_Header_Ltr_Array(i) = Mid(Header.Address, 2, 1)
            CS_72_Header_Num_Array(i) = col_Count
            Exit For
          End If
        Next Header
      Next i


      ' SUB - Copies Clinical Documentation to 72
      ''''''''''''''''''''''''''''''''''''''''''''

      ' Filters The Source column for current source
      Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(3), _
      Criteria1:=Source_Name, Operator:=xlAnd


      'Checks current table to determine if any cells are visible to copy
      Sheets(Val_Wk_Array(0)).Select
      Set tbl = Sheets(Val_Wk_Array(0)).ListObjects(1)

      If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        Table_ObjIsVisible = True
      Else
        Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
      End If

      'If data is visible, and this is CS 72 then copy visible data
      If Table_ObjIsVisible = True Then

        ' Finds the last row of the Clinical Documentation Sheet for copy Range
        Sheets(Val_Wk_Array(0)).Select
        LR = Range("A" & Rows.Count).End(xlUp).Row

        ' Copies Registry Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(0) & "2:" & Clin_Doc_Col_Ltr_Array(0) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(0) & "2")

        ' Copies Measure Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(1) & "2:" & Clin_Doc_Col_Ltr_Array(1) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(1) & "2")

        ' Copies Concept Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(2) & "2:" & Clin_Doc_Col_Ltr_Array(2) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(2) & "2")

        ' Copies Source Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(3) & "2:" & Clin_Doc_Col_Ltr_Array(3) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(3) & "2")

        ' Copies DocumentType Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(4) & "2:" & Clin_Doc_Col_Ltr_Array(4) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(4) & "2")

        ' Copies Name Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(5) & "2:" & Clin_Doc_Col_Ltr_Array(5) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(5) & "2")

        ' Copies Section Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(6) & "2:" & Clin_Doc_Col_Ltr_Array(6) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(6) & "2")

        ' Copies DTA Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(7) & "2:" & Clin_Doc_Col_Ltr_Array(7) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(7) & "2")

        ' Copies EventCode Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(8) & "2:" & Clin_Doc_Col_Ltr_Array(8) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(8) & "2")

        ' Copies EventDisplay Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(9) & "2:" & Clin_Doc_Col_Ltr_Array(9) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(9) & "2")

        ' Copies ESH Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(10) & "2:" & Clin_Doc_Col_Ltr_Array(10) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(10) & "2")

        ' Copies ControlType Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(11) & "2:" & Clin_Doc_Col_Ltr_Array(11) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(11) & "2")

        ' ' Copies NomenclatureID Column
        ' Sheets(Val_Wk_Array(0)).Select
        ' Range(Clin_Doc_Col_Ltr_Array(12) & "2:" & Clin_Doc_Col_Ltr_Array(12) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(12) & "2")
        '
        ' ' Copies Nomenclature Display Column
        ' Sheets(Val_Wk_Array(0)).Select
        ' Range(Clin_Doc_Col_Ltr_Array(13) & "2:" & Clin_Doc_Col_Ltr_Array(13) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(13) & "2")
        '
        ' ' Copies TaskAssay Column
        ' Sheets(Val_Wk_Array(0)).Select
        ' Range(Clin_Doc_Col_Ltr_Array(14) & "2:" & Clin_Doc_Col_Ltr_Array(14) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(14) & "2")
        '
        ' ' Copies the Nomenclature Notes Column
        ' Sheets(Val_Wk_Array(0)).Select
        ' Range(Clin_Doc_Col_Ltr_Array(15) & "2:" & Clin_Doc_Col_Ltr_Array(15) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(15) & "2")

        ' Copies the Social History Notes Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(16) & "2:" & Clin_Doc_Col_Ltr_Array(16) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(16) & "2")

        ' Copies the Grid Notes Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(17) & "2:" & Clin_Doc_Col_Ltr_Array(17) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(17) & "2")

        ' Copies the Freetext Notes Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(18) & "2:" & Clin_Doc_Col_Ltr_Array(18) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(18) & "2")

        ' Copies the Team Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(19) & "2:" & Clin_Doc_Col_Ltr_Array(19) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(19) & "2")

        ' Copies the Internal Data Source Column
        Sheets(Val_Wk_Array(0)).Select
        Range(Clin_Doc_Col_Ltr_Array(20) & "2:" & Clin_Doc_Col_Ltr_Array(20) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(23) & "2")

      End If


      '    SUB - Copies unmapped codes to 72 sheet
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Sheets(Val_Wk_Array(1)).Select

      ' Applies filters for only this source and code being currently reviewed.
      Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(3), _
      Criteria1:=Source_Name, Operator:=xlAnd

      ' Filters The Code Short Name Column for current code in loop
      Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(10), _
      Criteria1:=code, Operator:=xlAnd

      Set tbl = Sheets(Val_Wk_Array(1)).ListObjects(1)

      If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        Table_ObjIsVisible = True
      Else
        Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
      End If

      ' If data is visible then copy data
      If Table_ObjIsVisible = True Then

        ' Finds the last row of the data sheet for copying
        Sheets(Val_Wk_Array(1)).Select
        LR = Range("A" & Rows.Count).End(xlUp).Row

        ' Finds the next blank row on the code sheet
        Sheets(Code_Sheet).Select
        Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

        ' Copies the Registry Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(0) & "2:" & Unmapped_Col_Ltr_Array(0) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(0) & Next_Blank_Row)

        ' Copies the Measure Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(1) & "2:" & Unmapped_Col_Ltr_Array(1) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(1) & Next_Blank_Row)

        ' Copies the Concept Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(2) & "2:" & Unmapped_Col_Ltr_Array(2) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(2) & Next_Blank_Row)

        ' Copies the Source Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(3) & "2:" & Unmapped_Col_Ltr_Array(3) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(3) & Next_Blank_Row)

        ' Copies the Code System ID Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(4) & "2:" & Unmapped_Col_Ltr_Array(4) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(4) & Next_Blank_Row)

        ' Copies the Raw Code Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(5) & "2:" & Unmapped_Col_Ltr_Array(5) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(8) & Next_Blank_Row)

        ' Copies the Raw Code Display Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(6) & "2:" & Unmapped_Col_Ltr_Array(6) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(9) & Next_Blank_Row)

        ' Copies the Count Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(7) & "2:" & Unmapped_Col_Ltr_Array(7) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(10) & Next_Blank_Row)

        ' Copies the Notes Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(8) & "2:" & Unmapped_Col_Ltr_Array(8) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(16) & Next_Blank_Row)

        ' Copies the Team Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(9) & "2:" & Unmapped_Col_Ltr_Array(9) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(15) & Next_Blank_Row)

        ' Copies the Internal Data Source Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(11) & "2:" & Unmapped_Col_Ltr_Array(11) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(23) & Next_Blank_Row)

      End If


      '       SUB - Populates Health Maintenance visible data to CS 72
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      Sheets(Val_Wk_Array(2)).Select

      Sheets(Val_Wk_Array(2)).ListObjects("Health_Maint_Table").Range.AutoFilter Field:=Health_Maint_Num_Array(10), _
      Criteria1:=Source_Name, Operator:=xlAnd

      Set tbl = Sheets(Val_Wk_Array(2)).ListObjects(1)

      If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        Table_ObjIsVisible = True
      Else
        Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
      End If

      ' If data is visible then copy visible data
      If Table_ObjIsVisible = True Then

        ' Finds the last row of the data sheet for copying
        Sheets(Val_Wk_Array(2)).Select
        LR = Range("A" & Rows.Count).End(xlUp).Row

        ' Sets next blank row
        Sheets(Code_Sheet).Select
        Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

        ' Copies the Source Column to Source
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(10) & "2:" & Health_Maint_Ltr_Array(10) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(3) & Next_Blank_Row)

        ' Copies Expect_Meaning Column to Name
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(1) & "2:" & Health_Maint_Ltr_Array(1) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(5) & Next_Blank_Row)

        ' Copies Satisfier_Meaning Column to Section
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(6) & "2:" & Health_Maint_Ltr_Array(6) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(6) & Next_Blank_Row)

        ' Copies Entry_Type Column to ControlType
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(2) & "2:" & Health_Maint_Ltr_Array(2) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(11) & Next_Blank_Row)

        ' Copies Event_CD Column to EventCode
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(8) & "2:" & Health_Maint_Ltr_Array(8) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(8) & Next_Blank_Row)

        ' Copies Event_CD_DISP Column to EventDisplay
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(9) & "2:" & Health_Maint_Ltr_Array(9) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(9) & Next_Blank_Row)

        ' Copies the Internal Data Source Column
        Sheets(Val_Wk_Array(2)).Select
        Range(Health_Maint_Ltr_Array(11) & "2:" & Health_Maint_Ltr_Array(11) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(23) & Next_Blank_Row)

      End If

      ' SUB - Removes duplicates from the CS 72 sheet
      ''''''''''''''''''''''''''''''''''''''''''''''''
      ' Dynamically determines size of array
      ArraySize = UBound(CS_72_Header_Ltr_Array)

      ' Removes duplicates by source, EVcode, EventDisplay
      Sheets(Code_Sheet).Range("$A$1:" & CS_72_Header_Ltr_Array(ArraySize) & LR).RemoveDuplicates Columns:=Array(CS_72_Header_Num_Array(3), _
      CS_72_Header_Num_Array(8), CS_72_Header_Num_Array(9)), _
      Header:=xlYes


      '       SUB - Populates headers for all other sheets
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ElseIf code <> "Nomenclature - Patient Care" Then

      Off_Count = 0
      For i = 0 To UBound(Gen_Sht_Header_Name_Array)
        Range("A1").Offset(0, Off_Count).Value = Gen_Sht_Header_Name_Array(i)
        Off_Count = Off_Count + 1
      Next i

      ' Records the Addresses of the Non-CS72 sheet headers
      Sheets(Code_Sheet).Select
      Range("A1").Select
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(Gen_Sht_Header_Ltr_Array)
        col_Count = 0
        For Each Header In Range("Header_row")
          col_Count = col_Count + 1
          If LCase(Gen_Sht_Header_Num_Array(i)) = LCase(Header) Then
            Gen_Sht_Header_Ltr_Array(i) = Mid(Header.Address, 2, 1)
            Gen_Sht_Header_Num_Array(i) = col_Count
            Exit For
          End If
        Next Header
      Next i


      Sheets(Val_Wk_Array(1)).Select

      ' Filters unmapped codes table for current source
      Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(3), _
      Criteria1:=Source_Name, Operator:=xlAnd

      ' Filters for current code within loop
      Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(10), _
      Criteria1:=code, Operator:=xlAnd

      ' Sets variable to the table on the active sheet.
      Set tbl = Sheets(Val_Wk_Array(1)).ListObjects(1)

      If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        Table_ObjIsVisible = True
      Else
        Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
      End If

      ' If data is visible then copy data.
      If Table_ObjIsVisible = True Then

        ' Finds the last row of the data sheet for copying
        Sheets(Val_Wk_Array(1)).Select
        LR = Range("A" & Rows.Count).End(xlUp).Row

        ' Finds next blank row
        Sheets(Code_Sheet).Select
        Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

        ' Copies the Registry Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(0) & "2:" & Unmapped_Col_Ltr_Array(0) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(0) & Next_Blank_Row)

        ' Copies the Measure Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(1) & "2:" & Unmapped_Col_Ltr_Array(1) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(1) & Next_Blank_Row)

        ' Copies the Concept Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(2) & "2:" & Unmapped_Col_Ltr_Array(2) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(2) & Next_Blank_Row)

        ' Copies the Source Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(3) & "2:" & Unmapped_Col_Ltr_Array(3) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(3) & Next_Blank_Row)

        ' Copies the Code System ID Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(4) & "2:" & Unmapped_Col_Ltr_Array(4) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(4) & Next_Blank_Row)

        ' Copies the Raw Code Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(5) & "2:" & Unmapped_Col_Ltr_Array(5) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(8) & Next_Blank_Row)

        ' Copies the Raw Code Display Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(6) & "2:" & Unmapped_Col_Ltr_Array(6) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(9) & Next_Blank_Row)

        ' Copies the Count Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(7) & "2:" & Unmapped_Col_Ltr_Array(7) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(10) & Next_Blank_Row)

        ' Copies the Notes Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(8) & "2:" & Unmapped_Col_Ltr_Array(8) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(16) & Next_Blank_Row)

        ' Copies the Team Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(9) & "2:" & Unmapped_Col_Ltr_Array(9) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(15) & Next_Blank_Row)

        ' Copies the Internal Data Source Column
        Sheets(Val_Wk_Array(1)).Select
        Range(Unmapped_Col_Ltr_Array(11) & "2:" & Unmapped_Col_Ltr_Array(11) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(Gen_Sht_Header_Ltr_Array(19) & Next_Blank_Row)

      End If

    End If

  Next code

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' PRIMARY - Populates Nomenclature ID data onto the Nomenclature - PTCARE sheet
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Sheets("Clinical Documentation").Select
  '
  ' 'Filters the clinical doc table by source and makes sure nomenclature ID
  ' Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(3), _
  ' Criteria1:=Source_Name, Operator:=xlAnd
  '
  ' ' Filters to make sure nomenclature ID is not blank
  ' Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(12), _
  ' Criteria1:="<>"
  '
  ' Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(15), _
  ' Criteria1:= _
  ' "<>*This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure.*" _
  ' , Operator:=xlAnd, Criteria2:= _
  ' "<>*This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure.*"
  '
  '
  ' Set Table_Obj = ActiveSheet.ListObjects(1)
  '
  ' 'Checks filtered table for visible data.
  ' If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
  '   Table_ObjIsVisible = True
  ' Else
  '   Table_ObjIsVisible = False
  ' End If
  '
  ' ' If data is visible then copy data.
  ' If Table_ObjIsVisible = True Then
  '   ' Check to see if Nomenclature - Patient Care sheet already exists
  '   For Each Sheet In Worksheets
  '     exists = False
  '     If Sheet.Name = "Nomenclature - Patient Care" Then
  '       exists = True
  '       Exit For
  '     End If
  '   Next Sheet
  '
  '   ' If sheet does NOT exist, then create the sheet and populate headers
  '   If exists = False Then
  '     ActiveWorkbook.Sheets.Add(After:=Worksheets(1)).Name = "Nomenclature - Patient Care"
  '
  '     ' Finds the last row of the Clinical Documentation Sheet for copy Range
  '     Sheets(Val_Wk_Array(0)).Select
  '     LR = Range("A" & Rows.Count).End(xlUp).Row
  '
  '     Code_Sheet = "Nomenclature - Patient Care"
  '
  '     ' Populates the headers on the PTCare sheet
  '     Off_Count = 0
  '     For i = 0 To UBound(CS_72_Header_Name_Array)
  '       Sheets(Code_Sheet).Range("A1").Offset(0, Off_Count).Value = CS_72_Header_Name_Array(i)
  '       Off_Count = Off_Count + 1
  '     Next i
  '   End If
  '
  '   ' Records the Addresses of the PTCare headers
  '   Sheets(Code_Sheet).Select
  '   Range("A1").Select
  '   Range("A1", Selection.End(xlToRight)).Name = "Header_row"
  '
  '   For i = 0 To UBound(CS_72_Header_Num_Array)
  '     col_Count = 0
  '     For Each Header In Range("Header_row")
  '       col_Count = col_Count + 1
  '       If LCase(CS_72_Header_Num_Array(i)) = LCase(Header) Then
  '         CS_72_Header_Ltr_Array(i) = Mid(Header.Address, 2, 1)
  '         CS_72_Header_Num_Array(i) = col_Count
  '         Exit For
  '       End If
  '     Next Header
  '   Next i
  '
  '   ' Copies Registry Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(0) & "2:" & Clin_Doc_Col_Ltr_Array(0) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(0) & "2")
  '
  '   ' Copies Measure Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(1) & "2:" & Clin_Doc_Col_Ltr_Array(1) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(1) & "2")
  '
  '   ' Copies Concept Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(2) & "2:" & Clin_Doc_Col_Ltr_Array(2) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(2) & "2")
  '
  '   ' Copies Source Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(3) & "2:" & Clin_Doc_Col_Ltr_Array(3) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(3) & "2")
  '
  '   ' Copies DocumentType Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(4) & "2:" & Clin_Doc_Col_Ltr_Array(4) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(4) & "2")
  '
  '   ' Copies Name Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(5) & "2:" & Clin_Doc_Col_Ltr_Array(5) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(5) & "2")
  '
  '   ' Copies Section Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(6) & "2:" & Clin_Doc_Col_Ltr_Array(6) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(6) & "2")
  '
  '   ' Copies DTA Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(7) & "2:" & Clin_Doc_Col_Ltr_Array(7) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(7) & "2")
  '
  '   ' Copies EventCode Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(8) & "2:" & Clin_Doc_Col_Ltr_Array(8) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(8) & "2")
  '
  '   ' Copies EventDisplay Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(9) & "2:" & Clin_Doc_Col_Ltr_Array(9) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(9) & "2")
  '
  '   ' Copies ESH Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(10) & "2:" & Clin_Doc_Col_Ltr_Array(10) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(10) & "2")
  '
  '   ' Copies ControlType Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(11) & "2:" & Clin_Doc_Col_Ltr_Array(11) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(11) & "2")
  '
  '   ' Copies NomenclatureID Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(12) & "2:" & Clin_Doc_Col_Ltr_Array(12) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(12) & "2")
  '
  '   ' Copies Nomenclature Display Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(13) & "2:" & Clin_Doc_Col_Ltr_Array(13) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(13) & "2")
  '
  '   ' Copies TaskAssay Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(14) & "2:" & Clin_Doc_Col_Ltr_Array(14) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(14) & "2")
  '
  '   ' Copies the Nomenclature Notes Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(15) & "2:" & Clin_Doc_Col_Ltr_Array(15) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(15) & "2")
  '
  '   ' Copies the Social History Notes Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(16) & "2:" & Clin_Doc_Col_Ltr_Array(16) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(16) & "2")
  '
  '   ' Copies the Grid Notes Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(17) & "2:" & Clin_Doc_Col_Ltr_Array(17) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(17) & "2")
  '
  '   ' Copies the Freetext Notes Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(18) & "2:" & Clin_Doc_Col_Ltr_Array(18) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(18) & "2")
  '
  '   ' Copies the Team Column
  '   Sheets(Val_Wk_Array(0)).Select
  '   Range(Clin_Doc_Col_Ltr_Array(19) & "2:" & Clin_Doc_Col_Ltr_Array(19) & LR).SpecialCells(xlCellTypeVisible).Copy Sheets(Code_Sheet).Range(CS_72_Header_Ltr_Array(19) & "2")
  '
  '
  '   '  Removes dups on the PTCare sheet by source, nomenclature ID, Nomenclature Display
  '   Sheets(Code_Sheet).Range("$A$1:$W$700").RemoveDuplicates Columns:=Array(CS_72_Header_Num_Array(3), _
  '   CS_72_Header_Num_Array(12), CS_72_Header_Num_Array(13)), _
  '   Header:=xlYes
  '
  ' End If


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '   PRIMARY - Workbook Cleanup
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Deletes the extra sheets not needed
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

  Application.DisplayAlerts = True

  '
  ' SUB - Creates and Populates Index Sheet
  '''''''''''''''''''''''''''''''''''''''''

  ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "Index Sheet"

  Range("A1").Select
  Selection.Value = "Index Sheet"
  ActiveCell.Offset(1, 0).Select

  For Each Sheet In Worksheets
    If Sheet.Name <> "Index Sheet" Then
      ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheet.Name & "'" & "!A1", TextToDisplay:=Sheet.Name
      ActiveCell.Offset(1, 0).Select
    End If
  Next Sheet

  '
  ' SUB - Formats remaining sheets as tables for appearance
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For Each Sheet In Worksheets
    Sheet.Activate

    Set sht = Sheet
    Set StartCell = Range("A1")

    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
    Sheet_Name = Sheet.Name

    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    ' If User wants to clear all the formatting then clear all, else clear boarders
    If ClearFormatting = True Then
      Selection.ClearFormats
    Else
      Selection.Borders.LineStyle = xlNone
    End If

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = Sheet_Name
    tbl.TableStyle = "TableStyleLight9"
    Columns.AutoFit
    Range("A1").Select

    ' Clears all data validation which could be copied
    Sheet.Cells.Validation.Delete

  Next Sheet


  '
  '  SUB - Aligns Index Sheet
  '''''''''''''''''''''''

  Sheets("INDEX SHEET").Select
  Range("A2").Select
  Range(Selection, Selection.End(xlDown)).Select
  With Selection
    .HorizontalAlignment = xlLeft
  End With
  Range("A2").Select

  '
  '  SUB - Saves the new workbook
  ''''''''''''''''''''''''''''''''''''''

  ' Saves new workbook then Switches back to the validation form to begin next loop
  Workbooks(Source_Name & ".xlsx").Close SaveChanges:=True
  Windows(Validation_File_Name).Activate

  ' Start over with next source from list
Next Source_Name



End_Program:

'Re-enables previously disabled settings after all code has run.
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

' Clears the filesystem descriptor allowing you to delete the folder
Dir "C:\"
ChDir "C:\"

' Activates primary excel file
Windows(Validation_File_Name).Activate

'Notifies user that the program has completed.
MsgBox ("Your PCST Files have been created. Folder is loctated within your My Documents.")

Exit Sub

User_Exit:

' If the active workbook is not the validation form then close it without saving
If ActiveWorkbook.Name = (Source_Name & ".xlsx") Then
  Workbooks(Source_Name & ".xlsx").Close SaveChanges:=False
End If

' Activates primary excel file
Windows(Validation_File_Name).Activate

'Re-enables previously disabled settings after all code has run.
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

' Clears the filesystem descriptor allowing you to delete the folder
Dir "C:\"
ChDir "C:\"

MsgBox ("Program quitting per user action.")

Exit Sub


ErrHandler:

If Source_Name <> vbNullString Then
  Workbooks(Source_Name & ".xlsx").Close SaveChanges:=False
End If

' Activates primary excel file
Windows(Validation_File_Name).Activate

' Clears the filesystem descriptor allowing you to delete the folder
Dir "C:\"
ChDir "C:\"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

MsgBox ("Exiting program because of an issue." & vbNewLine & vbNewLine & "Sad Panda :(" & vbNewLine & vbNewLine & vbNewLine & Err.Number & vbNewLine & Err.Description)

End Sub
