Sub PCST_Exclusion_Run_Checker()

  Dim fso As Object
  Dim SheetName As Variant
  Dim FirstSheet As Variant
  Dim RawCodeColumn As Variant
  Dim lastrow As Long
  Dim sht As Worksheet
  Dim LastColumn As Long
  Dim StartCell As Range

  MsgBox ("PCST Exclusion Rules Checker Is About to Run. Please follow on screen prompts if any. Otherwise leave computer alone until BORIS is done.")


  ' PRIMARY - Imports the Exclusion Rules
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  ' SUB - Checks to confirm path to Access Database is mapped correctly.
  Set fso = CreateObject("scripting.filesystemobject")
  With fso
    Path_Checker = Len(Dir("Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb")) <> 0
  End With
  If Path_Checker = 0 Then
    MsgBox ("You have not mapped the shared network drive correctly for this program to run. Please check the wiki for instructions on how to map the network drive.")
  Else
    ' Do Nothing
  End If
  Set fso = Nothing



  ' Checks to see if the sheet already exists and if it does, deletes it.
  Application.DisplayAlerts = False

  For Each Sheet In Worksheets
    If Sheet.Name = "PCST_Exclusion_Rules" Then
      Sheet.Delete
    End If
  Next Sheet

  Application.DisplayAlerts = True

  With ThisWorkbook
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "PCST_Exclusion_Rules"
  End With

  With Sheets("PCST_Exclusion_Rules").ListObjects.Add(SourceType:=0, Source:=Array( _
  "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Y:\Data Intelligence\Code_Database\Data_Intelligence_Cod" _
  , _
  "e_Database.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:" _
  , _
  "Database Password=BORIS;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:" _
  , _
  "Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=" _
  , _
  "False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:" _
  , _
  "Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass Choice" _
  , "Field Validation=False"), Destination:=Range("$A$1")).QueryTable
    .CommandType = xlCmdTable
    .CommandText = Array("PCSTExclusionRules")
    .SourceDataFile = _
    "Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb"
    .ListObject.DisplayName = "Table_Data_Intelligence_Code_Database.accdb"
    .Refresh BackgroundQuery:=False
  End With

  ' Breaks the link so the database isn't locked in read only.
  Sheets("PCST_Exclusion_Rules").ListObjects("Table_Data_Intelligence_Code_Database.accdb").Unlink


  ' PRIMARY - Runs the rule Checker
  ''''''''''''''''''''''''''''''''''''

  ' Selects the first sheet in the workbook and cell A1 to make data easily visible to the user.
  Sheet1.Select
  Range("A1").Select

  ' Checks with the user to confirm the sheet and the data they want to run the checker against.
  FirstSheet = Sheet1.Name
  SheetChecker = MsgBox("BORIS found the sheet '" & FirstSheet & "' Is this the sheet with the data you want to review?", vbYesNo)

  If SheetChecker = vbYes Then
    SheetName = FirstSheet
  Else
    SheetName = InputBox("Please enter the name of the sheet containing the data you want to review")
  End If

  Sheets(SheetName).Select

  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For Each Header In Range("Header_row")
    Header_Finder = False
    If Header.Value = "Raw Code Display" Then
      Header_Location = Mid(Header.Address, 2, 1)
      Exit For
    End If
  Next Header

  RawCodeColumn = MsgBox("BORIS Found the header 'Raw Code Display' Is this the column you want to check?", vbYesNo)

  If RawCodeColumn = vbYes Then
    RawCodeColumn = Header_Location
  Else
    RawCodeColumn = InputBox("What is the column letter of the column you want to check?")
    RawCodeColumn = UCase(RawCodeColumn)
  End If


  ' SUB - Names the range of the Exclusion Rules for Looping
  Set sht = Worksheets("PCST_Exclusion_Rules")

  With sht
    Set StartCell = .Range("A2")
    'Find Last Row and Column
    lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    'Select Range
    sht.Range(StartCell, sht.Cells(lastrow, "B")).Name = "Exclusion_Rules"
  End With

  ' Names the Raw Code Column for looping
  Set sht = Sheets(SheetName)

  ' Names range for loop
  With sht
    Set StartCell = .Range(RawCodeColumn & "2")
    lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    sht.Range(StartCell, sht.Cells(lastrow, RawCodeColumn)).Name = "Codes"
  End With

  ' Sets the Results column
  With sht
    ' Checks if header already exists. If it finds a hit, delete that column and start again.
BeginAgain:
    Range("A1").Select
    Range("B2", Selection.End(xlToRight)).Name = "Header_row"

    For Each Header In Range("Header_row")
      If InStr(1, Header, "Exclusion Check Results") Then
        Header.EntireColumn.Delete
        GoTo BeginAgain
      End If
    Next Header

    ' Names Exclusion Check Results Range
    NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
    Range(NextBlank & "1") = "Exclusion Check Results"

    ' Names the results range.
    sht.Range(NextBlank & "2", sht.Cells(lastrow, NextBlank)).Name = "Results"
  End With

  ' Saves range to memory
  ExRules = Range("Exclusion_Rules").Value
  RawCodes = Range("Codes").Value
  ExclusionResults = Range("Results")

  ' Loops through cells for each row to find if it hits any rules
  For Rule = 1 To UBound(ExRules)
    For Code = 1 To UBound(RawCodes)
      CurrentRule = ExRules(Rule, 2)
      CurrentRuleNumber = ExRules(Rule, 1)
      CurrentCode = RawCodes(Code, 1)

      If InStr(1, CurrentCode, CurrentRule) > 0 Then
        ExclusionResults(Code, 1) = "Breaks Rule " & CurrentRuleNumber & " " & CurrentRule
      End If
    Next Code
  Next Rule

  ' Writes the rules back to the excel range.
  Range("Results") = ExclusionResults

  Sheets(SheetName).Select


  ' SUB - Delete the PCST Exclusion Rules Sheet. It is no longer needed.
  Application.DisplayAlerts = False
  Sheets("PCST_Exclusion_Rules").Delete
  Application.DisplayAlerts = True


  MsgBox ("BORIS is done! Check the new column 'Exclusion Check Results'. Rows with Violations were marked. Blank means no violation was found.")

End Sub
