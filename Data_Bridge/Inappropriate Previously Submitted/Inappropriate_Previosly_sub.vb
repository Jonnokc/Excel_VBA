Sub Inappropriate_Previously_Sub()



  Dim First_Sheet As String
  Dim Data_Headers As Variant
  Dim Previously_Sub_Headers As Variant
  Dim LastRow As Variant
  Dim NextBlankCol As Variant
  Dim DataCheckArray As Variant
  Dim DataCheckAnswerArray As Variant
  Dim PreviouslySubArray As Variant
  Dim Lookup As Variant
  Dim cell_Lookup As Variant
  Dim sResult_Value As Variant
  Dim sht As Worksheet

  ' This disables settings to improve macro performance.
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False


  Data_Headers = Array("Standardized Code", "Prev_Sub_Check")
  Previously_Sub_Headers = Array("Standard Code Display")

  ' PRIMARY - Import Previously Submitted
  '''''''''''''''''''''''''''''''''''''''''''''

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

  ' Deletes the previously submitted sheet if if already exists. This is for re-runs
  Application.DisplayAlerts = False

  For Each Sheet In Worksheets
    If Sheet.Name = "PreviouslySubChck" Then
      Sheet.Delete
    End If
  Next Sheet

  Application.DisplayAlerts = True

  With ThisWorkbook
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "PreviouslySubChck"
  End With

  ' SUB - Imports Previously Submitted Data
  With Sheets("PreviouslySubChck").ListObjects.Add(SourceType:=0, Source:=Array( _
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
    .CommandText = Array("Q_Inappropriate_Data_Activity_Submitted_Codes")
    .SourceDataFile = _
    "Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb"
    .ListObject.DisplayName = "Table_Data_Intelligence_Code_Database.accdb5"
    .Refresh BackgroundQuery:=False
  End With


  ' Breaks the link so the database isn't locked in read only.
  Sheets("PreviouslySubChck").ListObjects("Table_Data_Intelligence_Code_Database.accdb5").Unlink

  ' Finds the name of the first worksheet
  For Each Sheet In Worksheets
    If Sheet.Visible Then
      First_Sheet = Sheet.Name
      Exit For
    End If
  Next Sheet

  ' Creates named range for header loop
  Sheets(First_Sheet).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  ' Checks if header Prev_Sub_Check already exists. If it finds a hit, delete that column and start again.
BeginAgain1:
  For Each Header In Range("Header_row")
    If InStr(1, Header, "Prev_Sub_Check") Then
      Header.EntireColumn.Delete
      GoTo BeginAgain1
    End If
  Next Header


  Sheets(First_Sheet).Select
  ' Names The Column Matching Prev_Sub_Check Column
  NextBlankCol = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
  Range(NextBlankCol & "1") = "Prev_Sub_Check"


  ' Creates named range for header loop
  Sheets(First_Sheet).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  ' SUB - Finds columns by header name
  For i = 0 To UBound(Data_Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Data_Headers(i)) = LCase(Header) Then
        Data_Headers(i) = Mid(Header.Address, 2, 1)
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & Data_Headers(i) & " on the " & FirstSheet & " Sheet." & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        Data_Headers(i) = UCase(Header_User_Response)
      End If
    End If
  Next i


  ' SUB Finds the location of the Headers for the Previously Submitted Sheet.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Creates named range for header loop
  Sheets("PreviouslySubChck").Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  ' SUB - Finds columns by header name
  For i = 0 To UBound(Previously_Sub_Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Previously_Sub_Headers(i)) = LCase(Header) Then
        Previously_Sub_Headers(i) = Mid(Header.Address, 2, 1)
        Header_Check = True
        Exit For
      End If
    Next Header

    ' If no header was found then prompt the user for the column or allow the user to cancel the program
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & Previously_Sub_Headers(i) & " on the " & FirstSheet & " Sheet." & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        Previously_Sub_Headers(i) = UCase(Header_User_Response)
      End If
    End If
  Next i



  ' SUB - Stores data into memory
  ''''''''''''''''''''''''''''''''''

  Sheets(First_Sheet).Select
  Set sht = Sheets(First_Sheet)

  With Sheets(First_Sheet)
    LastRow = sht.Cells(sht.Rows.Count, Data_Headers(0)).End(xlUp).Row

    ' stores the standard code display column to memory
    Range(Data_Headers(0) & "2:" & Data_Headers(0) & LastRow).SpecialCells(xlCellTypeVisible).Name = "CodeLookup"
    DataCheckArray = Range("CodeLookup").Value

    ' stores the Prev_Sub_Check column to memory
    Range(Data_Headers(1) & "2:" & Data_Headers(1) & LastRow).SpecialCells(xlCellTypeVisible).Name = "Prev_Sub_Check"
    DataCheckAnswerArray = Range("Prev_Sub_Check")

  End With

  Sheets("PreviouslySubChck").Select
  Set sht = Sheets("PreviouslySubChck")

  With Sheets("PreviouslySubChck")
    LastRow = sht.Cells(sht.Rows.Count, Previously_Sub_Headers(0)).End(xlUp).Row

    ' Stores previously submitted Standard Code Display column to memory
    Range(Previously_Sub_Headers(0) & "2:" & Previously_Sub_Headers(0) & LastRow).SpecialCells(xlCellTypeVisible).Name = "Previously_Sub"
    PreviouslySubArray = Range("Previously_Sub").Value

  End With

  ' SUB - Loops through memory arrays for previously submitted codes

  For i = 1 To UBound(DataCheckArray)
    cell_Lookup = DataCheckArray(i, 1)
    sResult = Application.VLookup(cell_Lookup, Range("Previously_Sub"), 1, False)
    If IsError(sResult) Then
      sResult_Value = ""
      DataCheckAnswerArray(i, 1) = sResult_Value
    Else
      DataCheckAnswerArray(i, 1) = "Previously Submitted"
    End If
  Next i

  ' Write the updated DataRange Array to the excel file
  Sheets(First_Sheet).Select
  Range("Prev_Sub_Check").Value = DataCheckAnswerArray

  ' Close
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  MsgBox ("Program completed Check new column")


  Exit Sub

User_Exit:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  MsgBox ("Quitting Per User Action")

End Sub
