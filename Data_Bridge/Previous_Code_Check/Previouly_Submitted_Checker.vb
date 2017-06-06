Sub PrevCodeCheckBridge()

Dim sheet As Worksheet
Dim fso As Object
Dim SheetName As Variant
Dim FirstSheet As Variant
Dim RawCodeColumn As Variant
Dim lastrow As Long
Dim sht As Worksheet
Dim LastColumn As Long
Dim StartCell As Range
Dim Header_Check As Boolean
Dim Header_User_Response As Variant
Dim SheetArray As Variant
Dim UnmappedHeaderLocations As Variant
Dim HeaderNames As Variant
Dim EvCodeCheck As Variant
Dim EvCodeCheckAnswerArray As Variant
Dim EVCodeCheckArray As Variant
Dim EvCodeCheckHeader As Variant
Dim EvCodeConcat As Variant
Dim Client_Mnemonic As Variant
Dim Header As Variant
Dim LR As Long
Dim Lookup As Variant
Dim cell_Lookup As Variant
Dim sResult_Value As Variant

SheetArray = Array("Unmapped Raw", "CodeSystemCheck")
UnmappedHeaderLocations = Array("Coding System ID", "Raw Code ID", "EvCodeCheck", "CodeLookup")
HeaderNames = Array("Coding System ID", "Raw Code", "EvCodeCheck")
CodeSystemCheckHeaders = Array("UnmappedLookup")



    MsgBox ("PCST Previously Submitted Checker Is About to Run. Please follow on screen prompts if any. Otherwise leave computer alone until BORIS is done.")


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
    ' Checks to see if the CodeSystemCheck Sheet already exists.

    Application.DisplayAlerts = False

    For Each Sheet In Worksheets
      If Sheet.Name = "CodeSystemCheck" Then
        Sheet.Delete
      End If
    Next Sheet

    Application.DisplayAlerts = True

    With ThisWorkbook
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "CodeSystemCheck"
    End With


' PRIMARY - Imports the previously reviewed unmapped
''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With Sheets("CodeSystemCheck").ListObjects.Add(SourceType:=0, Source:=Array( _
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
        .CommandText = Array("QPCSTRequestedUnmappedCheck")
        .SourceDataFile = _
        "Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb"
        .ListObject.DisplayName = "Table_Data_Intelligence_Code_Database.accdb"
        .Refresh BackgroundQuery:=False
    End With


    ' Breaks the link so the database isn't locked in read only.
    Sheets("CodeSystemCheck").ListObjects("Table_Data_Intelligence_Code_Database.accdb").Unlink


        ' SUB - Finds the column for the code system ID on the unmapped codes Sheet
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ' SUB - Prompts user for current client mnemonic id
        Client_Mnemonic = InputBox("What is the full client mnemonic for this client?" & vbNewLine & vbNewLine & "ex. CERN_PH")

        If Client_Mnemonic = vbCancel Then
          Goto User_Exit
        End If

        ' Makes sure case is upper
        Client_Mnemonic = UCase(Client_Mnemonic)


        ' Selects the first sheet in the workbook and cell A1 to make data easily visible to the user.
        Sheet1.Select
        Range("A1").Select

        ' Adjusts the zoom so all the column headers can be seen.
        Cells.Select
        Selection.ColumnWidth = 14.86
        ActiveSheet.AutoFilterMode = False

        ' Checks with the user to confirm the sheet and the data they want to run the checker against.
        FirstSheet = Sheet1.Name
        SheetChecker = MsgBox("BORIS found the sheet '" & FirstSheet & "' Is this the sheet with the data you want to review?", vbYesNo)

        If SheetChecker = vbYes Then
          SheetName = FirstSheet
        Else
          SheetName = InputBox("Please enter the name of the sheet containing the data you want to review")
          If SheetName = vbNullString Then
              GoTo User_Exit
          End If
        End If

        ' Stores the Name of the data sheet in the Array
        SheetArray(0) = SheetName


        ' SUB - Checks if there already is a column titled EVCodeCheck, and code Check. If Not then Create them.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Sheets(SheetArray(0)).Select
        Range("A1:A2").Select
        Range(Selection, Selection.End(xlToRight)).Name = "Header_row"


        ' Checks if header EvCodeCheck already exists. If it finds a hit, delete that column and start again.
    BeginAgain1:
        For Each Header In Range("Header_row")
          If InStr(1, Header, "EvCodeCheck") Then
            Header.EntireColumn.Delete
            GoTo BeginAgain1
          End If
        Next Header

        ' Checks if header CodeLookup already exists. If it finds a hit, delete that column and start again.
    BeginAgain2:
        For Each Header In Range("Header_row")
          If InStr(1, Header, "CodeLookup") Then
            Header.EntireColumn.Delete
            GoTo BeginAgain2
          End If
        Next Header

        ' Names The Column CodeLookup Column
        NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
        Range(NextBlank & "1") = "CodeLookup"

        ' Names The Column EvCodeCheck Column
        NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
        Range(NextBlank & "1") = "EvCodeCheck"


        ' SUB - FINDS THE COLUMN LOCATIONS OF THE RAW DATA Sheet
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Confirms header range.
        Sheets(SheetArray(0)).Select
        Range("A1:A2").Select
        Range(Selection, Selection.End(xlToRight)).Name = "Header_row"

        For i = 0 To UBound(UnmappedHeaderLocations)

            ' Finds columns by header name
            Header_Check = False
            For Each Header In Range("Header_row")
                If LCase(UnmappedHeaderLocations(i)) = LCase(Header) Then
                    UnmappedHeaderLocations(i) = Mid(Header.Address, 2, 1)
                    Header_Check = True
                    Exit For
                End If
            Next Header

            ' If no header was found then prompt the user for the column or allow the user to cancel the program
            If Header_Check = False Then
                Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & HeaderNames(i) & " on the " & SheetArray(0) & " Sheet." & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
                If Header_User_Response = vbNullString Then
                    GoTo User_Exit
                Else
                    UnmappedHeaderLocations(i) = UCase(Header_User_Response)
                End If
            End If
        Next i

        ' SUB - Finds the Unmapped Lookup Header on the CodeSystemCheck SheetArray

        Sheets(SheetArray(1)).Select
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Name = "Header_row"

        For i = 0 To UBound(CodeSystemCheckHeaders)

            ' Finds columns by header name
            Header_Check = False
            For Each Header In Range("Header_row")
                If LCase(CodeSystemCheckHeaders(i)) = LCase(Header) Then
                    CodeSystemCheckHeaders(i) = Mid(Header.Address, 2, 1)
                    Header_Check = True
                    Exit For
                End If
            Next Header

            ' If no header was found then prompt the user for the column or allow the user to cancel the program
            If Header_Check = False Then
                Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & CodeSystemCheckHeaders(i) & " on the " & SheetArray(1) & " Sheet." & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
                If Header_User_Response = vbNullString Then
                    GoTo User_Exit
                Else
                    CodeSystemCheckHeaders(i) = UCase(Header_User_Response)
                End If
            End If
        Next i



        ' SUB - Creates Concat Column
        Sheets(SheetArray(0)).Select
        LR = Range(UnmappedHeaderLocations(0) & Rows.Count).End(xlUp).Row
        Range(UnmappedHeaderLocations(3) & "2:" & UnmappedHeaderLocations(3) & LR).Formula = "=CONCATENATE(" & CHR(34) & Client_Mnemonic & CHR(34) & "," & CHR(34) & "|" & CHR(34) & "," & UnmappedHeaderLocations(0) & "2" & "," & CHR(34) & "|" & CHR(34) & "," & UnmappedHeaderLocations(1) & "2" & ")"


        ' SUB - Assigns CodeLookup Column to an array in memory
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Sheets(SheetArray(0)).Select
        LR = Range(UnmappedHeaderLocations(0) & Rows.Count).End(xlUp).Row
        Range(UnmappedHeaderLocations(3) & "2:" & UnmappedHeaderLocations(3) & LR).SpecialCells(xlCellTypeVisible).Name = "CodeLookup"

        EVCodeCheckArray = Range("CodeLookup").Value

        ' SUB - Set EvCodeCheck answer range to array in memory
        Sheets(SheetArray(0)).Select
        Range(UnmappedHeaderLocations(2) & "2:" & UnmappedHeaderLocations(2) & LR).SpecialCells(xlCellTypeVisible).Name = "EvCodeCheck"

        EvCodeCheckAnswerArray = Range("EvCodeCheck")


        ' SUB - Assigns a name to the Previous ev code lookup column
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Sheets(SheetArray(1)).Select
        LR = Range("A" & Rows.Count).End(xlUp).Row
        Range(CodeSystemCheckHeaders(0) & "1:" & CodeSystemCheckHeaders(0) & LR).SpecialCells(xlCellTypeVisible).Name = "PreviousEvCodes"


        ' SUB - checks each cell in the EvCodeCheck for matches and either assigns a match or returns 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To UBound(EVCodeCheckArray)
            cell_Lookup = EVCodeCheckArray(i, 1)
            sResult = Application.VLookup(cell_Lookup, Range("PreviousEvCodes"), 1, False)
            If IsError(sResult) Then
                sResult_Value = ""
                EvCodeCheckAnswerArray(i, 1) = sResult_Value
            Else
                EvCodeCheckAnswerArray(i, 1) = "Previously Submitted"
            End If
        Next i

        ' Write the updated DataRange Array to the excel file
        Range("EvCodeCheck").Value = EvCodeCheckAnswerArray


        ' Deletes the extra sheet
        Application.DisplayAlerts = False

        For Each sheet In Worksheets
            If sheet.Name = "CodeSystemCheck" Then
              sheet.Delete
            End If
        Next sheet

        ' Deletes the Code Lookup column

        Sheets(SheetArray(0)).Select
        Range("A1:A2").Select
        Range(Selection, Selection.End(xlToRight)).Name = "Header_row"


        ' Checks if header CodeLookup already exists. If it finds a hit, delete that column and start again.

        For Each Header In Range("Header_row")
          If InStr(1, Header, "CodeLookup") Then
            Header.EntireColumn.Delete
          End If
        Next Header


        Application.DisplayAlerts = True

        ' Tells user program is completed
        Sheets(SheetArray(0)).Select
        Range("A1").Select
        MsgBox ("BORIS has completed the Previously Submitted Code Check. Check the newly added column. Blank = Not previously submitted.")
        Exit Sub

    User_Exit:
        MsgBox ("Exiting per user action")

    End Sub
