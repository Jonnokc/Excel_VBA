Sub EV_Code_Setup()

  Dim tbl As ListObject
  Dim sht As Worksheet
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim StartCell As Range
  Dim PvtTbl As PivotTable
  Dim EvCodeHeaderArray As Variant
  Dim ValidatedHeaderArray As Variant
  Dim missing_header As String


  EvCodeHeaderArray = Array("CODE_STATUS", "EVENT_CD")
  ValidatedHeaderArray = Array("CODE ID", "MAPPING STATUS")


  ' Disables settings to improve performance except calculation which is needed
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = False


  ' SUB - Finds column locations for the Event Codes Results Sheet Columns
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Sets header row range
  Set sht = Worksheets("Event Codes Results")
  With sht
    Set StartCell = .Range("A1")

    'Find Last Column
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
    ' Names range for loop
    sht.Range(StartCell, sht.Cells("1", LastColumn)).Name = "Header_row"


    For i = 0 To UBound(EvCodeHeaderArray)
      Header_Check = False
      For Each Header In Range("Header_row")
        If LCase(Header) = LCase(EvCodeHeaderArray(i)) Then
          EvCodeHeaderArray(i) = Mid(Header.Address, 2, 1)
          Header_Check = True
          Exit For
        End If
      Next Header

      ' If no header was found then prompt the user for the column or allow the user to cancel the program
      If Header_Check = False Then
        Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & EvCodeHeaderArray(i) & " on the Event Codes Results Sheet" & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
        If Header_User_Response = vbNullString Then
          GoTo User_Exit
        Else
          ' Renames the given column header with the missing one to be used down range for filtering etc. etc.
          missing_header = EvCodeHeaderArray(i)
          EvCodeHeaderArray(i) = UCase(Header_User_Response)
          .Range(EvCodeHeaderArray(i) & "1") = missing_header
        End If
      End If
    Next i
  End With


  ' SUB - Finds column locations for the Validated Code Sheet Column(s)
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Set sht = Worksheets("Validated Codes")
  With sht
    Set StartCell = .Range("A1")

    'Find Last Column
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
    ' Names range for loop
    sht.Range(StartCell, sht.Cells("1", LastColumn)).Name = "Header_row"

    For i = 0 To UBound(ValidatedHeaderArray)
      Header_Check = False
      For Each Header In Range("Header_row")
        If LCase(Header) = LCase(ValidatedHeaderArray(i)) Then
          ValidatedHeaderArray(i) = Mid(Header.Address, 2, 1)
          Header_Check = True
          Exit For
        End If
      Next Header

      ' If no header was found then prompt the user for the column or allow the user to cancel the program
      If Header_Check = False Then
        Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & ValidatedHeaderArray(i) & " on the Validated Codes Sheet" & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
        If Header_User_Response = vbNullString Then
          GoTo User_Exit
        Else
          ' Renames the given column header with the missing one to be used down range for filtering etc. etc.
          missing_header = ValidatedHeaderArray(i)
          ValidatedHeaderArray(i) = UCase(Header_User_Response)
          .Range(ValidatedHeaderArray(i) & "1") = missing_header
        End If
      End If
    Next i
  End With


  ' SUB - Formats Event Codes Results Sheet As Table
  ''''''''''''''''''''''''''''''''''

  Set sht = Worksheets("Event Codes Results")
  With sht
    If .ListObjects.Count > 0 Then
      With .ListObjects(1)
        Set rList = .Range
        .Unlist    ' convert the table back to a range
      End With
      With rList
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlLineStyleNone
      End With
    End If
    Set StartCell = .Range("A1")

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    Set tbl = .ListObjects.Add(xlSrcRange, .Range(StartCell, .Cells(LastRow, LastColumn)), , xlYes)
    tbl.Name = "EV_Results_Table"
    tbl.TableStyle = "TableStyleLight9"

    ' changes font color of header row to white
    With .Rows("1:1").Font
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = 0
    End With

    .Range("A2") = "=IFERROR(INDEX('Validated Codes'!" & ValidatedHeaderArray(1) & ":" & ValidatedHeaderArray(1) & ",MATCH(" & EvCodeHeaderArray(1) & "2,'Validated Codes'!" & ValidatedHeaderArray(0) & ":" & ValidatedHeaderArray(0) & ",0)),0)"
    .Range("A2").AutoFill Destination:=Range("EV_Results_Table[Mapped?]")
  End With


  ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
  "EV_Results_Table", Version:=6).CreatePivotTable TableDestination:= _
  "'Pivot Table'!R12C2", TableName:="EV_Pivot", DefaultVersion:=6

  ' set the Pivot Table "EV_Pilot" to an Object
  Set PvtTbl = Sheets("Pivot Table").PivotTables("EV_Pivot")
  With PvtTbl
    With .PivotFields("Mapped?")
      .Orientation = xlRowField
      .Position = 1
    End With
    With .PivotFields("CODE_STATUS")
      .Orientation = xlColumnField
      .Position = 1
    End With
    ' add field to Pivot Table as Count
    .AddDataField .PivotFields("EVENT_CD"), "Count of EVENT_CD", xlCount
  End With

  ' Selects the Validated count and changes color to red
  With Sheets("Pivot Table").Range("C15").Font
    .Color = -16776961
    .TintAndShade = 0
  End With

  With tbl
    ' Filters the results table column A for just "Validated"
    .Range.AutoFilter Field:=1, Criteria1:="Validated"
    ' Filters the Code_Status column for just "Active"
    .Range.AutoFilter Field:=3, Criteria1:="Active"
  End With

  're-enables settings previously disabled
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub

  ' Exits program and all called macros per user action
  User_Exit:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  MsgBox("Exiting Per User Action")
  End

End Sub
