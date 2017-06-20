Sub Inappropriate_Data_Act_Setup()


  Dim tbl As ListObject
  Dim First_Sheet As String
  Dim sht As Worksheet
  Dim StartCell As Range
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim Headers As Variant
  Dim Headers_Num As Variant
  Dim New_Headers As Variant
  Dim New_Headers_Num As Variant



  Headers = Array("Population ID", "Source", "Entity", "Raw Code", "Standardized Code", "Count")
  Headers_Num = Array("Population ID", "Source", "Entity", "Raw Code", "Standardized Code", "Count")
  New_Headers = Array("Raw Code Display", "Raw Code ID", "Raw Code System ID", "Standard Code Display", "Standard Code ID", "OID")
  New_Headers_Num = Array("Raw Code Display", "Raw Code ID", "Raw Code System ID", "Standard Code Display", "Standard Code ID", "OID")

  ' Finds the name of the first worksheet
  For Each Sheet In Worksheets
    If Sheet.Visible Then
      First_Sheet = Sheet.Name
      Exit For
    End If
  Next Sheet

  Sheets(First_Sheet).Select

  ' SUB - Creates a To_Review Sheet and copies data over

  ' Deletes the extra sheets not needed
  Application.DisplayAlerts = False

  For Each Sheet In Worksheets
    If Sheet.Name = "To_Review" Then
      Sheet.Delete
    End If
  Next Sheet

  Application.DisplayAlerts = True

  ' Copies the sheet to the "To_Review"
  Set sht = Worksheets(First_Sheet)
  With sht
    Sheets(First_Sheet).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "To_Review"
  End With


  ' SUB - Finds Column Header Locations for The Combined Data Export
  Sheets("To_Review").Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"


  ' SUB - Finds the columns by their headers
  For i = 0 To UBound(Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Headers(i)) = LCase(Header) Then
        Headers(i) = Mid(Header.Address, 2, 1)
        Headers_Num(i) = Range(Headers(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Headers(i) & "'" & " on the " & FirstSheet & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

      'If user hits cancel then close program.
      If Header_User_Response = vbNullString Then

      Else
        Headers(i) = Header_User_Response
        Headers_Num(i) = Range(Headers(i) & "1").Column
      End If
    End If

  Next i


  ' SUB - Formats First_Sheet Sheet as a table
  Sheets(First_Sheet).Select

  ' table can not be created if autofilters are on
  If Sheets(First_Sheet).AutoFilterMode = True Then
    Sheets(First_Sheet).AutoFilterMode = False
  End If

  ' Checks the current sheet. If it is in table format, convert it to range.
  If Sheets(First_Sheet).ListObjects.Count > 0 Then
    With Sheets(First_Sheet).ListObjects(1)
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

  Set sht = Worksheets(First_Sheet)
  Set StartCell = Range("A1")

  ' Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Raw_Data_tbl"
  tbl.TableStyle = "TableStyleLight9"


  With Sheets(First_Sheet)
    ' SUB - Sorts the sheet by Count from largest to smallest
    ActiveWorkbook.Worksheets(First_Sheet).ListObjects("Raw_Data_tbl").Sort. _
    SortFields.Clear
    ActiveWorkbook.Worksheets(First_Sheet).ListObjects("Raw_Data_tbl").Sort. _
    SortFields.Add Key:=Range(Headers(5) & "1:" & Headers(5) & LastRow), SortOn:=xlSortOnValues, Order:= _
    xlDescending, DataOption:=xlSortTextAsNumbers

  End With

  With ActiveWorkbook.Worksheets(First_Sheet).ListObjects("Raw_Data_tbl").Sort
    .Apply
  End With

  ' SUB - Inserts New Columns and Splits Standard Code
  With Sheets("To_Review")
    Sheets("To_Review").Select
    Columns(Headers(4) & ":" & Headers(4)).Select
    Columns(Headers(4) & ":" & Headers(4)).EntireColumn.Offset(0, 1).Insert
    Columns(Headers(4) & ":" & Headers(4)).EntireColumn.Offset(0, 1).Insert

  End With


  ' Deliminates the Standardized Code Column
  Application.DisplayAlerts = False

  Sheets("To_Review").Range(Headers(4) & "2").Select

  Range(Selection, Selection.End(xlDown)).Select
  Selection.TextToColumns Destination:=Range(Headers(4) & "2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
  :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
  TrailingMinusNumbers:=True

  Application.DisplayAlerts = True

  ' SUB - Adds the new column Headers

  ' Renames the Standard Code Display column
  Sheets("To_Review").Range(Headers(4) & "1").Select
  ActiveCell.Formula = "Standard Code Display"

  ' Adds the Standard Code ID Column
  Sheets("To_Review").Range(Headers(4) & "1").Offset(0, 1).Select
  ActiveCell.Formula = "Standard Code ID"

  ' Adds the OID column
  Sheets("To_Review").Range(Headers(4) & "1").Offset(0, 2).Select
  ActiveCell.Formula = "OID"


  ' SUB - Inserts New Columns and Splits raw Code Column
  Columns(Headers(3) & ":" & Headers(3)).Select
  Columns(Headers(3) & ":" & Headers(3)).EntireColumn.Offset(0, 1).Insert
  Columns(Headers(3) & ":" & Headers(3)).EntireColumn.Offset(0, 1).Insert


  ' Deliminates the Raw Code Column
  Application.DisplayAlerts = False
  Sheets("To_Review").Range(Headers(3) & "2").Select

  Range(Selection, Selection.End(xlDown)).Select

  Selection.TextToColumns Destination:=Range(Headers(3) & "2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
  :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
  TrailingMinusNumbers:=True

  Application.DisplayAlerts = True


  ' SUB - Renames and stores location of Raw Code Columns

  ' Renames the Raw Code Headers
  Sheets("To_Review").Range(Headers(3) & "1").Select
  ActiveCell.Formula = "Raw Code Display"

  ' Renames Raw Code ID Column
  Sheets("To_Review").Range(Headers(3) & "1").Offset(0, 1).Select
  ActiveCell.Formula = "Raw Code ID"

  ' Stores the location of the Raw Code System ID Column
  Sheets("To_Review").Range(Headers(3) & "1").Offset(0, 2).Select
  ActiveCell.Formula = "Raw Code System ID"



  ' SUB - Finds Column Header Locations for The Combined Data Export
  Sheets("To_Review").Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"


  ' SUB - Finds the columns of the new columns
  For i = 0 To UBound(New_Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(New_Headers(i)) = LCase(Header) Then
        New_Headers(i) = Mid(Header.Address, 2, 1)
        New_Headers_Num(i) = Range(New_Headers(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & New_Headers(i) & "'" & " on the " & FirstSheet & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

      'If user hits cancel then close program.
      If Header_User_Response = vbNullString Then

      Else
        New_Headers(i) = Header_User_Response
        New_Headers_Num(i) = Range(New_Headers(i) & "1").Column
      End If
    End If

  Next i


  ' SUB - Formats "To_Review" Sheet as a table
  Sheets("To_Review").Select

  ' table can not be created if autofilters are on
  If Sheets("To_Review").AutoFilterMode = True Then
    Sheets("To_Review").AutoFilterMode = False
  End If

  ' Checks the current sheet. If it is in table format, convert it to range.
  If Sheets("To_Review").ListObjects.Count > 0 Then
    With Sheets("To_Review").ListObjects(1)
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

  Set sht = Worksheets("To_Review")
  Set StartCell = Range("A1")

  ' Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "To_Review_tbl"
  tbl.TableStyle = "TableStyleLight9"


  ' Removes Duplicates from the Code ID Lists
  Sheets("To_Review").Range("To_Review_tbl[#All]").RemoveDuplicates Columns:=Array(Headers_Num(0), Headers_Num(1), Headers_Num(2), New_Headers_Num(3), New_Headers_Num(4)), Header:=xlYes

  ' Hides Columns Not needed
  With Sheets("To_Review")
    Columns(Headers(0)).EntireColumn.Hidden = True
    Columns(Headers(1)).EntireColumn.Hidden = True
    Columns(Headers(3)).EntireColumn.Hidden = True
    Columns("D").EntireColumn.Hidden = True
    Columns("G").EntireColumn.Hidden = True
    Columns("H").EntireColumn.Hidden = True
    Columns("I").EntireColumn.Hidden = True
    Columns("K").EntireColumn.Hidden = True
    Columns("L").EntireColumn.Hidden = True
  End With

  ' Re-finds the last row - would be false rows because of rows previously deleted, but not removed because we have not saved.
  LastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row

  ' Applies final sort
  With ActiveWorkbook.Worksheets("To_Review").ListObjects("To_Review_tbl").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("C2:C" & LastRow), SortOn:=xlSortOnValues, Order:= _
    xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range("G2:G" & LastRow), SortOn:=xlSortOnValues, Order:= _
    xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range("F2:F" & LastRow), SortOn:=xlSortOnValues, Order:= _
    xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=Range("E2:E" & LastRow), SortOn:=xlSortOnValues, Order:= _
    xlAscending, DataOption:=xlSortNormal

    .Apply
  End With

End Sub
