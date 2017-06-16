Sub Unmapped_Codes_Sheet_Setup()

  Dim tbl As ListObject
  Dim sht As Worksheet
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim StartCell As Range
  Dim Unmapped_Headers As Variant
  Dim Unmapped_Headers_Num As Variant

  Application.ScreenUpdating = False


  Unmapped_Headers = Array("Registry", "Measure", "Concept", "Raw Code Display", "Raw Code ID", "Source Name", "Coding System ID", "Record Count (SUM)")
  Unmapped_Headers_Num = Array("Registry", "Measure", "Concept", "Raw Code Display", "Raw Code ID", "Source Name", "Coding System ID", "Record Count (SUM)")

  ' Deletes the extra sheets not needed
  Application.DisplayAlerts = False

  For Each sheet In Worksheets
    If sheet.Name = "To_Review" Then
      sheet.Delete
    End If
  Next sheet

  Application.DisplayAlerts = True

  ' Finds the name of the first sheet
  For Each sheet In Worksheets
    If sheet.Visible Then
        FirstSheet = sheet.Name
        Exit For
        End If
    Next sheet

  Sheets(FirstSheet).Select

  If Sheets(FirstSheet).AutoFilterMode = True Then
    Sheets(FirstSheet).AutoFilterMode = False
  End If

  ' Checks the current sheet. If it is in table format, convert it to range.
  If Sheets(FirstSheet).ListObjects.Count > 0 Then
    With Sheets(FirstSheet).ListObjects(1)
      Set rList = .Range
      .Unlist
    End With
  End If

  Set sht = WorkSheets(FirstSheet)
  Set StartCell = Range("A1")


  'Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  'Select Range
  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  'Turn selected Range Into Table
  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Unmapped_Table"
  tbl.TableStyle = "TableStyleLight9"

  'changes font color of header row to white
  Rows("1:1").Select
  With Selection.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
  End With

  Range("A2").Select

  ' SUB - Finds Column Header Locations for The Combined Data Export
  Sheets(FirstSheet).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(Unmapped_Headers)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Unmapped_Headers(i)) = LCase(Header) Then
        Unmapped_Headers(i) = Mid(Header.Address, 2, 1)
        Unmapped_Headers_Num(i) = Range(Unmapped_Headers(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header
    If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Unmapped_Headers(i) & "'" & " on the " & FirstSheet & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

      'If user hits cancel then close program.
      If Header_User_Response = vbNullString Then

      Else
        Unmapped_Headers(i) = Header_User_Response
        Unmapped_Headers_Num(i) = Range(Unmapped_Headers(i) & "1").Column
      End If
    End If

  Next i

  ' Copies the raw codes to the "To Review Sheet"
  Set sht = WorkSheets(FirstSheet)
  With sht
    Sheets(FirstSheet).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "To_Review"
  End With


  ' Deletes all the data from the Registry, Measure, Concept Columns
  Sheets("To_Review").Select
  Range("A2:C2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Clear

  ' Formats the data on the "To Review" Sheet into a TableStyle
  Set sht = Worksheets("To_Review")
  Set StartCell = Range("A1")


  Set tbl = ActiveSheet.ListObjects(1)
  tbl.Name = "Unmapped_To_Rev"
  tbl.TableStyle = "TableStyleLight12"
  Columns.AutoFit

  ' Clears all formatting
  Cells.Select
  Selection.ClearFormats
  Sheets("To_Review").ListObjects("Unmapped_To_Rev").TableStyle = "TableStyleLight12"

  ' Removes Duplicates By Raw Code Display, Raw Code ID, Source Name, Coding System ID
  Sheets("To_Review").Range("Unmapped_To_Rev").RemoveDuplicates Columns:=Array(Unmapped_Headers_Num(3), Unmapped_Headers_Num(4), Unmapped_Headers_Num(5), Unmapped_Headers_Num(6)), Header:=xlYes

  Dim Temp_Cell As Variant
  Dim Rng As Range

  ' Converts numbers stored as text to numbers
  Sheets("To_Review").Select

  Range(Unmapped_Headers(4) & "2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Name = "Codes"

  Set Rng = Range("Codes")

  For Each cell In Rng
      If IsNumeric(cell) Then
          cell.Value = Val(cell.Value)
          cell.NumberFormat = "0"
      End If
  Next cell

  MsgBox ("Program is completed. Ready for your analysis")


End Sub
