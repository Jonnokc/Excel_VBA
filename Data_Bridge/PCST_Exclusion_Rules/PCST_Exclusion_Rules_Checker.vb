Sub ExclusionTesting()


    Dim SheetName As Variant
    Dim RawCodeColumn As Variant
    Dim lastrow As Long
    Dim sht As Worksheet
    Dim LastColumn As Long
    Dim StartCell As Range


    Sheet1.Select
    SheetName = InputBox("What is the name of the sheet with all the Unmapped codes?")
    RawCodeColumn = InputBox("What is the column letter for the 'Raw Code Display' Column?")


    ' SUB - Names the range of the Exclusion Rules for Looping
    Set sht = Worksheets("ExclusionRules")

    With sht
      Set StartCell = .Range("A2")
      'Find Last Row and Column
      lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      'Select Range
      sht.Range(StartCell, sht.Cells(lastrow, "B")).Name = "Exclusion_Rules"
    End With


    ' Names the Raw Code Column for looping
    Set sht = Sheets(SheetName)

    With sht
        Set StartCell = .Range(RawCodeColumn & "2")
        lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        ' Names range for loop
        sht.Range(StartCell, sht.Cells(lastrow, RawCodeColumn)).Name = "Codes"
    End With

    ' Sets the Results column
    With sht
      ' Checks if header already exists
      Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For Each Header In Range("Header_row")
        If InStr(1, Header, "Exclusion Check Results") > 0 Then
          header_location = Mid(Header.Address, 2, 1)
          Columns(header_location & ":" & header_location).Select
          Selection.Delete Shift:=xlToLeft
        End If
      Next Header


      ' Names Exclusion Check Results Range
      NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
      Range(NextBlank & "1") = "Exclusion Check Results"

      sht.Range(NextBlank & "2", sht.Cells(lastrow, NextBlank)).Name = "Results"
    End With


    ExRules = Range("Exclusion_Rules").Value
    RawCodes = Range("Codes").Value
    ExclusionResults = Range("Results")

    For Rule = 1 To UBound(ExRules)
      For Code = 1 To UBound(RawCodes)
        CurrentRule = ExRules(Rule, 2)
        CurrentRuleNumber = ExRules(Rule, 1)
        CurrentCode = RawCodes(Code, 1)

        If InStr(1, CurrentCode, CurrentRule) > 0 Then
          ExclusionResults(Code, 1) = "Breaks Rule " & CurrentRuleNumber & " " & CurrentRule
          Exit For
        End If
      Next Code
    Next Rule

  Range("Results") = ExclusionResults



End Sub
