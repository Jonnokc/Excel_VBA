Sub SHX_Setup()
'
' SHX REsults Macro. Formats range as table. Inserts lookup formulas and populates autofill.
'
Dim tbl As ListObject
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range
Dim rList As Range
Dim Header_Array As Variant
Dim Validated_Map_Headers As Variant
Dim Result_Check As Boolean
Dim Header_Check As Boolean
Dim Start_Check As Integer

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Error Handling
    On Error GoTo ErrHandler

    Header_Array = Array("NOMEN_ID", "CS_72", "CS_14003", "CS_4002165")
    Validated_Map_Headers = Array("CODE ID", "MAPPING STATUS")

    'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Start_Check = MsgBox("BORIS is about to run the SHX helper program. This can take a few minutes." & vbNewLine & vbNewLine & "Click OK to start or Cancel to exit. Follow on screen prompts otherwise leave your computer alone until BORIS is done.", vbOKCancel + vbQuestion, "There Can Only Be One BORIS")

    'If user hits cancel then close program.
    If Start_Check = vbCancel Then
        GoTo User_Exit
    End If

    Sheets("Social History Results").Select

    ActiveSheet.AutoFilterMode = False        'Disables autoFilter

    'If table exists on sheet then convert to range
    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With

    End If

    Set sht = Worksheets("Social History Results")
    Set StartCell = Range("A1")

    'Refresh UsedRange
    Worksheets("Social History Results").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "SHX_Results"
    tbl.TableStyle = "TableStyleLight12"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With


    ' SUB - Finds the locations of the lookup columns on the social history results sheet
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Re-enables screen update incase user needs to interact
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Sheets("Social History Results").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Name = "Header_Row"

    For i = 0 To UBound(Header_Array)
        Header_Check = False
        For Each Header In Range("Header_Row")
            If LCase(Header_Array(i)) = LCase(Header) Then
                Header_Array(i) = Mid(Header.Address, 2, 1)
                Header_Check = True
                Exit For
            End If
        Next Header
        If Header_Check = False Then
            Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Header_Array(i) & "'" & " on the Social History Results Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

            'If user hits cancel then close program.
            If Header_User_Response = vbNullString Then
                GoTo User_Exit
            Else
                Header_Array(i) = Header_User_Response
            End If
        End If
    Next i


    ' SUB - Finds the headers for the validated mappings sheet
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Validated Mappings").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Name = "Val_Headers"

    For i = 0 To UBound(Validated_Map_Headers)
        Header_Check = False
        For Each Header In Range("Val_Headers")
            If LCase(Validated_Map_Headers(i)) = LCase(Header) Then
                Validated_Map_Headers(i) = Mid(Header.Address, 2, 1)
                Header_Check = True
                Exit For
            End If
        Next Header
        If Header_Check = False Then
            Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Validated_Map_Headers(i) & "'" & " on the Social History Results Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

            'If user hits cancel then close program.
            If Header_User_Response = vbNullString Then
                GoTo User_Exit
            Else
                Validated_Map_Headers(i) = Header_User_Response
            End If
        End If
    Next i

    ' Disables screen update again after user interaction part
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Sets Code ID range for looping to check format
    Range(Validated_Map_Headers(0) & "2:" & Validated_Map_Headers(0) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "CodeID_Column"

    For Each cell In Range("CodeID_Column")
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell


    ' Checks code columns and makes sure values are in number format for lookup
    For i = 0 To UBound(Header_Array)
        'Selects all cells not empty in the code column and assigns named range for loop
        Range(Header_Array(i) & "2:" & Header_Array(i) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Code_Column"

        For Each cell In Range("Code_Column")
            If IsNumeric(cell) Then
                cell.Value = Val(cell.Value)
                cell.NumberFormat = "0"
            End If
        Next cell
    Next i


    ' CS Nomenclature Mapped

    Sheets("Social History Results").Range("A2").Formula = _
            "=IFERROR(INDEX('Validated Mappings'!" & Validated_Map_Headers(1) & ":" & Validated_Map_Headers(1) & ",MATCH(" & Header_Array(0) & "2,'Validated Mappings'!" & Validated_Map_Headers(0) & ":" & Validated_Map_Headers(0) & ",0))," & Header_Array(0) & "2)"

    ' CS 72 formula
    Sheets("Social History Results").Range("B2").Formula = _
            "=IFERROR(INDEX('Validated Mappings'!" & Validated_Map_Headers(1) & ":" & Validated_Map_Headers(1) & ",MATCH(" & Header_Array(1) & "2,'Validated Mappings'!" & Validated_Map_Headers(0) & ":" & Validated_Map_Headers(0) & ",0))," & Header_Array(1) & "2)"

    ' CS 14003 Formula
    Sheets("Social History Results").Range("C2").Formula = _
            "=IFERROR(INDEX('Validated Mappings'!" & Validated_Map_Headers(1) & ":" & Validated_Map_Headers(1) & ",MATCH(" & Header_Array(2) & "2,'Validated Mappings'!" & Validated_Map_Headers(0) & ":" & Validated_Map_Headers(0) & ",0))," & Header_Array(2) & "2)"

    ' CS 4002165
    Sheets("Social History Results").Range("D2").Formula = _
            "=IFERROR(INDEX('Validated Mappings'!" & Validated_Map_Headers(1) & ":" & Validated_Map_Headers(1) & ",MATCH(" & Header_Array(3) & "2,'Validated Mappings'!" & Validated_Map_Headers(0) & ":" & Validated_Map_Headers(0) & ",0))," & Header_Array(3) & "2)"


    'Centers cell values
    Sheets("Social History Results").Select
    Columns("A:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    'Re-enables Auto-calculate to make sure formulas are accurate
    Application.Calculation = xlCalculationAutomatic

    ' SUB - If the Review sheet does not exist (ie, working with an old file, then create the sheet)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Sheet In Worksheets
        Result_Check = False
        If Sheet.Name = "To_Review" Then
            ' Delete all data on the sheet
            Sheets("To_Review").Select
            Cells.Select
            Selection.Clear
            Result_Check = True
            Exit For
        End If
    Next Sheet

    If Result_Check = False Then
        With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "To_Review"
        End With
    End If

    ' SUB - Copies the data to the To_Review Sheet
    ''''''''''''''''''''''''''''''''''''''''''''''''

    ' Disables calculation again
    Application.Calculation = xlCalculationManual

    Sheets("Social History Results").Cells.Copy Sheets("To_Review").Range("A1")



    ' SUB - Formats the data on the To Review Sheet
    '''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("To_Review").AutoFilterMode = False

    'If table exists on this sheet, then convert to range
    If Sheets("To_Review").ListObjects.Count > 0 Then

        With Sheets("To_Review").ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With
        ' Changes header font color to white
        Rows("1:1").Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

    End If

    Set sht = Worksheets("To_Review")
    Set StartCell = Range("A1")

    Worksheets("Social History Results").UsedRange

    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "SHX_To_Review"
    tbl.TableStyle = "TableStyleLight9"

    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.Select
    Cells.EntireRow.AutoFit

    Columns("A:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Range("A1").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "GJ Team - All Done!", vbOKOnly, "BORIS"

    Exit Sub

User_Exit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Program quitting per user action.", vbOKOnly, ":("
    Exit Sub

ErrHandler:
    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Something went wrong that you won't be able to fix. (code vs. form version issue most likely) Contact code creator for troubleshooting" & vbNewLine & vbNewLine & "Sad Panda :(" & vbNewLine & vbNewLine & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly, ":("

End Sub
