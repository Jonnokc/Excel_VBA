Private Sub Nomenclature_Notes()

    '
    'This code will take values from a table and put them in an arrao.
    'Then it Will perform changes to the data within the array and then write the array back to the sheet.
    'This changes the values all at once instead of one at a time.
    '
    '

    Dim DataRange As Variant
    Dim Irow As Long
    Dim Icol As Integer
    Dim DocType As Variant
    Dim ControlArray As Variant
    Dim ControlTypeCheck As Variant
    Dim Nomenclature_Val_Check As Variant
    Dim EventCode_Val_Check As Variant
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim Sheet As Worksheet
    Dim rList As Range

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("New Lines").Select

    'Removes filters from sheet
    ActiveSheet.AutoFilterMode = False

    'convert the table back to a range
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


    Set sht = ActiveSheet
    Set StartCell = Range("A1")

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Creates the table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "New_Lines"    'Names the table
    tbl.TableStyle = "TableStyleLight12"    'Sets table color theme

    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    'Creates named Range starting at column E
    Sheets("New Lines").Range("E2:V2").Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Name = "Data_Range"

    'Array to check DocumentType
    ControlArray = Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi")
    UnmappedArray = Array("New Numeric", "Numeric", "Calculation", "Date Time")

    'Saves range to array
    DataRange = Range("Data_Range").Value    'writes the named data range to the array variable

    'Loops through all rows within the range.
    For Irow = 1 To UBound(DataRange)
        DocType = DataRange(Irow, 1)
        ControlTypeCheck = DataRange(Irow, 8)
        Nomenclature_Val_Check = DataRange(Irow, 18)
        EventCode_Val_Check = DataRange(Irow, 17)

        'Checks if control type is within the array.
        IsInControlArray = Not IsError(Application.Match(ControlTypeCheck, ControlArray, 0))
        IsInUnmappedArray = Not IsError(Application.Match(ControlTypeCheck, UnmappedArray, 0))

        If IsInControlArray = True _
                And Nomenclature_Val_Check = "0" _
                And EventCode_Val_Check = "0" _
                Then

            DataRange(Irow, 12) = "This nomenclature and event code are not mapped and should be if this will be used to complete the measure."
            DataRange(Irow, 16) = "PCST"


        ElseIf IsInControlArray = True _
                And Nomenclature_Val_Check = "Validated" _
                And EventCode_Val_Check = "0" _
                Then

            DataRange(Irow, 12) = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
            DataRange(Irow, 16) = "PCST"

        ElseIf IsInControlArray = True _
                And Nomenclature_Val_Check = "0" _
                And EventCode_Val_Check = "Validated" _
                Then

            DataRange(Irow, 12) = "This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure."
            DataRange(Irow, 16) = "Consulting"

        End If

        'If DocumentType is IView, then ignore the control type
        If LCase(DocType) = "iview" _
                And Nomenclature_Val_Check = "Validated" _
                And EventCode_Val_Check = "0" _
                Then

            DataRange(Irow, 12) = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
            DataRange(Irow, 16) = "PCST"

        ElseIf LCase(DocType) = "iview" _
                And Nomenclature_Val_Check = "0" _
                And EventCode_Val_Check = "Validated" _
                Then

            DataRange(Irow, 12) = "This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure."
            DataRange(Irow, 16) = "Consulting"

        ElseIf LCase(DocType) = "iview" _
                And Nomenclature_Val_Check = "0" _
                And EventCode_Val_Check = "0" _
                Then

            DataRange(Irow, 12) = "This nomenclature and event code are not mapped and should be if this will be used to complete the measure."
            DataRange(Irow, 16) = "PCST"

        End If

        'Unmapped Code comment
        If IsInUnmappedArray = True Then

            DataRange(Irow, 12) = "Unmapped code value that seems to be relevant to what we would want to measure in Registries."
            DataRange(Irow, 16) = "Consulting"

        End If

    Next Irow


    'Write the updated DataRange Array to the excel file
    Range("Data_Range").Value = DataRange

    're-enables settings previously disabled
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
