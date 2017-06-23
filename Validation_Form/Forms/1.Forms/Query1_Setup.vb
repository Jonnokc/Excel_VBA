Sub Query1_Range_Helper()
'
'Formats the data as a table for filtering. Then marks each row at each 100 then finally filters the results by the color.
'

    Dim cell As Range
    Dim tbl As ListObject

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    MsgBox ("Program is about to run. Please leave computer alone until completed")

    'Formats range as table

    Sheets("Query1").Select
    Worksheets("Query1").AutoFilterMode = False
    Range("I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Forms_Query1"
    tbl.TableStyle = "TableStyleLight12"


    'Sorts the ID column from smallest to largest



    Range("I1").Select

    ActiveWorkbook.Worksheets("Query1").ListObjects("Forms_Query1").Sort.SortFields _
            .Clear
    ActiveWorkbook.Worksheets("Query1").ListObjects("Forms_Query1").Sort.SortFields _
            .Add Key:=Range("Forms_Query1[[#All],[DCP_FORMS_REF_ID]]"), SortOn:= _
                 xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Query1").ListObjects("Forms_Query1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'populates the index column

    Range("K2").Select
    Selection.Value = 1
    Range("K3").Select
    Selection.Value = 2
    Range("K4").Select
    Selection.Value = 3
    Range("K5").Select
    Selection.Value = 4
    Range("K6").Select
    Selection.Value = 5
    Range("K7").Select
    Selection.Value = 6

    Range("K2:K6").Select
    Selection.AutoFill Destination:=Range("Forms_Query1[Index]")
    Range("Forms_Query1[Index]").Select

    'marks the rows that are at 100's
    For Each cell In Range("K:K")
        If cell.Value = 100 _
           Or cell.Value = 200 _
           Or cell.Value = 300 _
           Or cell.Value = 400 _
           Or cell.Value = 500 _
           Or cell.Value = 600 _
           Or cell.Value = 700 _
           Or cell.Value = 800 _
           Or cell.Value = 900 _
           Or cell.Value = 1000 _
           Or cell.Value = 1200 _
           Or cell.Value = 1300 _
           Or cell.Value = 1400 _
           Or cell.Value = 1500 _
           Or cell.Value = 1600 _
           Or cell.Value = 1700 _
           Then
            cell.Interior.Color = XlRgbColor.rgbLightGreen
        End If
    Next cell

    'Filters the table by color
    ActiveSheet.ListObjects("Forms_Query1").Range.AutoFilter Field:=3, Criteria1:=RGB _
                                                                                  (144, 238, 144), Operator:=xlFilterCellColor

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox ("Program Completed")


End Sub
