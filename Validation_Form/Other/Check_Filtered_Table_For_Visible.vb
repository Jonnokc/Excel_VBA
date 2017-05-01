'
'First code checks a filtered table to determine if there is any content visible. This code was not working for some unknown reason when run on the Health Maintenance
'Summary sheet.
'
'
Sub TestEmptyTable()

    Dim Table_Obj As ListObject
    Dim outputPasteRange As Range
    Dim Table_ObjIsVisible As Boolean

    Set Table_Obj = ActiveSheet.ListObjects(1)
    Set outputPasteRange = Range("B15")


    'Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
    If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        Table_ObjIsVisible = True
    Else:
        Table_ObjIsVisible = Table_Obj.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
        OR Table_ObjIsVisible = False
    End If

    'If Table_ObjIsVisible = True then "X"
    If Table_ObjIsVisible Then
        Table_Obj.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy _
                Destination:=outputPasteRange

        'If Table_ObjIsVisible = False then "Y"
    Else:
        MsgBox Table_Obj.Name & " has been filtered to no visible records", vbInformation

    End If

End Sub



'This is a second alternative for checking a table to determine if any rows are visible on a filtered table. Currently no errors with this process.
'

OR

Sub TestEmptyTable()

    Dim tbl As ListObject
    Dim outputPasteRange As Range
    Dim tblIsVisible As Boolean

    Set tbl = ActiveSheet.ListObjects(1)
    Set outputPasteRange = Range("B15")

    Code_Sheet = "72"    'This value should be the name of the sheet variable.
    Source_Name = "PowerWorks A"    'This value should be the name of the source variable.

    Sheets("Health Maintenance Summary").Select

    'Applies the filtering to the table to determine results.
    ActiveSheet.ListObjects("Health_Maint_Table").Range.AutoFilter Field:=11, _
                                                                   Criteria1:=Source_Name, Operator:=xlAnd

    'Checks the filtered table to determine if there are any visible cells.
    'If cells are visible then set "tblIsVisible" to TRUE
    If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        tblIsVisible = True
    Else:
        tblIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
    End If

    'If cells were visible then "X"
    If tblIsVisible = True Then
        MsgBox tbl.Name & "Has visible records", vbInformation

        'Cells were not visible "Y"
    Else
        MsgBox tbl.Name & " has been filtered to no visible records", vbInformation

    End If

End Sub
