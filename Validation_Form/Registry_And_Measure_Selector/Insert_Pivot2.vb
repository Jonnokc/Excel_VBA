Private Sub Insert_Pivot_2()

' Inserts second pivot table which is used to create CONCATENATE fields for matching to auto populate child rows.

    Dim PT As PivotTable

    Sheets("Raw_Pivot").Select

    Range("D1").Select
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveSheet.PivotTables("Concept_Pivot_Table").RepeatAllLabels xlRepeatLabels
    Set PT = ActiveSheet.PivotTables(1)
    PT.TableRange1.Select
    Selection.Copy
    Sheets("Pivot").Select
    Application.Goto Reference:="R1C27"
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("A1").Select
End Sub
