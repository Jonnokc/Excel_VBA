Private Sub Remove_Duplicates()

' A duplicate checker. It runs against the Raw Concept To Measure file looking for duplicates across col A, B, C.

    Sheets("Raw_Concept_To_Measure").Select
    Application.Goto Reference:="Raw_Table_Range"
    Selection.RemoveDuplicates Columns:=Array(1, 2, 3), _
            Header:=xlYes
End Sub
