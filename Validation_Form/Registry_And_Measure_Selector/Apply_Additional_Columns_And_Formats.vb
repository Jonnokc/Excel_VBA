Private Sub Apply_Additional_Columns_And_Formulas()

' Overall applies column headers and formulas to additional rows

    Sheets("Raw_Concept_To_Measure").Select

    ' Creates the column headers for the raw concet to measure sheet

    Range("D1").Select
    ActiveCell.Formula = "Flagged_Result"
    Range("E1").Select
    ActiveCell.Formula = "Registry Check"
    Range("F1").Select
    ActiveCell.Formula = "Measure Check"
    Range("G1").Select
    ActiveCell.Formula = "Concept Check"
    Range("H1").Select
    ActiveCell.Formula = "Registry and Measure"
    Range("J1").Select
    ActiveCell.Formula = "Combined - All"
    Range("K1").Select
    ActiveCell.Formula = "Flagged_Result2"

    ' Applies the formulas to the raw concept to measure additional columns

    Range("E2").Select
    ActiveCell.Formula = _
    "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!A2,Pivot!$AA:$AA,0))"

    Range("F2").Select
    ActiveCell.Formula = _
    "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!H2,Pivot!AD:AD,0))"

    Range("G2").Select
    ActiveCell.Formula = _
    "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!J2,Pivot!AE:AE,0))"

    Range("H2").Select
    ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2)"

    Range("J2").Select
    ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2,""|"",C2)"

    Range("K2").Select
    ActiveCell.Formula = "=D2"

    Range("D2").Select
    ActiveCell.Formula = _
    "=IF(G2=""Yes"", ""Yes"",IF(AND(E2=""Yes"",F2<>""No"",G2<>""No""),""Yes"",IF(AND(F2=""Yes"",G2<>""No""),""Yes"",""No"")))"

    ' Creates the column headers for the pivot table

    Sheets("Pivot").Select
    Range("D1").Select
    ActiveCell.Formula = "Flagged"
    Range("E1").Select
    ActiveCell.Formula = "Y/N"
    Range("AD1").Select
    ActiveCell.Formula = "Registry and Measure"
    Range("AE1").Select
    ActiveCell.Formula = "All Combined"
    Range("AG1").Select
    ActiveCell.Formula = "Result"

    ' Applies the formulas for the pivot table new columns

    Range("AD2").Select
    ActiveCell.Formula = "=CONCATENATE(AA2,""|"",AB2)"
    Range("AE2").Select
    ActiveCell.Formula = "=CONCATENATE(AA2,""|"",AB2,""|"",AC2)"
    Range("AF2").Select
    ActiveCell.Formula = "=E2"
    Range("D2").Select
    ActiveCell.Formula = _
    "=IFERROR(VLOOKUP(AE2,Raw_Concept_To_Measure!J:K,2,""FALSE""),"""")"

    Range("A1").Select

    ' Ensures calculations is set to automatic This is needed for lookups to work
    Application.Calculation = xlCalculationAutomatic


End Sub
