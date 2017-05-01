Sub Testing()

'
'This code will take values from a table and put them in an arrao.
'Then it Will perform changes to the data within the array and then write the array back to the sheet.
'This changes the values all at once instead of one at a time.
'
'


    Dim DataRange As Variant    'Declare array variable
    Dim Irow As Long    'The row variable
    Dim Icol As Integer    'The column variable if you need to loop through multiple columns
    Dim MyType As Variant    'Variable used to store column value


    DataRange = Range("Data_Range").Value    'writes the named data range to the array variable

    For Irow = 1 To UBound(DataRange)    'Loops through all rows within the range.
        MyType = DataRange(Irow, 1)    'Assigns current value to a variable

        If MyType = "PowerForm" Then
            DataRange(Irow, 13) = "Yes"    'If value is true, then update value in row IView, Column 13 to "X"
        End If
    Next Irow
    Range("Data_Range").Value = DataRange    'Write the updated DataRange Array to the excel file

End Sub
