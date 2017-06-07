Sub Validated_Mappings_Setup()

Dim Val_Headers As Variant
Dim Rng As Range
Dim cell As Range


'Helps improve performance
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Val_Headers = Array("Keywords (ATTR)", "Concept Alias", "Concepts Where Code is Normalized", "Measure", "Raw Code Display", "Raw Code ID", "Raw Code System", "Registry", "Standard Code Display", "Standard Code ID", "Standard Code System Display", "Status")


' SUB - Stores the location of the columns by their header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Finds Column Header Locations for The Combined Data Export
Sheet1.Select
Range("A1").Select
Range("A1", Selection.End(xlToRight)).Name = "Header_row"

For i = 0 To UBound(Val_Headers)
  Header_Check = False
  For Each Header In Range("Header_row")
    If LCase(Val_Headers(i)) = LCase(Header) Then
      Val_Headers(i) = Mid(Header.Address, 2, 1)
      Header_Check = True
      Exit For
    End If
  Next Header
  If Header_Check = False Then
    Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Val_Headers(i) & "'" & " on the " & Sheet1.Name & " Sheet....." & vbNewLine & vbNewLine & "However all is not lost! BORIS and you can do this!" & vbNewLine & vbNewLine & "To resolve the issue BORIS needs you to enter the letter of a column to use in place of the one he couldn't find." & vbNewLine & vbNewLine & "Look at the excel sheet behind this box and enter (in uppercase) the letter of the column you want to use in place of the missing one." & vbNewLine & vbNewLine & "If you don't want to replace data from another column in place of the missing one then enter the letter of an empty column(like T or something). If you would rather fix the issue within the file or program then click cancel.", "If I am BORIS who are you?")

    'If user hits cancel then close program.
    If Header_User_Response = vbNullString Then
      GoTo User_Exit
    Else
      Val_Headers(i) = Header_User_Response
    End If
  End If

Next i


  ' SUB - CONVERTS Numbers stored as text to numbers
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Selects all cells in column
    Range(Val_Headers(5) & "2:" & Val_Headers(5) & Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Number_Check"
    Set Rng = Range("Number_Check")

    'If cell is a number, then convert format to number
    For Each cell In Rng
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell

    're-enables updates
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

User_Exit:


End Sub
