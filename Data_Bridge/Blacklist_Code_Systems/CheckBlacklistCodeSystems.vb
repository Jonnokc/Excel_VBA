Sub CheckBlacklistCodeSystems()

Dim ws As Worksheet
Dim Header_Check As Boolean
Dim UnmappedCodeIdHeader As Variant
Dim Header_User_Response As Variant
Dim SheetArray As Variant
Dim HeaderLocations As Variant
Dim HeaderNames As Variant
Dim BlacklistArray As Variant
Dim BlacklistAnswerArray As Variant
Dim UnmappedCodeSystemArray As Variant
Dim BlackListHeader As String
Dim LR As Long
Dim Lookup As Variant
Dim cell_Lookup As Variant
Dim sResult_Value As Variant

SheetArray = Array("Unmapped Raw", "BlackList_Table")
HeaderLocations = Array("Coding System ID", "Blacklist Location")
HeaderNames = Array("Coding System ID", "BlacklistedCodeSystem")



' SUB - Finds the column for the code system ID on the unmapped codes Sheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 0 to UBound(SheetArray)
' Sets header row range
Sheets(SheetArray(i)).Select
Range("A1").Select
Range("A2", Selection.End(xlToRight)).Name = "Header_row"

' Finds the Code System ID header column

  For each header in Range("Header_row")
  Header_Check = False
    If LCase(header) = LCase("Coding System ID") Or LCase(header) = LCase("BlacklistedCodeSystem") Then
      HeaderLocations(i) = Mid(header.Address, 2, 1)
      Header_Check = True
      exit For
    End If
  next header

  ' If no header was found then prompt the user for the column or allow the user to cancel the program
  If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & HeaderNames(i) & " on the " & SheetArray(i) & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the column containing the Coding System ID's.","If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
          GoTo User_Exit
      Else
          HeaderLocations(i) = UCase(Header_User_Response)
      End If
  End If
Next i


' PRIMARY - Check Code Systems against the blacklist
'''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' SUB - Create a new column on unmapped code sheet at the end which will be the blacklist check response

  Sheets(SheetArray(0)).Select
  ' Checks if Blacklist Check column already exists

  Range("B2").Select
  Range(Selection, Selection.End(xlToRight)).Name = "Header_row"

  Header_Check = False
  For each header in Range("Header_row")
    If LCase(header) = LCase("Blacklist Check") then
      BlackListHeader = Mid(header.Address, 2, 1)
      Header_Check = True
      exit For
    End If
  next header


  ' Finds and sets letter of next available column on the unmapped code sheet
  If Header_Check = False Then
    BlackListHeader = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0,1).Address,2,1)
    Range(BlackListHeader & "2") = "Blacklist Check"
  End If


  ' SUB - Assigns Unmapped Codes Code System ID's to array in memory
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets(SheetArray(0)).Select
  LR = Range(HeaderLocations(0) & Rows.Count).End(xlUp).Row
  Range(HeaderLocations(0)&"3:"  & HeaderLocations(0) & LR).SpecialCells(xlCellTypeVisible).Name = "UnmappedCodeID_Range"

  UnmappedCodeSystemArray = Range("UnmappedCodeID_Range").Value

  ' SUB - Set blacklist answer range to array in memory
  Sheets(SheetArray(0)).Select
  Range(BlackListHeader &"3:" & BlackListHeader & LR).SpecialCells(xlCellTypeVisible).Name = "BlackListAnswers"

  BlacklistAnswerArray = Range("BlackListAnswers")

  ' SUB - Assigns Blacklist Codes from the BlackList Table to an array in memory
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets(SheetArray(1)).Select
  LR = Range(HeaderLocations(1) & Rows.Count).End(xlUp).Row
  Range(HeaderLocations(1) &"1:" & HeaderLocations(1) & LR).SpecialCells(xlCellTypeVisible).Name = "BlackList_Range"

  BlacklistArray = Range("BlackList_Range").Value



  For i = 1 to UBound(UnmappedCodeSystemArray)

      cell_Lookup = UnmappedCodeSystemArray(i,1)  'The cell value you want to look for a match with

      sResult = Application.VLookup(cell_Lookup, Range("BlackList_Range"), 1, False)

      If IsError(sResult) Then
          sResult_Value = ""
          BlacklistAnswerArray(i,1) = sResult_Value
      Else
          BlacklistAnswerArray(i,1) = "On Code System Blacklist"

      End If

  Next i

  ' Deletes extra sheet
  Application.DisplayAlerts = False

  For Each sheet In Worksheets
      If sheet.Name = "BlackList_Table" _
              Then
          sheet.Delete
      End If
  Next sheet

  Application.DisplayAlerts = True

  'Write the updated DataRange Array to the excel file
  Range("BlackListAnswers").Value = BlacklistAnswerArray

  Sheets(SheetArray(0)).Select
  MsgBox ("BORIS has completed the Blacklist Code System check")

  Exit Sub

  User_Exit:
    MsgBox ("Exiting per user action")


End Sub
