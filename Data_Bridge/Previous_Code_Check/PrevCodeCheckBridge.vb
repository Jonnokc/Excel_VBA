Sub PrevCodeCheckBridge()

Dim sheet As Worksheet


    ' Checks to see if the CodeSystemCheck Sheet already exists.

    Application.DisplayAlerts = False

    For Each sheet In Worksheets
        If sheet.Name = "CodeSystemCheck" Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "CodeSystemCheck"
    End With


' PRIMARY - Imports the previously reviewed unmapped
''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With Sheets("CodeSystemCheck").ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Y:\Data Intelligence\Code_Submittion_Database\CodeFeedba" _
        , _
        "ckDatabase.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:" _
        , _
        "Database Password=BORIS;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:" _
        , _
        "Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=" _
        , _
        "False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:" _
        , _
        "Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass Choice" _
        , "Field Validation=False"), Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array("QUnmappedEvCodeCheck")
        .SourceDataFile = _
        "Y:\Data Intelligence\Code_Submittion_Database\CodeFeedbackDatabase.accdb"
        .ListObject.DisplayName = "PreviousCodes"
        .Refresh BackgroundQuery:=False
    End With
End Sub
