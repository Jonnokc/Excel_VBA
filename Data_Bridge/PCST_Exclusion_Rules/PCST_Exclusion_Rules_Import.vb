Sub get_data()
'
' get_data Macro
'
'
    ' Checks to see if the CodeSystemCheck Sheet already exists.

    Application.DisplayAlerts = False

    For Each Sheet In Worksheets
        If Sheet.Name = "ExclusionRules" Then
            Sheet.Delete
        End If
    Next Sheet

    Application.DisplayAlerts = True

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "ExclusionRules"
    End With

    With Sheets("ExclusionRules").ListObjects.Add(SourceType:=0, Source:=Array( _
      "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Y:\Data Intelligence\Code_Submittion_Database\CodeFeedba" _
      , _
      "ckDatabase.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:" _
      , _
      "Database Password="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Glo" _
      , _
      "bal Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=Fal" _
      , _
      "se;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Sup" _
      , _
      "port Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceFie" _
      , "ld Validation=False"), Destination:=Range("$A$1")).QueryTable
      .CommandType = xlCmdTable
      .CommandText = Array("PCSTExclusionRules")
      .SourceDataFile = _
      "Y:\Data Intelligence\Code_Submittion_Database\CodeFeedbackDatabase.accdb"
      .ListObject.DisplayName = "ExclusionRules"
      .Refresh BackgroundQuery:=False
    End With



End Sub
