Private Sub Summary_Add_HyperLinks()
'
' Summary_Validated_Sheet_Link Macro
'
	Sheets("Summary View").Select
	Range("E2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Potential Mapping Issues'!A1"

	Range("F2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Unmapped Codes'!A1"

	Range("G2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Clinical Documentation'!A1"

	Range("B2").Select


End Sub
