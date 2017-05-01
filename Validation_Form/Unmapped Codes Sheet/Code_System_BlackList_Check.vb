Formula for exact match
=IFERROR(INDEX(Sheet1!$A:$A,MATCH(Unmapped!C2,Sheet1!$A:$A,0)),0)


Formula for Partial MATCH    'If code system is within text on the list
{=INDEX(Sheet1!C2:C9,MATCH(TRUE,ISNUMBER(SEARCH(Sheet1!C2:C9,Unmapped!C3)),0))}
'MUST BE WITHIN BRACKETS. THIS IS AN ARRAY SEARCH
