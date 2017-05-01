'
' This is the most efficient way to handle copy and paste.
' Range to copy is selected then output range is entered and is pasted even without code specifying
' Paste.
'
'
'

Sub CopyRange2()

    Range("A1:A12").Copy Range("C1")

End Sub
