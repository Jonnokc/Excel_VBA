Sub String_Combiner()

' Program continues to add values to a string.

Dim Source_Combined As String
Dim Current_Source As String

For Each Source_Name In Range("Sources_List")
  Current_Source = Source_Name

  Source_Combined = Source_Combined & Current_Source & vbNewLine


Next Source_Name

MsgBox("Here is a list of the sources" & vbNewLine & vbNewLine & Source_Combined)

End Sub
