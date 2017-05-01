Sub Testing()

    Dim Lookup As Variant
    Dim cell_Lookup As Long
    Dim sResult_Value As String

    For Each cell In Range("Event_Range")

        cell_Lookup = cell.Offset(0, 5).Value    'The cell value you want to look for a match with

        On Error GoTo NoMatch
        sResult = Application.VLookup(cell_Lookup, Range("Validated_Range"), 6, False)

        If sResult = "Validated" Then
            sResult_Value = sResult
            cell.Value = sResult_Value

        Else
NoMatch:
            cell.Value = "0"
            Resume ClearError
ClearError:
        End If

    Next cell
End Sub
