
Function WORDDIF(rngA As Range, rngB As Range) As String

    Dim WordsA As Variant
    Dim WordsB As Variant
    Dim nextA As Long
    Dim nextB As Long
    Dim strTemp As String

    WordsA = Split(rngA.Text, " ")
    WordsB = Split(rngB.Text, " ")

    For nextB = LBound(WordsB) To UBound(WordsB)
        For nextA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(LCase(nextA)), WordsB(LCase(nextB)), vbTextCompare) = 0 Then
                WordsA(nextA) = vbStringA
                WordsB(nextB) = vbStringB
                Exit For
            End If
        Next nextA
    Next nextB

    For nextA = LBound(WordsA) To UBound(WordsA)
        If WordsA(LCase(nextA)) <> LCase(vbStringA) Then strTemp = strTemp & LCase(WordsA(nextA)) & " "
    Next nextA

    For nextB = LBound(WordsB) To UBound(WordsB)
        If WordsB(LCase(nextB)) <> LCase(vbStringB) Then strTemp = strTemp & LCase(WordsB(nextB)) & " "
    Next nextB


    WORDDIF = Trim(StrConv(strTemp, vbProperCase))

End Function
