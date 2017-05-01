Function WORDSAME(rngA As Range, rngB As Range) As String

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
                strTemp = strTemp & WordsA(nextA) & " "
                Exit For
            End If
        Next nextA
    Next nextB


    WORDSAME = Trim(StrConv(strTemp, vbProperCase))


End Function
