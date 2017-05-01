Function CONCEPTMATCH(rngA As Range, rngB As Range) As String

    Dim WordsA As Variant
    Dim WordsB As Variant
    Dim nextA As Long
    Dim nextB As Long
    Dim strTemp As String
    Dim matchCount As Integer
    Dim WordsACount As Single
    Dim WordsBCount As Single
    Dim uniqueCount As Single
    Dim finalMatch As Single


    WordsA = Split(rngA.Text, " ")
    WordsB = Split(rngB.Text, " ")

    For nextA = LBound(WordsA) To UBound(WordsA)
        WordsACount = WordsACount + 1
    Next nextA

    For nextB = LBound(WordsB) To UBound(WordsB)
        WordsBCount = WordsBCount + 1
    Next nextB

    For nextB = LBound(WordsB) To UBound(WordsB)
        For nextA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(LCase(nextA)), WordsB(LCase(nextB)), vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                Exit For
            End If
        Next nextA
    Next nextB



    For nextA = LBound(WordsA) To UBound(WordsA)
        If WordsA(LCase(nextA)) <> LCase(vbStringA) Then uniqueCount = uniqueCount + 1
    Next nextA

    For nextB = LBound(WordsB) To UBound(WordsB)
        If WordsB(LCase(nextB)) <> LCase(vbStringB) Then uniqueCount = uniqueCount + 1
    Next nextB


    For nextB = LBound(WordsB) To UBound(WordsB)
        For nextA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(LCase(nextA)), WordsB(LCase(nextB)), vbTextCompare) = 0 Then
                WordsA(nextA) = vbStringA
                WordsB(nextB) = vbStringB
                Exit For
            End If
        Next nextA
    Next nextB


    finalMatch = matchCount / (WordsACount - 1)


    CONCEPTMATCH = (Format(finalMatch, "0.00"))



End Function
