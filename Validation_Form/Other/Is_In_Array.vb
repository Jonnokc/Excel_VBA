Sub Demo()

' Function that checks an array to see if the variable is within it.
' In this case the function checks the variable "i" to see if it is within the array.
' If it is then it returns "True" if not then it returns "False"

    Dim arr(2) As String
    Dim i As Variant

    arr(0) = "100"
    arr(1) = "50"
    arr(2) = "2"
    i = "20"
    MsgBox IsInArray(CStr(i), arr)
End Sub


Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
