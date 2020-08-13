Attribute VB_Name = "Util"
Public Function getStringArrayLength(arr() As String) As Long
    getStringArrayLength = UBound(arr) - LBound(arr) + 1
End Function

Public Function getLetterFromNumber(number As Integer) As String
    ' Return the corresponding letter from a number, 1 returns "A", 2 returns "B", and so on.
    getLetterFromNumber = Chr(Asc("A") + number - 1)
End Function
