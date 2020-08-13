Attribute VB_Name = "Util"
Public Function getStringArrayLength(arr() As String) As Long
    getStringArrayLength = UBound(arr) - LBound(arr) + 1
End Function

Public Function getLetterFromNumber(number As Integer) As String
    ' Return the corresponding letter from a number, 1 returns "A", 2 returns "B", and so on.
    getLetterFromNumber = Chr(Asc("A") + number - 1)
End Function

Public Function getDifferenceMinutes(startTime As Date, endTime As Date) As Integer
    getDifferenceMinutes = DateDiff("s", startTime, endTime) \ 60
End Function

Public Function getTimeDifference(startTime As Date, endTime As Date) As String
    Dim mins As String
    Dim secs As String
    Dim totalSecs As Long
    totalSecs = CLng(DateDiff("s", startTime, endTime))
    mins = Format(CStr(totalSecs \ 60), "00")
    secs = Format(CStr(totalSecs Mod 60), "00")
    getTimeDifference = mins & ":" & secs
End Function

Public Function getVSAPBMDHeader() As String()
    Dim header() As String
    header = Split("Duration (mm:ss), Scan Type, Ballot Cast Status, Poll Pass Used", ", ")
    getHeader = header
End Function
