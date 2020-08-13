VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DICE_Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Enum DICEState
    ' The device is ready to accept a new ballot
    Ready = 0
    ' A ballot has just been inserted into the device for scanning
    BallotInserted = 1
End Enum

Private Const s_ballotInserted As String = "[Voting] A ballot has been inserted into the unit."
Private Const s_ballotCast As String = "[Voting] Ballot successfully cast and dropped into ballot box"
Private Const s_endOfSession As String = "[Voting] The end of the current session. The system is ready to accept a new ballot."

Private state As DICEState
Private startTime As Date
Private writer As OutputWriter

Public Sub class_initialize()
    state = DICEState.Ready
End Sub

Public Sub setWriter(ow As OutputWriter)
    Set writer = ow
End Sub

Private Function getTimeDifference(startTime As Date, endTime As Date) As String
    Dim mins As String
    Dim secs As String
    mins = Format(CStr(CLng(DateDiff("n", startTime, endTime))), "00")
    secs = Format(CStr(CLng(DateDiff("s", startTime, endTime)) Mod 60), "00")
    getTimeDifference = mins & ":" & secs
End Function

Private Function getMonthOrder(month As String) As Integer
    Dim m As String
    Dim o As Integer
    Dim months() As String
    m = Left(LCase(month), 3)
    months = Split("jan feb mar apr may jun jul aug sep oct nov dec", " ")
    o = -1
    For i = 0 To 11
        If months(i) = m Then
            o = i
            Exit For
        End If
    Next i
    getMonthOrder = o + 1
End Function

Private Function get_timestamp_from_DICE_line(line As String) As Date
    Dim line_words() As String
    Dim time As String
    Dim d, m, y As Integer
    line_words = Split(line, " ")
    d = CInt(line_words(0))
    m = getMonthOrder(line_words(1))
    y = CInt(line_words(2))
    time = line_words(3)
    get_timestamp_from_DICE_line = CDate(CStr(m) & "/" & CStr(d) & "/" & CStr(y) & " " & time)
End Function

Private Sub writeBallotCastNormalRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(2) As String
    lineArr(0) = getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = "Ballot cast normally"
    writer.writeLine lineArr
End Sub

Public Sub readLine(line As String)
    Dim thisTime As Date
    Dim col_pos As Integer
    Dim rest As String

    ' The timestamp is 20 characters long
    If Len(line) < 21 Then
        ' There is nothing useful in the line, do nothing
        Exit Sub
    End If
    ' Get the position of the first colon after the timestamp
    col_pos = InStr(21, line, ":")
    If col_pos = 0 Then
        ' Do nothing if the colon was not found
        Exit Sub
    End If
    
    thisTime = get_timestamp_from_DICE_line(line)
    rest = Trim(Mid(line, col_pos + 1))
    
    If rest = s_ballotInserted Then
        state = BallotInserted
        startTime = thisTime
    ElseIf rest = s_ballotCast Then
        state = Ready
        writeBallotCastNormalRecord startTime, thisTime
    End If
End Sub

