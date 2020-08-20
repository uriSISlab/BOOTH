VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DICE_Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LogProcessor

Private Enum DICEState
    ' The device is ready to accept a new ballot
    Ready = 0
    ' A ballot has just been inserted into the device for scanning
    BallotInserted = 1
    ' A ballot was just cast into the machine
    BallotCast = 2
    ' An accessible voting session has been started
    AccessibleStarted = 3
    ' A ballot has been inserted for accessible voting
    AccessibleBallotInserted = 4
    ' A ballot has been accepted for accessible voting
    AccessibleBallotAccepted = 5
End Enum

Private Const s_ballotInserted As String = "[Voting] A ballot has been inserted into the unit."
Private Const s_invokingBallotReview As String = "[Voting] Invoking ballot review for the current ballot."
Private Const s_ballotCast As String = "[Voting] Ballot successfully cast and dropped into ballot box"
Private Const s_endOfSession As String = "[Voting] The end of the current session. The system is ready to accept a new ballot."

Private Const s_accessibleStarted As String = "[Admin] Accessible Voting (Contest by Contest) started"
Private Const s_accessibleBallotInserted As String = "[Accessible Voting] A ballot has been inserted into the unit"
Private Const s_accessibleBallotAccepted As String = "[Accessible Voting] A ballot has been accepted by the system."
Private Const s_accessibleMarkingCompleted As String = "[Accessible Voting] A ballot marking has been completed."

Private Const s_ballotMisread As String = "[Pixel Count] Ballot misread."
Private Const s_paperJam As String = "[Scanner] Error scanning ballot: Possible paper jam. Code:"
Private Const s_noManifestation As String = "[Pixel Count] No ballot manifestation for determined ballot Id"

Private state As DICEState
Private startTime As Date
Private writer As OutputWriter
Private misreads As Integer
Private reviewed As Boolean
Private paperJam As Boolean
Private ballotNotRecognized As Boolean
Private lastTimestamp As Date
Private fileName As String

Public Sub class_initialize()
    clearState
    fileName = ""
End Sub

Public Sub LogProcessor_setWriter(ow As OutputWriter)
    Set writer = ow
End Sub

Public Sub LogProcessor_setFileName(fn As String)
    fileName = fn
End Sub

Private Sub clearState()
    state = DICEState.Ready
    misreads = 0
    reviewed = False
    paperJam = False
    ballotNotRecognized = False
End Sub

Private Sub writeToWriter(lineArr() As String)
    ' Write the given array to writer after possibly appending the filename
    If Len(fileName) > 0 Then
        writer.writeLineArr Util.appendToArray(lineArr, fileName)
    Else
        writer.writeLineArr lineArr
    End If
End Sub

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
    Dim lineArr(4) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Ballot cast normally"
    lineArr(3) = CStr(misreads)
    If reviewed Then
        lineArr(4) = "Yes"
    Else
        lineArr(4) = "No"
    End If
    writeToWriter lineArr
End Sub

Private Sub writeBallotNotCastRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(4) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Ballot not cast"
    lineArr(3) = misreads
    If reviewed Then
        lineArr(4) = "Yes"
    Else
        lineArr(4) = "No"
    End If
    writeToWriter lineArr
End Sub

Private Sub writePaperJamRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(4) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Paper jam when accepting ballot"
    lineArr(3) = misreads
    lineArr(4) = "-"
    writeToWriter lineArr
End Sub

Private Sub writeBallotNotRecognizedRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(4) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Ballot not recognized"
    lineArr(3) = misreads
    lineArr(4) = "-"
    writeToWriter lineArr
End Sub

Private Sub writeAccessibleBallotMarkedRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(4) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Accessible voting : ballot marked"
    lineArr(3) = "-"
    ' TODO find out if accessible marking mode review is possible
    lineArr(4) = "-"
    writeToWriter lineArr
End Sub

Private Sub writeRecordAccordingToState(startTimestamp As Date, endTimestamp As Date)
    If state = BallotInserted Then
        If paperJam Then
            writePaperJamRecord startTimestamp, endTimestamp
        ElseIf ballotNotRecognized Then
            writeBallotNotRecognizedRecord startTimestamp, endTimestamp
        Else
            writeBallotNotCastRecord startTimestamp, endTimestamp
        End If
    ElseIf state = BallotCast Then
        writeBallotCastNormalRecord startTimestamp, endTimestamp
    End If
End Sub

Public Sub LogProcessor_writeHeader()
    ' Write the header
    Dim headerArr() As String
    If Len(fileName) > 0 Then
        headerArr = Split("Duration,Timestamp,Event,Misreads,Ballot Reviewed,Filename", ",")
    Else
        headerArr = Split("Duration,Timestamp,Event,Misreads,Ballot Reviewed", ",")
    End If
    writer.writeLineArr headerArr
End Sub

Public Sub LogProcessor_readLine(line As String)
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
    
    If rest = s_accessibleStarted Then
        state = AccessibleStarted
        startTime = thisTime
    ElseIf rest = s_accessibleMarkingCompleted Then
        state = Ready
        writeAccessibleBallotMarkedRecord startTime, thisTime
    ElseIf rest = s_ballotInserted Then
        If Util.getDifferenceMinutes(lastTimestamp, thisTime) > 2 Then
            ' If there's a large difference in time here, the earlier session
            ' was probably abandoned due to something. So we clear the state.
            clearState
        End If
        state = BallotInserted
        startTime = thisTime
    ElseIf rest = s_ballotCast Then
        state = BallotCast
    ElseIf rest = s_endOfSession Then
        writeRecordAccordingToState startTime, thisTime
        clearState
    ElseIf rest = s_ballotMisread Then
        misreads = misreads + 1
    ElseIf rest = s_invokingBallotReview Then
        reviewed = True
    ElseIf InStr(rest, s_paperJam) <> 0 Then
        paperJam = True
    ElseIf rest = s_noManifestation Then
        ballotNotRecognized = True
    End If
    
    lastTimestamp = thisTime
End Sub

Public Function LogProcessor_isThisLog(sheet As Worksheet) As Boolean
    LogProcessor_isThisLog = InStr(sheet.Cells(1, 2), "Logging service initialized") <> 0
End Function

Public Function LogProcessor_getSeparator() As String
    LogProcessor_getSeparator = " "
End Function
