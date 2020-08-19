VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DICX_Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LogProcessor

Private Enum DICXState
    ' The device is ready to accept a new ballot
    Ready = 0
    ' A new voting session has been started
    Started = 1
    ' Voter has voted and is reviewing ballot
    BallotReview = 2
    ' Ballot is being cast
    BallotCasting = 3
End Enum

Private Const s_startingVotingSession = "Starting new voting session."
Private Const s_ballotPresented = "Ballot is presented to voter."
Private Const s_ballotReview = "Ballot review"
Private Const s_prepareBallot = "Prepare ballot for cast."
Private Const s_ballotCast = "Ballot cast successfully"

Private state As DICXState
Private startTime As Date
Private writer As OutputWriter
Private fileName As String

Public Sub class_initialize()
    clearState
    fileName = ""
End Sub

Public Sub LogProcessor_setWriter(ow As OutputWriter)
    Set writer = ow
End Sub

Private Sub clearState()
    state = DICXState.Ready
End Sub

Private Sub writeToWriter(lineArr() As String)
    ' Write the given array to writer after possibly appending the filename
    If Len(fileName) > 0 Then
        writer.writeLineArr Util.appendToArray(lineArr, fileName)
    Else
        writer.writeLineArr lineArr
    End If
End Sub

Private Sub writeBallotCastNormalRecord(startTimestamp As Date, endTimestamp As Date)
    Dim lineArr(2) As String
    lineArr(0) = Util.getTimeDifference(startTimestamp, endTimestamp)
    lineArr(1) = endTimestamp
    lineArr(2) = "Ballot cast normally"
    writeToWriter lineArr
End Sub

Public Sub LogProcessor_setFileName(fn As String)
    fileName = fn
End Sub

Public Sub LogProcessor_writeHeader()
    'Write the header
    If Len(fileName) > 0 Then
        writer.writeLine "Duration", "Timestamp", "Event", "Filename"
    Else
        writer.writeLine "Duration", "Timestamp", "Event"
    End If
End Sub

Public Sub LogProcessor_readLine(line As String)
    Dim thisTime As Date
    Dim rest As String
    
    If Len(line) < 24 Then
        ' There is nothing useful in the line, do nothing
        Exit Sub
    End If

    ' Check the position of the timestamp-log divider
    If InStr(1, line, " - ") = 0 Then
        ' Do nothing if the divider was not found
        Exit Sub
    End If
    
    thisTime = CDate(Left(line, 19)) ' Timestamp is in the first 19 characters
    rest = Trim(Mid(line, 23)) ' Next three characters are " - "
    
    If rest = s_startingVotingSession Then
        state = Started
        startTime = thisTime
    ElseIf rest = s_ballotReview Then
        state = BallotReview
    ElseIf rest = s_prepareBallot Then
        state = BallotCasting
    ElseIf rest = s_ballotCast Then
        state = Ready
        writeBallotCastNormalRecord startTime, thisTime
    End If
    lastTimestamp = thisTime
End Sub

Public Function LogProcessor_isThisLog(sheet As Worksheet) As Boolean
    Dim idRange As Range
    Set idRange = sheet.UsedRange.Find(What:="Audit Log file is saved.", _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
    LogProcessor_isThisLog = Not idRange Is Nothing
End Function


Public Function LogProcessor_getSeparator() As String
    LogProcessor_getSeparator = " - "
End Function
