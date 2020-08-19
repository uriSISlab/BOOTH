VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VSAPBMD_Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LogProcessor

Private fileName As String
Private startTime As Date
Private state As BMDState
Private pollPassUsed As Boolean
Private writer As OutputWriter

Const loadingBallotLog As String = "Loading Ballot"
Const languageSelectedLog As String = "Language Selected"
Const removedBallotLog As String = "Voter removed ballot before read by BMD"
Const ballotActivatedLog As String = "Ballot Activated and User session is ended"
Const printedBallotLog As String = "Printed ballot successfully"
Const castBallotLog As String = "Casted ballot successfullly" ' Typo in "sucessfully" as it appears in logs
Const removedPrintedBallotLog As String = "Ballot removed after printing"
Const provisionalBallotEjectedLog As String = "Provisonal Ballot ejected" ' Typo in "provisional" as it appears in logs
Const pollPassScannedLog As String = "poll-pass successfully scanned"
Const votingSessionLockedLog As String = "voting session locked after timeout done (Ballot not in BMD)"
Const errorScanningBPMLog As String = "Error scanning BPM - BPM not present"
Const quitVotingLog As String = "Returning ballot - quit voting"
Const startLog As String = "screen diagnostics Successful"

'Enum to represent states of the VSAP BMD
Enum BMDState
    ' Initial state
    INIT = 0
    ' Loading state is entered after the ballot loading has begun
    Loading = 1
    ' Ballot has been activated, user can now vote (or cast their vote
    ' if ballot is already voted in)
    Activated = 2
    ' Ballot has been printed
    printed = 3
    ' An out-of-place removed ballot log has occured
    UnexpectedRemovedBallot = 4
End Enum

Public Sub class_initialize()
    clearState
End Sub

Public Sub LogProcessor_setWriter(ow As OutputWriter)
    Set writer = ow
End Sub

Public Sub LogProcessor_setFileName(fn As String)
    fileName = fn
End Sub

Private Sub clearState()
    ' TODO differentiate between state machine state and i/o state
    fileName = ""
    Set destination = Nothing
    pollPassUsed = False
    state = BMDState.INIT
End Sub

Private Sub writeLineToWriter(line As String)
    Dim lineArr() As String
    lineArr = Split(line, ", ")
    Dim fullLineArr() As String
    If Len(fileName) > 0 Then
        ReDim fullLineArr(Util.getStringArrayLength(lineArr) + 1)
        ' Copy contents of lineArr to fullLineArr first
        For i = LBound(lineArr) To UBound(lineArr)
            fullLineArr(i) = lineArr(i)
        Next i
        fullLineArr(UBound(fullLineArr)) = fileName
    Else
        fullLineArr = lineArr
    End If
    writer.writeLineArr fullLineArr
End Sub

Public Sub LogProcessor_writeHeader()
    If Len(fileName) > 0 Then
        writer.writeLine "Duration (mm:ss)", "Scan Type", "Ballot Cast Status", "Poll Pass Used", "Filename"
    Else
        writer.writeLine "Duration (mm:ss)", "Scan Type", "Ballot Cast Status", "Poll Pass Used"
    End If
End Sub

Public Sub writeBallotRemovedRecordNoTime()
    writeLineToWriter "-, Voter removed ballot before read by BMD, Unsuccessful, -"
End Sub

Public Sub writeBallotRemovedRecord(duration As String)
   writeLineToWriter duration & ", Voter removed ballot before read by BMD, Unsuccessful, -"
End Sub
Public Sub writeBallotCastRecord(duration As String, printed As Boolean, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration
    If printed Then
        outline = outline & ", Ballot printed and cast, Successful"
        If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
        End If
    Else
        outline = outline & ", Pre-printed ballot cast, Successful, -"
    End If
    writeLineToWriter outline
End Sub

Public Sub writeProvisionalBallotEjectedRecord(duration As String, printed As Boolean, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration
    If printed Then
        outline = outline & ", Provisional ballot printed and ejected, Successful"
        If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
        End If
    Else
        outline = outline & ", Pre-printed provisional ballot ejected, Successful, -"
    End If
    writeLineToWriter outline
End Sub

Public Sub writePrintedBallotRemovedRecord(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Ballot printed and removed, Unsuccessful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    writeLineToWriter outline
End Sub

Public Sub writeVotingTimedOutLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voting session timed out, Unsuccessful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    writeLineToWriter outline
End Sub

Public Sub writeBPMScanErrorLog(duration As String)
    writeLineToWriter duration & ", BPM Scan Error, Unsuccessful, -"
End Sub

Public Sub writeQuitVotingLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voter quit voting, Unsuccessful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    writeLineToWriter outline
End Sub

Public Sub writeMachineRestartedLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voting machine restarted unexpectedly, Unsuccessful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    writeLineToWriter outline
End Sub
Private Function parseDate(dateString As String) As Date
    Dim dateObj As Date
    Dim dateAndTime() As String
    dateAndTime = Split(dateString, "T")
    dateObj = dateAndTime(0) & " " & Split(dateAndTime(1), ".")(0)
    parseDate = dateObj
End Function
Public Sub LogProcessor_readLine(line As String)
    Dim thisTime As Date
    Dim thisLog As String
    Dim elements() As String
    
    elements = Split(line, "|")
    If UBound(elements) - LBound(elements) + 1 = 7 Then
        thisTime = parseDate(elements(1))
        thisLog = elements(6)
        If Not state = BMDState.INIT And Not state = BMDState.UnexpectedRemovedBallot And Len(startTime) > 0 Then
            If Util.getDifferenceMinutes(startTime, thisTime) > 60 Then
                ' A more than 60 minute difference probably indicates something suspicious.
                ' So we reset state here.
                state = BMDState.INIT
            End If
        End If
        If state = BMDState.INIT Then
            If Trim(thisLog) = loadingBallotLog Then
                startTime = thisTime
                state = BMDState.Loading
                pollPassUsed = False
            ElseIf Trim(thisLog) = removedBallotLog Then
                state = BMDState.UnexpectedRemovedBallot
            End If
        ElseIf state = BMDState.UnexpectedRemovedBallot Then
            If Trim(thisLog) = loadingBallotLog Then
                ' Since this loading ballot log appears after an unexpected removed ballot log,
                ' we will assume that this log line is the one that should have come before
                ' the unexpected one we encountered before.
                writeBallotRemovedRecordNoTime
                state = BMDState.INIT
            End If
        ElseIf state = BMDState.Loading Then
            If Trim(thisLog) = loadingBallotLog Then
                ' I wfe encounter another "loading" log at this state,
                ' the first one most probably came after a mis-ordered one
                writeBallotRemovedRecordNoTime
                pollPassUsed = False
                startTime = thisTime
            ElseIf Trim(thisLog) = removedBallotLog Then
                ' This means the ballot was removed from the machine before it could
                ' be read and activated. We need to record it and reset state.
                writeBallotRemovedRecord Util.getTimeDifference(startTime, thisTime)
                state = BMDState.INIT
            ElseIf Trim(thisLog) = ballotActivatedLog Then
                state = BMDState.Activated
            ElseIf Trim(thisLog) = errorScanningBPMLog Then
                writeBPMScanErrorLog Util.getTimeDifference(startTime, thisTime)
                state = BMDState.INIT
            ElseIf Trim(thisLog) = startLog Then
                writeMachineRestartedLog Util.getTimeDifference(startTime, thisTime), False
                state = BMDState.INIT
            ElseIf Trim(thisLog) = quitVotingLog Then
                writeQuitVotingLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            End If
        ElseIf state = BMDState.Activated Then
            If Trim(thisLog) = pollPassScannedLog Then
                pollPassUsed = True
            ElseIf Trim(thisLog) = printedBallotLog Then
                state = BMDState.printed
            ElseIf Trim(thisLog) = castBallotLog Then ' Typo in "sucessfully" as it appears in logs
                ' If the ballot was cast without being printed in this transaction.
                ' This means a pre-printed ballot was inserted.
                writeBallotCastRecord Util.getTimeDifference(startTime, thisTime), False, pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                writeProvisionalBallotEjectedRecord Util.getTimeDifference(startTime, thisTime), False, pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = votingSessionLockedLog Then
                writeVotingTimedOutLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = quitVotingLog Then
                writeQuitVotingLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = startLog Then
                writeMachineRestartedLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = languageSelectedLog Then
                writeQuitVotingLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            End If
        ElseIf state = BMDState.printed Then
            If Trim(thisLog) = removedPrintedBallotLog Then
                writePrintedBallotRemovedRecord Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = castBallotLog Then
                writeBallotCastRecord Util.getTimeDifference(startTime, thisTime), True, pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                writeProvisionalBallotEjectedRecord Util.getTimeDifference(startTime, thisTime), True, pollPassUsed
                state = BMDState.INIT
            ElseIf Trim(thisLog) = startLog Then
                writeMachineRestartedLog Util.getTimeDifference(startTime, thisTime), pollPassUsed
                state = BMDState.INIT
            End If
            ' TODO Find out whether provisional ballots can be cast just after printing
        End If
    End If
End Sub

Public Function LogProcessor_isThisLog(sheet As Worksheet) As Boolean
    LogProcessor_isThisLog = Trim(sheet.Range("B1")) = "Logger.js-Loading page-Manual Diagnostic Status"
End Function

Public Function LogProcessor_getSeparator() As String
    LogProcessor_getSeparator = "|"
End Function
