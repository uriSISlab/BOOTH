VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VSAPBMD_Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public fileName As String
Public row As Long
Public destination As Object
Public Sub class_initialize()
    fileName = ""
    row = 2
    Set destination = Nothing
End Sub

Public Sub writeLine(line As String)
    Dim rowStr As String
    Dim rangeEnd As String
    rowStr = line
    rangeEnd = "D" & Me.row
    If Not Me.fileName = "" Then
        rowStr = rowStr & ", " & Me.fileName
        rangeEnd = "E" & Me.row
    End If
    If TypeOf Me.destination Is Worksheet Then
        Me.destination.Range("A" & Me.row, rangeEnd) = Split(rowStr, ", ")
    Else
        Me.destination.writeLine rowStr
    End If
    Me.row = Me.row + 1
End Sub
Public Sub writeHeader()
    Dim rowStr As String
    Dim rangeEnd As String
    rowStr = "Duration (mm:ss), Scan Type, Ballot Cast Status, Poll Pass Used"
    rangeEnd = "D1"
    If Not Me.fileName = "" Then
        rowStr = rowStr & ", Filename"
        rangeEnd = "E1"
    End If
    If TypeOf Me.destination Is Worksheet Then
        With Me.destination
            .Range("A1", rangeEnd) = Split(rowStr, ", ")
            .Range("A1", rangeEnd).Font.Bold = True
            .Range("A1", rangeEnd).HorizontalAlignment = xlCenter
        End With
    Else
        Me.destination.writeLine rowStr
    End If
End Sub
Public Sub writeBallotRemovedRecordNoTime()
    Me.writeLine "-, Voter removed ballot before read by BMD, Unsuccessful, -"
End Sub

Public Sub writeBallotRemovedRecord(duration As String)
    Me.writeLine duration & ", Voter removed ballot before read by BMD, Unsuccessful, -"
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
    Me.writeLine outline
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
    Me.writeLine outline
End Sub

Public Sub writePrintedBallotRemovedRecord(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Ballot printed and removed, Unsuccesful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    Me.writeLine outline
End Sub

Public Sub writeVotingTimedOutLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voting session timed out, Unsuccesful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    Me.writeLine outline
End Sub

Public Sub writeBPMScanErrorLog(duration As String)
    Dim outline As String
    outline = duration & ", BPM Scan Error, Unsuccesful, -"
    Me.writeLine outline
End Sub

Public Sub writeQuitVotingLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voter quit voting, Unsuccesful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    Me.writeLine outline
End Sub

Public Sub writeMachineRestartedLog(duration As String, pollPassUsed As Boolean)
    Dim outline As String
    outline = duration & ", Voting machine restarted unexpectedly, Unsuccesful"
    If pollPassUsed Then
            outline = outline & ", Yes"
        Else
            outline = outline & ", No"
    End If
    Me.writeLine outline
End Sub
