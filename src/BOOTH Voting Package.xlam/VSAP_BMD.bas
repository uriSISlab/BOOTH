Attribute VB_Name = "VSAP_BMD"

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

Sub Import_VSAPBMD_data(control As IRibbonControl)
    
    Dim lrow As Long
    Dim l2row As Long
    Dim l3row As Long
    Dim i As Long
    Dim t As Long
    Dim ret1 As String
    Dim j As Long
    Dim intResult As Long
    Dim strPath As String
    Dim arraylen As Long
    Dim tbook As ThisWorkbook
    Dim f As Long
    Dim w As Long
    
    'When File Explorer opens, only display text log files
    With Application.FileDialog(msoFileDialogFilePicker)
    Application.FileDialog(msoFileDialogFilePicker).Filters.Clear
    Application.FileDialog(msoFileDialogFilePicker).Filters.Add "Log files", "*.log"
    End With
    
    'Open the file explorer and allow the selection of multiple files
    Application.FileDialog(msoFileDialogFilePicker).Show
    Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = True
    
    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False
    
    'Loop to process multiple files consecutively
    For j = 1 To Application.FileDialog(msoFileDialogFilePicker).SelectedItems.count
    
        'Adds an additional Worksheet to write VSAP BMD data to if only one sheet is open
        If ActiveWorkbook.Sheets.count = 1 Then
            ActiveWorkbook.Sheets.Add after:=ActiveSheet
        End If
    
        'Check for duplicate precincts and delete the duplicate sheets
        c = 1
        While c < ActiveWorkbook.Sheets.count + 1
            If ActiveWorkbook.Sheets(c).Name = "Precinct " & Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".txt", ""), 10) Then
                GoTo skipit
            Else: c = c + 1
            End If
        Wend
    
        'Add an additional sheet and activate it to populate it with VSAP BMD data
        ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(j)
        ActiveWorkbook.Sheets(j + 1).Activate
    
    
        'Pulling file path for a specific file
        Dim nam As String
        nam = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)
    
        'importing text file as a query table
        With ActiveSheet.QueryTables.Add(Connection:= _
               "TEXT;" & nam _
               , destination:=Range("$A$1"))
               .Name = "Precinct " & j
               .FieldNames = True
               .RowNumbers = False
               .FillAdjacentFormulas = False
               .PreserveFormatting = True
               .RefreshOnFileOpen = False
               .RefreshStyle = xlInsertDeleteCells
               .SavePassword = False
               .SaveData = True
               .AdjustColumnWidth = True
               .RefreshPeriod = 0
               .TextFilePromptOnRefresh = False
               .TextFilePlatform = 437
               .TextFileStartRow = 1
               .TextFileParseType = xlDelimited
               .TextFileTextQualifier = xlTextQualifierDoubleQuote
               .TextFileConsecutiveDelimiter = False
               .TextFileTabDelimiter = False
               .TextFileSemicolonDelimiter = False
               .TextFileCommaDelimiter = False
               .TextFileSpaceDelimiter = False
               .TextFileOtherDelimiter = "|"
               .TextFileColumnDataTypes = Array(xlSkipColumn, xlGeneralFormat, xlSkipColumn, xlTextFormat, xlTextFormat, xlTextFormat, xlTextFormat)
               .TextFileTrailingMinusNumbers = True
               .Refresh BackgroundQuery:=False
        End With
    
        'Rename the Worksheet to the file name of the selected data file
        ActiveWorkbook.ActiveSheet.Name = "Precinct " & Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".log", ""), 10)
skipit:
    
    Next j
    
    'Deletes any blank sheets while more than one sheet is open
    d = ActiveWorkbook.Sheets.count
    For t = 1 To d
        If t <= d And t > 1 Then
            If IsEmpty(ActiveWorkbook.Sheets(t).Range("A1")) = True Then
                ActiveWorkbook.Worksheets(t).Delete
                d = ActiveWorkbook.Sheets.count
                t = 0
            End If
        End If
        d = ActiveWorkbook.Sheets.count
    Next t
    
    'Allow the Excel file to actively update
    Application.ScreenUpdating = True
    
    
    End Sub
Function parseDate(dateString As String) As Date
    Dim dateObj As Date
    Dim dateAndTime() As String
    dateAndTime = Split(dateString, "T")
    dateObj = dateAndTime(0) & " " & Split(dateAndTime(1), ".")(0)
    parseDate = dateObj
End Function
Function getTimeDifference(startTime As String, endTime As String) As String
    startLocal = parseDate(startTime)
    endlocal = parseDate(endTime)
    Dim mins As String
    Dim secs As String
    mins = Format(CStr(CLng(DateDiff("n", startLocal, endlocal))), "00")
    secs = Format(CStr(CLng(DateDiff("s", startLocal, endlocal)) Mod 60), "00")
    getTimeDifference = mins & ":" & secs
    'Debug.Assert (mins < 100)
End Function
Function getDifferenceMinutes(startTime As String, endTime As String) As Long
    startLocal = parseDate(startTime)
    endlocal = parseDate(endTime)
    getDifferenceMinutes = CLng(DateDiff("n", startLocal, endlocal))
End Function
Function Process_Single_VSAPBMD_Data_From_Stream(source As TextStream, destination As Object, startRow As Long, _
            writeHeader As Boolean, writeFileName As Boolean, fileName As String) As Long

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
        
        Dim writer As VSAPBMD_Writer
        Set writer = New VSAPBMD_Writer
        Set writer.destination = destination
        If writeFileName Then
            writer.fileName = fileName
        End If
                
        ' Timestamp (string) of the instant when current transaction was started
        Dim startTime As String
        startTime = ""
        Dim thisTime As String
        Dim j As Long
        Dim pollPassUsed As Boolean
        State = BMDState.INIT
        writer.row = startRow
        
        If writeHeader Then
            writer.writeHeader
        End If
        Dim i As Long
        Dim line As String
        i = 1
        Do While Not source.AtEndOfStream
                line = source.ReadLine
                elements = Split(line, "|")
                If UBound(elements) - LBound(elements) + 1 = 7 Then
                    thisTime = elements(1)
                    thisLog = elements(6)
                    If Not State = BMDState.INIT And Len(startTime) > 0 Then
                        If getDifferenceMinutes(startTime, thisTime) > 60 Then
                            ' A more than 60 minute difference probably indicates something suspicious.
                            ' So we reset state here.
                            State = BMDState.INIT
                        End If
                    End If
                    If State = BMDState.INIT Then
                        If Trim(thisLog) = loadingBallotLog Then
                            startTime = thisTime
                            State = BMDState.Loading
                            pollPassUsed = False
                        ElseIf Trim(thisLog) = removedBallotLog Then
                            State = BMDState.UnexpectedRemovedBallot
                        End If
                    ElseIf State = BMDState.UnexpectedRemovedBallot Then
                        If Trim(thisLog) = loadingBallotLog Then
                            ' Since this loading ballot log appears after an unexpected removed ballot log,
                            ' we will assume that this log line is the one that should have come before
                            ' the unexpected one we encountered before.
                            writer.writeBallotRemovedRecordNoTime
                            State = BMDState.INIT
                        End If
                    ElseIf State = BMDState.Loading Then
                        If Trim(thisLog) = loadingBallotLog Then
                            ' I wfe encounter another "loading" log at this state,
                            ' the first one most probably came after a mis-ordered one
                            writer.writeBallotRemovedRecordNoTime
                            pollPassUsed = False
                            startTime = thisTime
                        ElseIf Trim(thisLog) = removedBallotLog Then
                            ' This means the ballot was removed from the machine before it could
                            ' be read and activated. We need to record it and reset state.
                            writer.writeBallotRemovedRecord getTimeDifference(startTime, thisTime)
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = ballotActivatedLog Then
                            State = BMDState.Activated
                        ElseIf Trim(thisLog) = errorScanningBPMLog Then
                            writer.writeBPMScanErrorLog getTimeDifference(startTime, thisTime)
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = startLog Then
                            writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), False
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = quitVotingLog Then
                            writer.writeQuitVotingLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        End If
                    ElseIf State = BMDState.Activated Then
                        If Trim(thisLog) = pollPassScannedLog Then
                            pollPassUsed = True
                        ElseIf Trim(thisLog) = printedBallotLog Then
                            State = BMDState.printed
                        ElseIf Trim(thisLog) = castBallotLog Then ' Typo in "sucessfully" as it appears in logs
                            ' If the ballot was cast without being printed in this transaction.
                            ' This means a pre-printed ballot was inserted.
                            writer.writeBallotCastRecord getTimeDifference(startTime, thisTime), False, pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                            writer.writeProvisionalBallotEjectedRecord getTimeDifference(startTime, thisTime), False, pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = votingSessionLockedLog Then
                            writer.writeVotingTimedOutLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = quitVotingLog Then
                            writer.writeQuitVotingLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = startLog Then
                            writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = languageSelectedLog Then
                            writer.writeQuitVotingLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        End If
                    ElseIf State = BMDState.printed Then
                        If Trim(thisLog) = removedPrintedBallotLog Then
                            writer.writePrintedBallotRemovedRecord getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = castBallotLog Then
                            writer.writeBallotCastRecord getTimeDifference(startTime, thisTime), True, pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                            writer.writeProvisionalBallotEjectedRecord getTimeDifference(startTime, thisTime), True, pollPassUsed
                            State = BMDState.INIT
                        ElseIf Trim(thisLog) = startLog Then
                            writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), pollPassUsed
                            State = BMDState.INIT
                        End If
                        ' TODO Find out whether provisional ballots can be cast just after printing
                    End If
                End If
                i = i + 1
            Loop
        
        ' Return the next row (the one after the last written row)
        Process_Single_VSAPBMD_Data_From_Stream = writer.row
End Function
Function Process_Single_VSAPBMD_Data_From(source As Worksheet, destination As Object, startRow As Long, _
            writeHeader As Boolean, writeFileName As Boolean, fileName As String) As Long
        ' Count the number of rows in source sheet
        l2row = source.UsedRange.Rows.count
        
        Const loadingBallotLog As String = "Loading Ballot"
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
        
        Dim writer As VSAPBMD_Writer
        Set writer = New VSAPBMD_Writer
        Set writer.destination = destination
        If writeFileName Then
            writer.fileName = fileName
        End If
                
        ' Timestamp (string) of the instant when current transaction was started
        Dim startTime As String
        Dim thisTime As String
        Dim j As Long
        Dim pollPassUsed As Boolean
        State = BMDState.INIT
        writer.row = startRow
        
        If writeHeader Then
            writer.writeHeader
        End If
            
        With source
            ' Main loop
            For i = 1 To l2row
                thisLog = CStr(.Range("E" & i))
                thisTime = CStr(.Range("A" & i))
                If State = BMDState.INIT Then
                    If Trim(thisLog) = loadingBallotLog Then
                        startTime = thisTime
                        State = BMDState.Loading
                        pollPassUsed = False
                    ElseIf Trim(thisLog) = removedBallotLog Then
                        State = BMDState.UnexpectedRemovedBallot
                    End If
                ElseIf State = BMDState.UnexpectedRemovedBallot Then
                    If Trim(thisLog) = loadingBallotLog Then
                        ' Since this loading ballot log appears after an unexpected removed ballot log,
                        ' we will assume that this log line is the one that should have come before
                        ' the unexpected one we encountered before.
                        writer.writeBallotRemovedRecordNoTime
                        State = BMDState.INIT
                    End If
                ElseIf State = BMDState.Loading Then
                    If Trim(thisLog) = loadingBallotLog Then
                        ' If we encounter another "loading" log at this state,
                        ' the first one most probably came after a mis-ordered one
                        writer.writeBallotRemovedRecordNoTime
                        pollPassUsed = False
                        startTime = thisTime
                    ElseIf Trim(thisLog) = removedBallotLog Then
                        ' This means the ballot was removed from the machine before it could
                        ' be read and activated. We need to record it and reset state.
                        writer.writeBallotRemovedRecord getTimeDifference(startTime, thisTime)
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = ballotActivatedLog Then
                        State = BMDState.Activated
                    ElseIf Trim(thisLog) = errorScanningBPMLog Then
                        writer.writeBPMScanErrorLog getTimeDifference(startTime, thisTime)
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = startLog Then
                        writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), False
                        State = BMDState.INIT
                    End If
                ElseIf State = BMDState.Activated Then
                    If Trim(thisLog) = pollPassScannedLog Then
                        pollPassUsed = True
                    ElseIf Trim(thisLog) = printedBallotLog Then
                        State = BMDState.printed
                    ElseIf Trim(thisLog) = castBallotLog Then ' Typo in "sucessfully" as it appears in logs
                        ' If the ballot was cast without being printed in this transaction.
                        ' This means a pre-printed ballot was inserted.
                        writer.writeBallotCastRecord getTimeDifference(startTime, thisTime), False, pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                        writer.writeProvisionalBallotEjectedRecord getTimeDifference(startTime, thisTime), False, pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = votingSessionLockedLog Then
                        writer.writeVotingTimedOutLog getTimeDifference(startTime, thisTime), pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = quitVotingLog Then
                        writer.writeQuitVotingLog getTimeDifference(startTime, thisTime), pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = startLog Then
                        writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), pollPassUsed
                        State = BMDState.INIT
                    End If
                ElseIf State = BMDState.printed Then
                    If Trim(thisLog) = removedPrintedBallotLog Then
                        writer.writePrintedBallotRemovedRecord getTimeDifference(startTime, thisTime), pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = castBallotLog Then
                        writer.writeBallotCastRecord getTimeDifference(startTime, thisTime), True, pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = provisionalBallotEjectedLog Then
                        writer.writeProvisionalBallotEjectedRecord getTimeDifference(startTime, thisTime), True, pollPassUsed
                        State = BMDState.INIT
                    ElseIf Trim(thisLog) = startLog Then
                        writer.writeMachineRestartedLog getTimeDifference(startTime, thisTime), pollPassUsed
                        State = BMDState.INIT
                    End If
                    ' TODO Find out whether provisional ballots can be cast just after printing
                End If
            Next i
        End With
        
        ' Return the next row (the one after the last written row)
        Process_Single_VSAPBMD_Data_From = writer.row
End Function
Sub Process_VSAPBMD_Data_Single_To_Worksheet()
    
    Dim u As Long
    Dim lrow As Long
    Dim var As String
    Dim k As Long
    Dim Name As String
    Dim pctCom As Single
    
    'Displays the progress bar
    UserForm1.Show vbModeless
    
    'Updates the progress bar
    pctCom = 0
    progress pctCom
    
    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False
    
    'TODO Verify that this is a good test for VSAP BMDs.
    If Trim(Range("B1")) = "Logger.js-Loading page-Manual Diagnostic Status" Then
        Name = ActiveWorkbook.ActiveSheet.Name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).Name = Name & " Processed" Then
                Exit Sub
            End If
        Next n
      
      'Updates the progress bar
      pctCom = 1 / 4 * 100
      progress pctCom
      
        'Add a Worksheet in which processed precinct data will be populated
        ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
        
        'Name the created Worksheet to the name of the precinct selected with the "Processed" qualifier
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count).Name = Name & " Processed"
       
        'Copies the data from the current Worksheet to the newly created worksheet
        ActiveWorkbook.Sheets(Name).Activate
        lrow = Cells(ActiveWorkbook.ActiveSheet.Rows.count, 1).End(xlUp).row
        Range("A1", "E" & lrow).Copy
        ActiveWorkbook.Sheets(Name & " Processed").Activate
        Range("A1", "E" & lrow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
             :=True, Transpose:=False
       
       'Updates the progress bar
       pctCom = 2 / 4 * 100
       progress pctCom
        
        'Clears the information of the first row, which is an arbitrary log recording by the DS200
        Range("1:1").ClearContents
    
        
        'Updates the progress bar
        pctCom = 3 / 4 * 100
        progress pctCom
        
        l2row = Cells(Rows.count, 1).End(xlUp).row
        next_row = Process_Single_VSAPBMD_Data_From(ActiveWorkbook.ActiveSheet, ActiveWorkbook.ActiveSheet, 2, True, False, "")
        ActiveWorkbook.ActiveSheet.Columns("A:E").AutoFit
        ActiveWorkbook.ActiveSheet.Range("A" & next_row, "E" & l2row).Clear
        ActiveWorkbook.ActiveSheet.Range("E1", "E" & (next_row - 1)).Clear
    Else
        'If the file does not contain VSAP BMD Data, the program exits
        MsgBox "Action can not be done on this WorkSheet"
    End If
    
    'Begin refreshing the Excel document in real time
        Application.ScreenUpdating = True
        
    'Updates the progress bar
    pctCom = 4 / 4 * 100
    progress pctCom
    Unload UserForm1
    
    End Sub

Sub Process_VSAPBMD_Data_Single()
    
    Dim u As Long
    Dim lrow As Long
    Dim var As String
    Dim k As Long
    Dim Name As String

    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False
    
    'TODO Verify that this is a good test for VSAP BMDs.
    If Trim(Range("B1")) = "Logger.js-Loading page-Manual Diagnostic Status" Then
        Name = ActiveWorkbook.ActiveSheet.Name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).Name = Name & " Processed" Then
                Exit Sub
            End If
        Next n
        
        Dim filePath As String
        filePath = ActiveWorkbook.Path & "\output1.csv"
        
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(filePath)

        next_row = Process_Single_VSAPBMD_Data_From(ActiveWorkbook.ActiveSheet, fileStream, 2, True, False, "")
        fileStream.Close
    Else
        'If the file does not contain VSAP BMD Data, the program exits
        MsgBox "Action can not be done on this WorkSheet"
    End If
    
    'Begin refreshing the Excel document in real time
        Application.ScreenUpdating = True
    
    End Sub



