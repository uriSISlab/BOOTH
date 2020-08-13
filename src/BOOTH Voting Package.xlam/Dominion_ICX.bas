Attribute VB_Name = "Dominion_ICX"
' Log processing for Dominion ImageCast X Ballot Scanning and Marking Device
Sub Import_DICX_data(control As IRibbonControl)
    
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
    
        'Add an additional sheet and activate it to populate it with Dominion ICE data
        ActiveWorkbook.Sheets.Add after:=ActiveSheet

        'Pulling file path for a specific file
        Dim filePath As String
        filePath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)
        
        Import_DICX_File_Into_Sheet filePath, ActiveWorkbook.ActiveSheet
        
        'Rename the Worksheet to the file name of the selected data file
        'TODO: check if name is already taken
        Dim parts() As String
        parts = Split(filePath, "\")
        ActiveWorkbook.ActiveSheet.name = parts(UBound(parts))
    Next j
    
    'Allow the Excel file to actively update
    Application.ScreenUpdating = True
    
End Sub

Sub Import_DICX_File_Into_Sheet(filePath As String, sheet As Worksheet)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim inputStream As TextStream
    'Open the file as a text stream for reading
    Set inputStream = fso.OpenTextFile(filePath, ForReading, False)
    
    Dim lineStr, rest As String
    Dim sheetWriter As OutputWriter
    Set sheetWriter = New OutputWriter
    sheetWriter.setOutputSheet sheet
    Do While Not inputStream.AtEndOfStream
        lineStr = inputStream.readLine
        Dim lineArr(2) As String
        lineArr(0) = Left(lineStr, 19) ' Timestamp is in the first 19 characters
        lineArr(1) = Mid(lineStr, 23) ' Next three characters are " - ", so the rest of the line starts from 23.
        sheetWriter.writeLine lineArr
    Loop
    inputStream.Close
End Sub

Public Function is_DICX_Log(sheet As Worksheet) As Boolean
    Dim idRange As Range
    Set idRange = sheet.UsedRange.Find(What:="Audit Log file is saved.", _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
    is_DICX_Log = Not idRange Is Nothing
End Function

Private Function get_clipped_processed_name(name As String) As String
    Dim processedName As String
    processedName = name & " Processed"
    If Len(processedName) > 31 Then
        processedName = Right(processedName, 31)
    End If
    get_clipped_processed_name = processedName
End Function


Sub Process_DICX_Data_Single()
    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False

    If is_DICX_Log(ActiveWorkbook.ActiveSheet) Then
        Dim name As String
        name = ActiveWorkbook.ActiveSheet.name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).name = get_clipped_processed_name(name) Then
                Exit Sub
            End If
        Next n

        'Add a Worksheet in which processed precinct data will be populated
        ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
        
        'Name the created Worksheet to the name of the precinct selected with the "Processed" qualifier
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count).name = get_clipped_processed_name(name)
       
        'Copies the data from the current Worksheet to the newly created worksheet
        ActiveWorkbook.Sheets(get_clipped_processed_name(name)).Activate
        
        Dim processor As DICX_Processor
        Set processor = New DICX_Processor
        Dim writer As OutputWriter
        Set writer = New OutputWriter
        writer.setOutputSheet ActiveWorkbook.ActiveSheet
        processor.setWriter writer
        
        'Write the header
        Dim headerArr() As String
        headerArr = Split("Duration,Timestamp,Event", ",")
        writer.writeLine headerArr
        
        Process_DICX_Data_From_Sheet ActiveWorkbook.Sheets(name), processor
        ActiveWorkbook.ActiveSheet.Range("A1:C1").Font.Bold = True
        ActiveWorkbook.ActiveSheet.UsedRange.Columns.AutoFit
    Else
        'If the file does not contain VSAP BMD Data, the program exits
        MsgBox "Action can not be done on this WorkSheet"
    End If
    
    'Begin refreshing the Excel document in real time
    Application.ScreenUpdating = True
End Sub

Sub Process_DICX_Data_From_Sheet(sheet As Worksheet, processor As DICX_Processor)
    Dim rows As Long
    Dim line As String
    rows = sheet.UsedRange.rows.count
    
    For i = 1 To rows
        line = CStr(sheet.Range("A" & i).Text) & " - " & CStr(sheet.Range("B" & i).Text)
        processor.readLine line
    Next i
End Sub


