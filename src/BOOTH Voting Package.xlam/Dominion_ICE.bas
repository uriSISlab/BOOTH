Attribute VB_Name = "Dominion_ICE"
' Log processing for Dominion ImageCast Evolution Ballot Scanning and Marking Device
Sub Import_DICE_data(control As IRibbonControl)
    
    'When File Explorer opens, only display text log files
    With Application.FileDialog(msoFileDialogFilePicker)
    Application.FileDialog(msoFileDialogFilePicker).Filters.Clear
    Application.FileDialog(msoFileDialogFilePicker).Filters.Add "Text files", "*.txt"
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
        
        Import_DICE_File_Into_Sheet filePath, ActiveWorkbook.ActiveSheet
        
        'Rename the Worksheet to the file name of the selected data file
        'TODO: check if name is already taken
        Dim parts() As String
        parts = Split(filePath, "\")
        ActiveWorkbook.ActiveSheet.Name = parts(UBound(parts))
    Next j
    
    'Allow the Excel file to actively update
    Application.ScreenUpdating = True
    
End Sub

Sub Import_DICE_File_Into_Sheet(filePath As String, sheet As Worksheet)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim inputStream As TextStream
    'Open the file as a text stream for reading
    Set inputStream = fso.OpenTextFile(filePath, ForReading, False)
    
    Dim lineStr, rest As String
    Dim timestamp As Date
    Dim col_pos As Integer
    Dim sheetWriter As OutputWriter
    Set sheetWriter = New OutputWriter
    sheetWriter.setOutputSheet sheet
    Do While Not inputStream.AtEndOfStream
        lineStr = inputStream.readLine
        Dim lineArr(2) As String
        lineArr(0) = Left(lineStr, 20) ' Timestamp is in the first 20 characters
        lineArr(1) = Mid(lineStr, 21)
        sheetWriter.writeLine lineArr
    Loop
    inputStream.Close
End Sub

Sub Process_DICE_Data_Single()
    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False

    If InStr(ActiveWorkbook.ActiveSheet.Cells(1, 2), "Logging service initialized") <> 0 Then
        Name = ActiveWorkbook.ActiveSheet.Name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).Name = Name & " Processed" Then
                Exit Sub
            End If
        Next n

        'Add a Worksheet in which processed precinct data will be populated
        ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
        
        'Name the created Worksheet to the name of the precinct selected with the "Processed" qualifier
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count).Name = Name & " Processed"
       
        'Copies the data from the current Worksheet to the newly created worksheet
        ActiveWorkbook.Sheets(Name & " Processed").Activate
        
        Dim processor As DICE_Processor
        Set processor = New DICE_Processor
        Dim writer As OutputWriter
        Set writer = New OutputWriter
        writer.setOutputSheet ActiveWorkbook.ActiveSheet
        processor.setWriter writer
        
        'Write the header
        Dim headerArr() As String
        headerArr = Split("Duration,Timestamp,Event,Misreads,Ballot Reviewed", ",")
        writer.writeLine headerArr
        
        Process_DICE_Data_From_Sheet ActiveWorkbook.Sheets(Name), processor
        ActiveWorkbook.ActiveSheet.Range("A1:E1").Font.Bold = True
        ActiveWorkbook.ActiveSheet.UsedRange.Columns.AutoFit
    Else
        'If the file does not contain VSAP BMD Data, the program exits
        MsgBox "Action can not be done on this WorkSheet"
    End If
    
    'Begin refreshing the Excel document in real time
    Application.ScreenUpdating = True
End Sub

Sub Process_DICE_Data_From_Sheet(sheet As Worksheet, processor As DICE_Processor)
    Dim rows As Long
    Dim line As String
    rows = sheet.UsedRange.rows.count
    
    For i = 1 To rows
        line = CStr(sheet.Range("A" & i).Text) & " " & CStr(sheet.Range("B" & i).Text)
        processor.readLine line
    Next i
End Sub

