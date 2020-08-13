Attribute VB_Name = "VSAP_BMD"

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
    
        'Add an additional sheet and activate it to populate it with VSAP BMD data
        ActiveWorkbook.Sheets.Add after:=ActiveSheet

        'Pulling file path for a specific file
        Dim filePath As String
        filePath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)
    
        'importing text file as a query table
        With ActiveSheet.QueryTables.Add(Connection:= _
               "TEXT;" & filePath _
               , destination:=Range("$A$1"))
               .name = "Precinct " & j
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
        'TODO: check if name is already taken
        Dim parts() As String
        parts = Split(filePath, "\")
        ActiveWorkbook.ActiveSheet.name = parts(UBound(parts))
skipit:
    
    Next j
    
    'Allow the Excel file to actively update
    Application.ScreenUpdating = True
    
    
    End Sub

Sub Process_Single_VSAPBMD_Data_From_Stream(source As TextStream, processor As VSAPBMD_Processor)
        Dim line As String
        Do While Not source.AtEndOfStream
            line = source.readLine
            processor.readLine line
        Loop
End Sub
Sub Process_Single_VSAPBMD_Data_From(source As Worksheet, processor As VSAPBMD_Processor)
        ' Count the number of rows in source sheet
        l2row = source.UsedRange.rows.count

        Dim i, j As Integer
        With source
            ' Main loop
            For i = 1 To l2row
                Dim line As String
                line = i & "|" & CStr(.Range("A" & i)) & "|placeholder"
                For j = 2 To 5
                    ' Join the row with pipes
                    line = line & "|" & .Range(getLetterFromNumber(j) & i)
                Next j
                processor.readLine line
            Next i
        End With
End Sub
Sub Process_VSAPBMD_Data_Single_To_Worksheet()
    
    Dim u As Long
    Dim lrow As Long
    Dim var As String
    Dim k As Long
    Dim name As String
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
        name = ActiveWorkbook.ActiveSheet.name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).name = name & " Processed" Then
                Exit Sub
            End If
        Next n
      
      'Updates the progress bar
      pctCom = 1 / 4 * 100
      progress pctCom
      
        'Add a Worksheet in which processed precinct data will be populated
        ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
        
        'Name the created Worksheet to the name of the precinct selected with the "Processed" qualifier
        ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count).name = name & " Processed"
       
        'Copies the data from the current Worksheet to the newly created worksheet
        ActiveWorkbook.Sheets(name).Activate
        lrow = Cells(ActiveWorkbook.ActiveSheet.rows.count, 1).End(xlUp).row
        Range("A1", "E" & lrow).Copy
        ActiveWorkbook.Sheets(name & " Processed").Activate
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
        
        l2row = Cells(rows.count, 1).End(xlUp).row
        Dim writer As OutputWriter
        Set writer = New OutputWriter
        Dim processor As VSAPBMD_Processor
        Set processor = New VSAPBMD_Processor
        writer.setOutputSheet ActiveWorkbook.ActiveSheet
        processor.setWriter writer
        processor.writeHeader
        
        Process_Single_VSAPBMD_Data_From ActiveWorkbook.ActiveSheet, processor
        
        next_row = writer.getRowNum
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
    Dim name As String

    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False
    
    'TODO Verify that this is a good test for VSAP BMDs.
    If Trim(Range("B1")) = "Logger.js-Loading page-Manual Diagnostic Status" Then
        name = ActiveWorkbook.ActiveSheet.name
        'Check if the data chosen was already processed
        For n = 1 To ActiveWorkbook.Sheets.count
            If ActiveWorkbook.Sheets(n).name = name & " Processed" Then
                Exit Sub
            End If
        Next n
        
        Dim filePath As String
        filePath = ActiveWorkbook.Path & "\output1.csv"
        
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(filePath)
        Dim writer As OutputWriter
        Set writer = New OutputWriter
        Dim processor As VSAPBMD_Processor
        Set processor = New VSAPBMD_Processor
        writer.setOutputStream fileStream
        processor.setWriter writer
        processor.writeHeader

        Process_Single_VSAPBMD_Data_From ActiveWorkbook.ActiveSheet, processor

        fileStream.Close
    Else
        'If the file does not contain VSAP BMD Data, the program exits
        MsgBox "Action can not be done on this WorkSheet"
    End If
    
    'Begin refreshing the Excel document in real time
        Application.ScreenUpdating = True
    
    End Sub



