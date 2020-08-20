Attribute VB_Name = "Dispatch"
Private Function addSheetForOutput(after As Worksheet) As Worksheet
    ActiveWorkbook.Worksheets.Add after:=after
    ActiveWorkbook.ActiveSheet.name = Util.getProcessedName(after.name)
    Set addSheetForOutput = ActiveWorkbook.ActiveSheet
End Function

Public Sub processSheetForLogType(sheet As Worksheet, t As LogType)
    Dim processor As LogProcessor
    Dim writer As SheetWriter
    Dim reader As SheetReader
    Set processor = Util.createProcessor(t)
    Set writer = New SheetWriter
    Set reader = New SheetReader
    
    'Check if the data chosen was already processed
    For n = 1 To ActiveWorkbook.Sheets.count
        If ActiveWorkbook.Sheets(n).name = Util.getProcessedName(sheet.name) Then
            Exit Sub
        End If
    Next n
  
    reader.setSheetAndSeparator sheet, processor.getSeparator
    writer.setOutputSheet addSheetForOutput(sheet)
    Util.runPipeline reader, processor, writer, True
    
    writer.formatPretty
End Sub

Function CountFiles(glob As String) As Long
    Dim fileName As String
    Dim count As Long
    count = 0
    fileName = Dir(glob)
    Do While Len(fileName) > 0
        count = count + 1
        fileName = Dir
    Loop
    CountFiles = count
End Function

Public Sub openAndProcessDirectory(control As IRibbonControl)
    Dim t As LogType
    If control.ID = "VSAPBMD_folder" Then
        t = LogType.VSAP_BMD
    ElseIf control.ID = "DICE_folder" Then
        t = LogType.DICE
    ElseIf control.ID = "DICX_folder" Then
        t = LogType.DICX
    Else
        MsgBox "This feature has not been implemented yet."
        Exit Sub
    End If
    processEntireDirectory t
End Sub

Public Sub processEntireDirectory(t As LogType)
    Dim folder As String
    ' Create folder picker
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            folder = CStr(.SelectedItems(1))
        Else
            Exit Sub
        End If
    End With
    
    'Prevent showing Excel document updates to improve performance
    Application.ScreenUpdating = False
    
    Dim outputRow As Long
    Dim fileCount As Long
    Dim fileNum As Long
    Dim outputFileName As String
    fileNum = 1
    outputRow = 2
    outputFileName = folder & "\processed_all.csv"
    
    Dim fileName As String
    Dim filesGlob As String
    filesGlob = folder & "\" & Util.getFileNamePatternForLog(t)
    fileCount = CountFiles(filesGlob)
    fileName = Dir(filesGlob)

    ' Show progress bar
    UserForm1.Show vbModeless
    progress 0
    
    Dim writer As FileWriter
    Set writer = New FileWriter
    writer.setFilePath outputFileName
    
    'Loop to process multiple files consecutively
    Do While Len(fileName) > 0
        Dim processor As LogProcessor
        Set processor = Util.createProcessor(t)
        Dim reader As FileReader
        Set reader = New FileReader
        reader.setFilePath folder & "\" & fileName
        processor.setFilename fileName

        Util.runPipeline reader, processor, writer, (outputRow = 2)
        
        outputRow = writer.OutputWriter_getRowNum

        fileName = Dir
        fileNum = fileNum + 1
        If fileNum Mod 5 = 0 Then
            progress fileNum / fileCount * 100
        End If
    Loop

    writer.OutputWriter_done
        
    ' Stop showing progress bar
    Unload UserForm1
    
    MsgBox "Processed output saved to " & outputFileName
    
    'Allow the Excel file to actively update
    Application.ScreenUpdating = True

End Sub
