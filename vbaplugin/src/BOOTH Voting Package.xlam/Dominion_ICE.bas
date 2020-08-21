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
        ActiveWorkbook.ActiveSheet.name = parts(UBound(parts))
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
    Dim writer As SheetWriter
    Set writer = New SheetWriter
    writer.setOutputSheet sheet
    Do While Not inputStream.AtEndOfStream
        lineStr = inputStream.readLine
        Dim lineArr(2) As String
        lineArr(0) = Left(lineStr, 20) ' Timestamp is in the first 20 characters
        lineArr(1) = Mid(lineStr, 21)
        writer.OutputWriter_writeLineArr lineArr
    Loop
    inputStream.Close
End Sub

