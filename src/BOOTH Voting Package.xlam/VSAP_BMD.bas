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
               .TextFileColumnDataTypes = Array(xlTextFormat, xlGeneralFormat, xlTextFormat, xlTextFormat, xlTextFormat, xlTextFormat, xlTextFormat)
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
