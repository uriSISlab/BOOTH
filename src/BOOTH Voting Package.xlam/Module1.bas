Attribute VB_Name = "Module1"
Sub Import_DS200_data(control As IRibbonControl)

Dim lrow As Long
Dim l2row As Long
Dim l3row As Long
Dim i As Integer
Dim t As Long
Dim ret1 As String
Dim j As Integer
Dim intResult As Integer
Dim strPath As String
Dim arraylen As Integer
Dim tbook As ThisWorkbook
Dim f As Integer
Dim w As Long


'When File Explorer opens, only display text files
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
For j = 1 To Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count

    'Adds an additional Worksheet to write DS200 data to if only one sheet is open
    If ActiveWorkbook.Sheets.Count = 1 Then
        ActiveWorkbook.Sheets.Add after:=ActiveSheet
    End If

    'Check for duplicate precincts and delete the duplicate sheets
    c = 1
    While c < ActiveWorkbook.Sheets.Count + 1
        If ActiveWorkbook.Sheets(c).Name = "Precinct " & Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".txt", ""), 10) Then
            GoTo skipit
        Else: c = c + 1
        End If
    Wend

    'Add an additional sheet and activate it to populate it with DS200 data
    ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(j)
    ActiveWorkbook.Sheets(j + 1).Activate


    'Pulling file path for a specific file
    Dim nam As String
    nam = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)

    'importing text file as a query table
    With ActiveSheet.QueryTables.Add(Connection:= _
           "TEXT;" & nam _
           , Destination:=Range("$A$1"))
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
           .TextFileTabDelimiter = True
           .TextFileSemicolonDelimiter = False
           .TextFileCommaDelimiter = True
           .TextFileSpaceDelimiter = False
           .TextFileColumnDataTypes = Array(1, 9, 2, 9, 9, 2, 2)
           .TextFileTrailingMinusNumbers = True
           .Refresh BackgroundQuery:=False
    End With

    'Rename the Worksheet to the file name of the selected data file
    ActiveWorkbook.ActiveSheet.Name = "Precinct " & Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".txt", ""), 10)
skipit:

Next j

'Deletes any blank sheets while more than one sheet is open
d = ActiveWorkbook.Sheets.Count
For t = 1 To d
    If t <= d And t > 1 Then
        If IsEmpty(ActiveWorkbook.Sheets(t).Range("A1")) = True Then
            ActiveWorkbook.Worksheets(t).Delete
            d = ActiveWorkbook.Sheets.Count
            t = 0
        End If
    End If
    d = ActiveWorkbook.Sheets.Count
Next t

'Allow the Excel file to actively update
Application.ScreenUpdating = True

End Sub


Sub Process_DS200_Data_Single()

Dim u As Long
Dim lrow As Long
Dim var As String
Dim k As Long
Dim Name As String
Dim PCTCom As Single

'Displays the progress bar
UserForm1.Show vbModeless

'Updates the progress bar
PCTCom = 0
progress PCTCom

'Prevent showing Excel document updates to improve performance
Application.ScreenUpdating = False

'Checks that current sheet has raw DS200 data
If Range("A1") = 1114111 Then
    Name = ActiveWorkbook.ActiveSheet.Name
    'Check if the data chosen was already processed
    For n = 1 To ActiveWorkbook.Sheets.Count
        If ActiveWorkbook.Sheets(n).Name = Name & " Processed" Then
            Exit Sub
        End If
    Next n
  
  'Updates the progress bar
  PCTCom = 1 / 4 * 100
  progress PCTCom
  
    'Add a Worksheet in which processed precinct data will be populated
    ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    
    'Name the created Worksheet to the name of the precinct selected with the "Processed" qualifier
    ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count).Name = Name & " Processed"
  
    'Defining loop variables
    u = 2
   
    'Copies the data from the current Worksheet to the newly created worksheet
    ActiveWorkbook.Sheets(Name).Activate
    lrow = Cells(ActiveWorkbook.ActiveSheet.Rows.Count, 1).End(xlUp).Row
    Range("A1", "E" & lrow).Copy
    ActiveWorkbook.Sheets(Name & " Processed").Activate
    Range("A1", "E" & lrow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         :=True, Transpose:=False
   
   'Updates the progress bar
   PCTCom = 2 / 4 * 100
   progress PCTCom
    'Deletes rows not pertaining to scanning a ballot. Keeps only the start scan, stop scan, and error code lines
    For i = 1 To lrow - 1
         If Range("A" & lrow) <> 1004115 And Range("A" & lrow) <> 1004163 And Range("A" & lrow) <> 3013006 And Range("A" & lrow) <> 1004138 And Range("A" & lrow) <> 1004016 And Range("A" & lrow) <> 1004056 And Range("A" & lrow) <> 1004022 And Range("A" & lrow) <> 1004111 And Range("A" & lrow) <> 1004113 And Range("A" & lrow) <> 3013005 And Range("A" & lrow) <> 3003337 And Range("A" & lrow) <> 3013001 And Range("A" & lrow) <> 3013004 And Range("A" & lrow) <> 3013008 And Range("A" & lrow) <> 3013002 And Range("A" & lrow) <> 7003009 And Range("A" & lrow) <> 3013003 And Range("A" & lrow) <> 3013007 And Range("A" & lrow) <> 3013009 And Range("A" & lrow) <> 3003335 And Range("A" & lrow) <> 3003336 And Range("A" & lrow) <> 3003339 And Range("A" & lrow) <> 3003340 And Range("A" & lrow) <> 3003318 And Range("A" & lrow) <> 3003341 And Range("A" & lrow) <> 1004122 And Range("A" & lrow) <> 1004112 And Range("A" & lrow) <> 1004114 And Range("A" & lrow) <> 1004328 Then
            Range("A" & lrow).EntireRow.Delete
            lrow = lrow - 1
         Else: lrow = lrow - 1
         End If
    Next i
    
    'Clears the information of the first row, which is an arbitrary log recording by the DS200
    Range("1:1").ClearContents
    
    ' Recount number of rows for next loop
    l2row = Cells(Rows.Count, 1).End(xlUp).Row
 
    ' Trims the space that precedes the time stamp to allow for mathematical operations
    If Left(Range("B2").Value, 1) = " " Then
        For i = 1 To l2row
            var = LTrim(Range("B" & i).Value)
            Range("B" & i) = var
        Next

    End If
    
    'Updates the progress bar
    PCTCom = 3 / 4 * 100
    progress PCTCom

    'Refreshes the cells which were trimmed in order to set them in a number format capable of mathmatical operations
    Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

With ActiveWorkbook.ActiveSheet
   'Calculates the duration of an instance given a set of parameters
   For i = 2 To l2row
        'Calculates scan durations and records event descriptions
        If .Range("A" & i) = 1004115 And .Range("A" & (i + 1)) <> 1004113 And .Range("A" & (i + 1)) <> 1004111 Then
            .Range("E" & i) = .Range("B" & (i + 1)) * 1 - .Range("B" & i) * 1
            .Range("H" & i) = .Range("E" & i) * 86400
            If .Range("A" & (i + 1)) <> 1004022 Then
                .Range("G" & i) = "Unsuccessful"
                .Range("F" & i) = .Range("D" & (i + 1))
            Else
                .Range("G" & i) = "Successful"
                .Range("F" & i) = "No Error"
            End If
                i = i + 1
        Else
            If .Range("A" & i) = 1004115 And .Range("A" & (i + 2)) = 1004022 Then
                .Range("E" & i) = .Range("B" & (i + 2)) * 1 - .Range("B" & i) * 1
                .Range("H" & i) = .Range("E" & i) * 86400
                .Range("G" & i) = "Successful"
                .Range("F" & i) = .Range("D" & (i + 1))
                i = i + 2
            Else
                If .Range("A" & i) = 3013004 And .Range("A" & (i - 1)) <> 1004115 And .Range("A" & (i + 1)) = 1004328 Then
                    .Range("E" & i) = .Range("B" & (i + 1)) * 1 - .Range("B" & i) * 1
                    .Range("H" & i) = .Range("E" & i) * 86400
                    .Range("G" & i) = "Jam"
                    .Range("F" & i) = .Range("D" & i)
                    i = i + 1
                Else
                    If .Range("A" & i) = 1004016 And .Range("A" & (i - 1)) = 1004163 And .Range("A" & (i + 1)) = 1004056 Then
                        .Range("E" & i) = .Range("B" & (i + 1)) * 1 - .Range("B" & i) * 1
                        .Range("H" & i) = .Range("E" & i) * 86400
                        .Range("G" & i) = "Shutdown"
                        .Range("F" & i) = Range("D" & i)
                        i = i + 1
                    End If
                End If
            End If
        End If
    Next i

    'Clear irrelavent columns
    .Range("A1", "D" & l2row).EntireColumn.Delete

    'Deletes the rows that are empty
    While u <= l2row
       If IsEmpty(.Range("A" & u)) = True Then
            .Range("A" & u).EntireRow.Delete
            u = u
            l2row = l2row - 1
       Else: u = u + 1
       End If
    Wend

    'Formats the column headers and enters their titles
    .Range("A:A").NumberFormat = "mm:ss"
    .Range("A1") = "Duration (mm:ss)"
    .Range("B1") = "Scan Type"
    .Range("C1") = "Ballot Cast Status"
    .Range("D1") = "Simio Input (seconds)"
    .Range("D:D").NumberFormat = "general"
    .Columns("A:D").AutoFit
    .Range("A1", "D1").Font.Bold = True
    .Range("A1", "C1").HorizontalAlignment = xlCenter
    'Delete any stray data
    .Range("E1", "K" & l2row).ClearContents
End With
   
Else
    'If the file does not contain raw DS200 data, the program exits
    MsgBox "Action can not be done on this WorkSheet"
    
End If

'Begin refreshing the Excel document in real time
    Application.ScreenUpdating = True
    
'Updates the progress bar
PCTCom = 4 / 4 * 100
progress PCTCom
Unload UserForm1

  End Sub

Sub Process_DS200_Data_Multiple(control As IRibbonControl)

Dim u As Long
Dim lrow As Long
Dim var As String
Dim k As Long
Dim Name As String
Dim PCTCom As Single
Dim TWS As Integer
Dim y As Long

'Variable to determine progress
TWS = ActiveWorkbook.Sheets.Count

'Shows the progress bar
UserForm1.Show vbModeless

'Prevent showing Excel document updates to improve performance
Application.ScreenUpdating = False

'Loops for every open Worksheet
For y = 1 To ActiveWorkbook.Sheets.Count

    ActiveWorkbook.Sheets(y).Activate
  
'Tests if the worksheet selected is in the expected format and whether it is a PollPad file or a DS200 file
If ActiveWorkbook.ActiveSheet.Cells(1, 1).NumberFormat = "General" And ActiveWorkbook.ActiveSheet.Cells(1, 2).NumberFormat = "@" And ActiveWorkbook.ActiveSheet.Cells(1, 3).NumberFormat = "@" And ActiveWorkbook.ActiveSheet.Cells(1, 4).NumberFormat = "@" Then
    'If the worksheet is identified as a DS200 data, the DS200 processing function is called
    Call Process_DS200_Data_Single
Else
    If ActiveWorkbook.ActiveSheet.Cells(2, 1).NumberFormat = "General" And ActiveWorkbook.ActiveSheet.Cells(2, 2).NumberFormat = "m/d/yyyy h:mm" And ActiveWorkbook.ActiveSheet.Cells(2, 3).NumberFormat = "General" Then
        'If the worksheet is identified as PollPad data, the PollPad processing function is called
        Call PollPadProcessing
    Else
        'If the sheet is already processed or blank, it is skipped
        If WorksheetFunction.CountA(ActiveWorkbook.ActiveSheet.UsedRange) = 0 Or Right(ActiveWorkbook.ActiveSheet.Name, 9) = "Processed" Or ActiveWorkbook.ActiveSheet.Cells(2, 3).NumberFormat = "h:mm" Then
            GoTo ExitIf
        Else
            'Indicating which file(s) has issues
            MsgBox ("The sheet: " & ActiveWorkbook.Sheets(y).Name & " does not contain compatible data.")
            GoTo ExitIf
        End If
    End If
End If

ExitIf:
       

Next y

'Begin updating the Excel document again
Application.ScreenUpdating = True

'Updates the progress bar
PCTCom = 100
progress PCTCom

Unload UserForm1
MsgBox "All Worksheets Have Been Processed."

End Sub

Sub progress(PCTCom As Single)

'Progress bar function
UserForm1.Text.Caption = Round(PCTCom, 0) & "% Completed"
UserForm1.Bar.Width = Round(PCTCom * 2, 0)

DoEvents

End Sub
Sub PollPadImport(control As IRibbonControl)

Dim lrow As Long
Dim l2row As Long
Dim l3row As Long
Dim i As Integer
Dim t As Long
Dim ret1 As String
Dim j As Integer
Dim intResult As Integer
Dim strPath As String
Dim arraylen As Integer
Dim tbook As ThisWorkbook
Dim f As Integer
Dim w As Long
Dim FileNam As String
Dim Name As String


'When File Explorer opens, only display text files
With Application.FileDialog(msoFileDialogFilePicker)
Application.FileDialog(msoFileDialogFilePicker).Filters.Clear
Application.FileDialog(msoFileDialogFilePicker).Filters.Add "PollPad Files", "*.csv, *.txt"
End With

'Open the file explorer and allow the selection of multiple files
Application.FileDialog(msoFileDialogFilePicker).Show
Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = True

'Prevent showing Excel document updates to improve performance
Application.ScreenUpdating = False

'Loop to process multiple files consecutively
For j = 1 To Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count
 FileNam = Left(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), 10)
 'Check for duplicate precincts and delete the duplicate sheets
    c = 1
    While c < ActiveWorkbook.Sheets.Count + 1
        If ActiveWorkbook.Sheets(c).Name = Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".csv", ""), 10) & " PollPad" Then
                MsgBox (FileNam & " shares the first 10 characters with a current worksheet. Please rename the file and import again.")
                GoTo skipit
        Else:
            c = c + 1
        End If
    Wend
     'Add an additional sheet and activate it to populate it with DS200 data
    ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(j)
    ActiveWorkbook.Sheets(j + 1).Activate
 
    'Names the worksheet after the file name
    ActiveWorkbook.ActiveSheet.Name = Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".csv", ""), 10) & " PollPad"
    
    'Pulling file path for a specific file
    Dim nam As String
    nam = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)

    'importing text file as a query table
    With ActiveSheet.QueryTables.Add(Connection:= _
           "TEXT;" & nam _
           , Destination:=Range("$A$1"))
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
           .TextFileTabDelimiter = True
           .TextFileSemicolonDelimiter = False
           .TextFileCommaDelimiter = True
           .TextFileSpaceDelimiter = False
          
           .TextFileTrailingMinusNumbers = True
           .Refresh BackgroundQuery:=False
    End With
    
skipit:

Next j

'Allow the Excel file to actively update
Application.ScreenUpdating = True


End Sub

Sub PollPadProcessing()


Application.ScreenUpdating = False

ColNum = ActiveWorkbook.ActiveSheet.UsedRange.Columns.Count

'Loops through worksheet to format data, separating date and time
For i = 1 To ColNum

If ActiveWorkbook.ActiveSheet.Cells(2, i).NumberFormat = "m/d/yyyy h:mm" Then
With ActiveWorkbook.ActiveSheet
    .Columns(i + 1).Insert
    .Columns(i).Copy ActiveWorkbook.ActiveSheet.Columns(i + 1)
    .Columns(i + 1).NumberFormat = "h:mm"
    .Cells(1, i + 1) = "Time"
    .Columns(i + 1).Insert
    .Columns(i).Copy ActiveWorkbook.ActiveSheet.Columns(i + 1)
    .Columns(i + 1).NumberFormat = "m/d/yyyy"
    .Cells(1, i + 1) = "Date"
End With
'Breaks free of the loop if completed
GoTo endloophere
Else:
End If

Next i

endloophere:

Application.ScreenUpdating = True



End Sub

Sub TestDataSet(control As IRibbonControl)

'Tests which type of data is in the worksheet in order to call the appropriate single sheet processing function
If ActiveWorkbook.ActiveSheet.Cells(1, 1).NumberFormat = "General" And ActiveWorkbook.ActiveSheet.Cells(1, 2).NumberFormat = "@" And ActiveWorkbook.ActiveSheet.Cells(1, 3).NumberFormat = "@" And ActiveWorkbook.ActiveSheet.Cells(1, 4).NumberFormat = "@" Then
    Call Process_DS200_Data_Single
Else
    If ActiveWorkbook.ActiveSheet.Cells(2, 1).NumberFormat = "General" And ActiveWorkbook.ActiveSheet.Cells(2, 2).NumberFormat = "m/d/yyyy h:mm" And ActiveWorkbook.ActiveSheet.Cells(2, 3).NumberFormat = "General" Then
        Call PollPadProcessing
    Else
        'Provides error message when incompatible data is selected
        If WorksheetFunction.CountA(ActiveWorkbook.ActiveSheet.UsedRange) = 0 Or Right(ActiveWorkbook.ActiveSheet.Name, 9) = "Processed" Or ActiveWorkbook.ActiveSheet.Cells(2, 3).NumberFormat = "h:mm" Or ActiveWorkbook.ActiveSheet.Cells(2, 4).NumberFormat = "h:mm" Then
            GoTo ExitIf
        Else
            MsgBox ("The sheet: " & ActiveWorkbook.ActiveSheet.Name & " does not contain compatible data.")
            GoTo ExitIf
        End If
    End If
End If

ExitIf:

End Sub
Sub TestforStat(control As IRibbonControl)

'Identifies data type for statistical functions to be called
If ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Duration (mm:ss)" And ActiveWorkbook.ActiveSheet.Cells(1, 2) = "Scan Type" Then
    Call DSStatTable
Else
    If ActiveWorkbook.ActiveSheet.Cells(2, 1).NumberFormat = "General" And ActiveWorkbook.ActiveSheet.Cells(2, 2).NumberFormat = "m/d/yyyy h:mm" And ActiveWorkbook.ActiveSheet.Cells(2, 3).NumberFormat = "m/d/yyyy" And ActiveWorkbook.ActiveSheet.Cells(2, 4).NumberFormat = "h:mm" Then
        Call pivottablepollpad
    'Provides error message when incompatible data is selected
    Else: MsgBox ("The sheet: " & ActiveWorkbook.ActiveSheet.Name & " does not contain compatible data.")
        GoTo ExitIf
    End If
End If

ExitIf:

End Sub

Sub pivottablepollpad()
Dim Early As Date
Dim Late As Date

Application.ScreenUpdating = False

'Storing shortened file names
FirstName = ActiveWorkbook.ActiveSheet.Name
SecondName = Left(ActiveWorkbook.ActiveSheet.Name, 10) + " PrecinctTurnout"
ThirdName = Left(ActiveWorkbook.ActiveSheet.Name, 10) + " TotalTurnout"
RawRows = ActiveWorkbook.ActiveSheet.Cells(ActiveWorkbook.ActiveSheet.Rows.Count, 1).End(xlUp).Row
  
  
i = 0

'Tests to see if sheet name is already taken
For y = 1 To ActiveWorkbook.Sheets.Count
If SecondName = ActiveWorkbook.Sheets(y).Name Then
MsgBox ("Sheet name already taken for precinct turnout, please rename the sheet.")
i = 1
GoTo CheckName1
Else
End If
Next y

'Filters the PollPad data by time in ascending order
If ActiveWorkbook.ActiveSheet.AutoFilterMode = True Then
ActiveWorkbook.ActiveSheet.AutoFilterMode = False
Else
End If
ActiveWorkbook.ActiveSheet.Range("C1").Select
    Selection.AutoFilter
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key _
        :=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sets the starting point of the dayan hour before the first observation
    Early = DateAdd("h", -1, WorksheetFunction.MRound(ActiveWorkbook.ActiveSheet.Range("D2"), "1:00"))
    Early2 = DateAdd("h", -1, WorksheetFunction.MRound(ActiveWorkbook.ActiveSheet.Range("D2"), "1:00"))
     
    'Filters the data by time in descending order
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key _
        :=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sets the end time an hour after the latest observation
   Late = WorksheetFunction.MRound(ActiveWorkbook.ActiveSheet.Range("D2"), "1:00")
   
'Adds new worksheet and names it after the file
ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.ActiveSheet
ActiveWorkbook.ActiveSheet.Name = SecondName

'Pulls precinct number and name from the data sheet to the stats sheet
ActiveWorkbook.Sheets(FirstName).Range("F:F").Copy ActiveWorkbook.ActiveSheet.Range("BD1")
ActiveSheet.Range("BD:BD").RemoveDuplicates Columns:=1, Header:=xlYes
rowscount = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 56).End(xlUp).Row
Range("BD1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "BD2:BD" & rowscount), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("BD1:BD" & rowscount)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
ActiveWorkbook.Sheets("" & FirstName & "").Range("F:F").Copy ActiveWorkbook.Sheets("" & SecondName & "").Range("BF1")
ActiveWorkbook.Sheets("" & FirstName & "").Range("E:E").Copy ActiveWorkbook.Sheets("" & SecondName & "").Range("BG1")
ActiveSheet.Range("BF:BG").RemoveDuplicates Columns:=1, Header:=xlYes
  
'Creates a table out of the precinct numbers for the dropdown menus
ActiveSheet.ListObjects.Add(xlSrcRange, Range("$BD$1:$BD$" & rowscount), , xlYes).Name _
        = "Table23"
       Range("B2").Select
    'Creates the dropdown menus
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$BD$2:$BD$" & rowscount & ""
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B2").Value = Range("BD2").Value
    
    Range("G2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$BD$2:$BD$" & rowscount & ""
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("G2").Value = Range("BD2").Value

'Loops to calculate the counts and percentages of voter turnout
t = 5
While Early <= Late
With ActiveWorkbook.ActiveSheet
    .Range("B" & t) = Early
    .Range("C" & t).Formula = "=COUNTIFS('" & FirstName & "'!C6,""=""&'" & SecondName & "'!R2C2,'" & FirstName & "'!C4,"">=""&'" & SecondName & "'!RC[-1],'" & FirstName & "'!C4,""<""&'" & SecondName & "'!R[1]C[-1])"
    .Range("D" & t) = "=C" & t & "/E2*100"
    .Range("E" & t) = "=C" & t & "/E2"

    .Range("G" & t).Formula = "=COUNTIFS('" & FirstName & "'!C6,""=""&'" & SecondName & "'!R2C7,'" & FirstName & "'!C4,"">=""&'" & SecondName & "'!RC2,'" & FirstName & "'!C4,""<""&'" & SecondName & "'!R[1]C2)"
    .Range("H" & t) = "=G" & t & "/J2*100"
    .Range("I" & t) = "=G" & t & "/J2"
End With
Early = DateAdd("h", 1, Early)
t = t + 1
Wend

'Formatting of worksheet and aesthetics
With ActiveWorkbook.ActiveSheet
    .Range("BD1") = "Precinct List"
    .Range("BD1").Font.Bold = True
    .Range("A2").Font.Bold = True
    .Range("D2").Font.Bold = True
    .Range("B4").Font.Bold = True
    .Range("C4").Font.Bold = True
    .Range("D4").Font.Bold = True
    .Range("E4").Font.Bold = True
    
    .Range("F2").Font.Bold = True
    .Range("G4").Font.Bold = True
    .Range("H4").Font.Bold = True
    .Range("I4").Font.Bold = True
    .Range("I2").Font.Bold = True
    
    .Range("A2") = "Select Precint Number From List:"
    .Range("F2") = "Select Precint Number From List:"
    .Range("D2") = "Total Count:"
    .Range("E2") = "=sum(C5:C" & t - 1 & ")"
    .Range("I2") = "Total Count:"
    .Range("J2") = "=sum(G5:G" & t - 1 & ")"
    .Range("B4") = "Time"
    .Range("C4") = "Precinct 1 Count"
    .Range("G4") = "Precinct 2 Count"
    .Range("D4") = "Percent"
    .Range("E4") = "Simio Input"
    .Range("H4") = "Percent"
    .Range("I4") = "Simio Input"
    .Range("B:B").NumberFormat = "h:mm AM/PM"
    .Columns("A:I").AutoFit
    .Range("A3") = "=VLOOKUP(B2,BF:BG,2,FALSE)"
    .Range("F3") = "=VLOOKUP(G2,BF:BG,2,FALSE)"
    
    .Range("A1").EntireRow.Insert
    .Range("C1") = "Precinct Specific"
    .Range("A1:J1").Merge
    .Range("A1").HorizontalAlignment = xlCenter
    .Range("A1").Font.Bold = True
    .Range("A2:E2").Merge
    .Range("F2:J2").Merge
    .Range("A2") = "First Precinct (Green)"
    .Range("F2") = "Second Precinct (Orange)"
    .Range("A2").HorizontalAlignment = xlCenter
    .Range("G2").HorizontalAlignment = xlCenter
    .Range("A2").Font.Bold = True
    .Range("F2").Font.Bold = True
    .Range("B4:C" & t - 1).Select
End With

'Creates figures of hourly voter turnout by count and percent per hour
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
With ActiveChart
    .SetSourceData Source:=ActiveWorkbook.ActiveSheet.Range("$B$5:$C$" & t)
    .ChartTitle.Caption = "='" & SecondName & "'!$A$4"
    .Parent.Top = -100
    .Parent.Left = 810
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Count of Voters"
    .SeriesCollection.Add _
        Source:=ActiveWorkbook.ActiveSheet.Range("$G$6:$G$" & t)
    .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 128, 0)
    .SeriesCollection(1).MarkerForegroundColor = RGB(0, 128, 0)
    .SeriesCollection(1).MarkerBackgroundColor = RGB(0, 128, 0)
End With

ActiveWorkbook.ActiveSheet.Range("B4:B" & t - 1 & ",D4:D" & t - 1).Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
With ActiveChart
    .SetSourceData Source:=ActiveWorkbook.ActiveSheet.Range("B5:B" & t & ",D5:D" & t)
    .ChartTitle.Caption = "='" & SecondName & "'!$A$4"
    .Parent.Top = 215
    .Parent.Left = 810
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percent of Voters"
    .SeriesCollection.Add _
        Source:=ActiveWorkbook.ActiveSheet.Range("$H$6:$H$" & t)
    .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 128, 0)
    .SeriesCollection(1).MarkerForegroundColor = RGB(0, 128, 0)
    .SeriesCollection(1).MarkerBackgroundColor = RGB(0, 128, 0)
End With
    
    ActiveWorkbook.ActiveSheet.Range("A:A").NumberFormat = "Text"
  
'More formatting and aesthetics
With Range("B3").Borders
.LineStyle = xlContinuous
.Weight = xlThick
End With
Range("B3").NumberFormat = "General"


With Range("G3").Borders
.LineStyle = xlContinuous
.Weight = xlThick
End With
Range("G3").NumberFormat = "General"

Range("B3").Select

'NEXT SHEET
CheckName1:

'Check for sheets with same name
For y = 1 To ActiveWorkbook.Sheets.Count
If ThirdName = ActiveWorkbook.Sheets(y).Name Then
MsgBox ("Sheet name already taken for total turnout, please rename the sheet.")
i = 1
GoTo ENDIT
Else
End If
Next y

'Adds new worksheet and names it
ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.ActiveSheet
ActiveWorkbook.ActiveSheet.Name = ThirdName

'Formatting and labeling
ActiveWorkbook.ActiveSheet.Range("A1") = "All Precincts"
ActiveWorkbook.ActiveSheet.Range("A1:E1").Merge

p = 5

'Calculates counts and percentages for turnout across the entire data sheet
While Early2 <= Late
With ActiveWorkbook.ActiveSheet
    .Range("B" & p) = Early2
    .Range("C" & p).Formula = "=COUNTIFS('" & FirstName & "'!C4,"">=""&'" & ThirdName & "'!RC[-1],'" & FirstName & "'!C4,""<""&'" & ThirdName & "'!R[1]C[-1])"
    .Range("D" & p) = "=C" & p & "/C2*100"
    .Range("E" & p) = "=C" & p & "/C2"
End With
Early2 = DateAdd("h", 1, Early2)
p = p + 1
Wend

'Formatting and labeling
With ActiveWorkbook.ActiveSheet
    .Range("B2") = "Total Count:"
    .Range("C2") = "=sum(C5:C" & p - 1 & ")"
    .Range("B4") = "Time"
    .Range("C4") = "Count"
    .Range("D4") = "Percent"
    .Range("E4") = "Simio Input"
    .Range("B:B").NumberFormat = "h:mm AM/PM"
    .Columns("A").AutoFit
    .Columns("B").AutoFit
    .Columns("E").AutoFit
    .Range("A1").HorizontalAlignment = xlCenter
End With

'Creates figures to display turnout counts and percentages per hour
ActiveWorkbook.ActiveSheet.Range("B4:C" & p - 1).Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
With ActiveChart
    .SetSourceData Source:=ActiveWorkbook.ActiveSheet.Range("$B$4:$C$" & p - 1)
    .ChartTitle.Text = "All Precincts"
    .Parent.Top = -100
    .Parent.Left = 300
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Count of Voters"
    .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(200, 0, 255)
    .SeriesCollection(1).MarkerForegroundColor = RGB(200, 0, 255)
    .SeriesCollection(1).MarkerBackgroundColor = RGB(200, 0, 255)
End With

ActiveWorkbook.ActiveSheet.Range("B4:B" & p - 1 & ",D4:D" & p - 1).Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
With ActiveChart
    .SetSourceData Source:=ActiveWorkbook.ActiveSheet.Range("B4:B" & p - 1 & ",D4:D" & p - 1)
    .ChartTitle.Text = "All Precincts"
    .Parent.Top = 215
    .Parent.Left = 300
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percent of Voters"
    .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(200, 0, 255)
    .SeriesCollection(1).MarkerForegroundColor = RGB(200, 0, 255)
    .SeriesCollection(1).MarkerBackgroundColor = RGB(200, 0, 255)
End With
    
    
'Aesthetics
With ActiveWorkbook.ActiveSheet
    .Range("A2").Font.Bold = True
    .Range("A1").Font.Bold = True
    .Range("B2").Font.Bold = True
    .Range("B4").Font.Bold = True
    .Range("C4").Font.Bold = True
    .Range("D4").Font.Bold = True
    .Range("E4").Font.Bold = True
    .Range("A3").Select
End With

ENDIT:

Application.ScreenUpdating = True

End Sub

Sub DSStatTable()


Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

'Store name information
Name = Left(ActiveWorkbook.ActiveSheet.Name, 21) + "... Stats"
i = 0

'Check if sheet name is already taken
For y = 1 To ActiveWorkbook.Sheets.Count
If Name = ActiveWorkbook.Sheets(y).Name Then
MsgBox ("Sheet name already taken, please rename the sheet.")
i = 1
GoTo ENDIT
Else
End If
Next y

'Determine the data range you want to pivot
  SrcData = ActiveWorkbook.ActiveSheet.Name & "!" & ActiveWorkbook.ActiveSheet.UsedRange.Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable2")

pvt.AddDataField pvt.PivotFields("Scan Type"), "Count of Scan Type", xlCount
pvt.AddDataField pvt.PivotFields("Scan Type"), "Percent of Scan Type", xlCount
pvt.PivotFields("Percent of Scan Type").Calculation = xlPercentOfColumn
pvt.AddDataField pvt.PivotFields("Duration (mm:ss)"), "Average Duration of Scan Type", xlAverage
pvt.AddDataField pvt.PivotFields("Duration (mm:ss)"), "Max Duration of Scan Type", xlMax
pvt.AddDataField pvt.PivotFields("Duration (mm:ss)"), "Standard Deviation of Scan Type", xlStDev
pvt.PivotFields("Average Duration of Scan Type").NumberFormat = "mm:ss"
pvt.PivotFields("Max Duration of Scan Type").NumberFormat = "mm:ss"
pvt.PivotFields("Standard Deviation of Scan Type").NumberFormat = "mm:ss"


pvt.PivotFields("Scan Type").Orientation = xlRowField

'Formatting and labeling
ActiveSheet.Name = Name
ActiveSheet.Range("A2").Font.Bold = True
ActiveSheet.Range("A2") = Name


ENDIT:

End Sub


