VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataSourceForm 
   Caption         =   "Data Source"
   ClientHeight    =   3780
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4290
   OleObjectBlob   =   "DataSourceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataSourceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub DominionImagecastData_Click()
SubmitSource.Enabled = True
End Sub

Private Sub DS200Data_Click()
SubmitSource.Enabled = True
End Sub

Private Sub PollPadData_Click()
SubmitSource.Enabled = True
End Sub

Private Sub SubmitSource_Click()
If PollPadData = True Then
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
            C = 1
            While C < ActiveWorkbook.Sheets.Count + 1
                If ActiveWorkbook.Sheets(C).Name = Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".csv", ""), 10) & " PollPad" Then
                        MsgBox (FileNam & " shares the first 10 characters with a current worksheet. Please rename the file and import again.")
                        GoTo skipit
                Else:
                    C = C + 1
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


Else
    If DS200Data = True Then
                             
                
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
                
                                
                
                    'Check for duplicate precincts and delete the duplicate sheets
                    C = 1
                    While C < ActiveWorkbook.Sheets.Count + 1
                        If ActiveWorkbook.Sheets(C).Name = "Precinct " & Left(Replace(Right$(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), Len(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j)) - InStrRev(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(j), "\")), ".txt", ""), 10) Then
                            GoTo skipit2
                        Else: C = C + 1
                        End If
                    Wend
                
                    'Add an additional sheet and activate it to populate it with DS200 data
                    ActiveWorkbook.Sheets.Add after:=ActiveWorkbook.Sheets(j)
                    ActiveWorkbook.Sheets(j + 1).Activate
                
                
                    'Pulling file path for a specific file

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
skipit2:
                
                Next j
                              
                'Allow the Excel file to actively update
                Application.ScreenUpdating = True
    Else
        If DominionImagecastData = True Then
            'call
        Else: End If
    End If
End If
    
PollPadData = False
DS200Data = False
DominionImagecastData = False
SubmitSource.Enabled = False

    
End Sub

Private Sub UserForm_Initialize()
SubmitSource.Enabled = False
End Sub
