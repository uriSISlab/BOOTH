Attribute VB_Name = "Automate"
    Option Explicit
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
    
    Sub openAndProcessDirectory(control As IRibbonControl)
        Const ONE_FILE_LIMIT As Long = 5000
        
        Dim folder As String
        ' Create folder picker
        With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        .AllowMultiSelect = False
        folder = CStr(.SelectedItems(1))
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
        filesGlob = folder & "\BEL_*_*.log"
        fileCount = CountFiles(filesGlob)
        fileName = Dir(filesGlob)
    
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(outputFileName)
        Dim inputStream As TextStream
    
        ' Show progress bar
        UserForm1.Show vbModeless
        progress 0
        
        'Loop to process multiple files consecutively
        Do While Len(fileName) > 0
            Set inputStream = fso.OpenTextFile(folder & "\" & fileName, ForReading, False)
            Dim writeHeader As Boolean
            If outputRow = 2 Then
                writeHeader = True
            Else
                writeHeader = False
            End If
            outputRow = Process_Single_VSAPBMD_Data_From_Stream(inputStream, fileStream, outputRow, writeHeader, True, fileName)
            ActiveWorkbook.ActiveSheet.UsedRange.ClearContents
            fileName = Dir
            fileNum = fileNum + 1
            If fileNum Mod 5 = 0 Then
                progress fileNum / fileCount * 100
            End If
            inputStream.Close
        Loop
        fileStream.Close
        
        ' Stop showing progress bar
        Unload UserForm1
        
        'Allow the Excel file to actively update
        Application.ScreenUpdating = True
    End Sub
