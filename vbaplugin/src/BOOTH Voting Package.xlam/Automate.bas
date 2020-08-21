Attribute VB_Name = "Automate"
    Option Explicit
    Function CountFiles(glob As String) As Long
        Dim filename As String
        Dim count As Long
        count = 0
        filename = Dir(glob)
        Do While Len(filename) > 0
            count = count + 1
            filename = Dir
        Loop
        CountFiles = count
    End Function
    
Sub openAndProcessDirectory(control As IRibbonControl)
        Const ONE_FILE_LIMIT As Long = 5000
        
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
        
        Dim filename As String
        Dim filesGlob As String
        filesGlob = folder & "\BEL_*_*.log"
        fileCount = CountFiles(filesGlob)
        filename = Dir(filesGlob)
    
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(outputFileName)
        Dim inputStream As TextStream
    
        ' Show progress bar
        UserForm1.Show vbModeless
        progress 0
        
        Dim writer As OutputWriter
        Set writer = New OutputWriter
        writer.setOutputStream fileStream
        'TODO Do this dynamically
        Dim header() As String
        header = Util.getVSAPBMDHeader
        Dim fileNameArr(1) As String
        
        'Loop to process multiple files consecutively
        Do While Len(filename) > 0
            Set inputStream = fso.OpenTextFile(folder & "\" & filename, ForReading, False)

            Dim processor As VSAPBMD_Processor
            Set processor = New VSAPBMD_Processor
            processor.setWriter writer
            processor.setFileName filename
            If outputRow = 2 Then
                processor.writeHeader
            End If

            Process_Single_VSAPBMD_Data_From_Stream inputStream, processor
            outputRow = writer.getRowNum
            filename = Dir
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