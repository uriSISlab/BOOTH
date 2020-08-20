Attribute VB_Name = "Util"
Public Enum LogType
    VSAP_BMD
    DICE
    DICX
    Unknown
End Enum

Public Enum IOType
    File
    sheet
End Enum

Public Function getAllLogTypes() As LogType()
    Dim logTypes(3) As LogType
    logTypes(0) = VSAP_BMD
    logTypes(1) = DICE
    logTypes(2) = DICX
    getAllLogTypes = logTypes
End Function

Public Function getFileNamePatternForLog(t As LogType) As String
    If t = VSAP_BMD Then
        getFileNamePatternForLog = "BEL_*_*.log"
    ElseIf t = DICE Then
        getFileNamePatternForLog = "*.TXT"
    ElseIf t = DICX Then
        getFileNamePatternForLog = "ICX_AUDIT_LOG.*.log"
    Else
        getFileNamePatternForLog = "*"
    End If
End Function

Public Function getStringArrayLength(arr() As String) As Long
    getStringArrayLength = UBound(arr) - LBound(arr) + 1
End Function

Public Function getLetterFromNumber(number As Integer) As String
    ' Return the corresponding letter from a number, 1 returns "A", 2 returns "B", and so on.
    getLetterFromNumber = Chr(Asc("A") + number - 1)
End Function

Public Function getDifferenceMinutes(startTime As Date, endTime As Date) As Integer
    getDifferenceMinutes = DateDiff("s", startTime, endTime) \ 60
End Function

Public Function getTimeDifference(startTime As Date, endTime As Date) As String
    Dim mins As String
    Dim secs As String
    Dim totalSecs As Long
    totalSecs = CLng(DateDiff("s", startTime, endTime))
    mins = format(CStr(totalSecs \ 60), "00")
    secs = format(CStr(totalSecs Mod 60), "00")
    getTimeDifference = mins & ":" & secs
End Function

Public Function createProcessor(t As LogType) As LogProcessor
    Dim processor As LogProcessor
    If t = VSAP_BMD Then
        Set processor = New VSAPBMD_Processor
    ElseIf t = DICE Then
        Set processor = New DICE_Processor
    ElseIf t = DICX Then
        Set processor = New DICX_Processor
    End If
    Set createProcessor = processor
End Function

Public Function createReader(t As IOType) As InputReader
    Dim reader As InputReader
    If t = IOType.File Then
        Set reader = New FileReader
    ElseIf t = IOType.sheet Then
        Set reader = New SheetReader
    End If
    createReader = reader
End Function

Public Function createWriter(t As IOType) As OutputWriter
    Dim writer As OutputWriter
    If t = IOType.File Then
        Set writer = New FileWriter
    ElseIf t = IOType.sheet Then
        Set writer = New SheetWriter
    End If
    createWriter = writer
End Function

Public Function getLogTypeForLog(sheet As Worksheet) As LogType
    Dim logTypes() As LogType
    Dim processor As LogProcessor
    logTypes = Util.getAllLogTypes
    For i = LBound(logTypes) To UBound(logTypes)
        Set processor = Util.createProcessor(logTypes(i))
        If processor.isThisLog(sheet) Then
            getLogTypeForLog = logTypes(i)
            Exit Function
        End If
    Next i
    getCorrectProcessorForLog = LogType.Unknown
End Function

Public Function getProcessorForLog(sheet As Worksheet) As LogProcessor
    Dim logTypes() As LogType
    Dim processor As LogProcessor
    logTypes = Util.getAllLogTypes
    For i = LBound(logTypes) To UBound(logTypes)
        processor = Util.createProcessor(logTypes(i))
        If processor.isThisLog(sheet) Then
            getCorrectProcessorForLog = processor
            Exit Function
        End If
    Next i
    getCorrectProcessorForLog = Nothing
End Function

Public Sub runPipeline(reader As InputReader, processor As LogProcessor, writer As OutputWriter, writeHeader As Boolean)
    processor.setWriter writer
    If writeHeader Then
        processor.writeHeader
    End If
    Do While Not reader.noMoreLines
        Dim line As String
        line = reader.readLine
        processor.readLine line
    Loop
End Sub

Public Function getProcessedName(name As String) As String
    Dim processedName As String
    processedName = name & " Processed"
    If Len(processedName) > 31 Then
        processedName = Right(processedName, 31)
    End If
    getProcessedName = processedName
End Function

Public Function appendToArray(arr() As String, item As String) As String()
    Dim fullLineArr() As String
    If Len(item) > 0 Then
        ReDim fullLineArr(Util.getStringArrayLength(arr))
        ' Copy contents of arr to fullLineArr first
        For i = LBound(arr) To UBound(arr)
            fullLineArr(i) = arr(i)
        Next i
        fullLineArr(UBound(fullLineArr)) = item
    Else
        fullLineArr = arr
    End If
    appendToArray = fullLineArr
End Function

