VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements InputReader

Private filePath As String
Private fso As FileSystemObject
Private stream As TextStream
Private endReached As Boolean
Private begun As Boolean
Private skipLines As Integer

Public Sub setFilePath(fp As String)
    filePath = fp
    Set fso = New FileSystemObject
    Set stream = fso.OpenTextFile(filePath, ForReading, False)
    endReached = False
    begun = False
End Sub

Public Sub InputReader_setSkipLines(skipCount As Integer)
    skipLines = skipCount
End Sub

Public Function InputReader_noMoreLines() As Boolean
    If endReached Then
        InputReader_noMoreLines = True
    Else
        If stream.AtEndOfStream Then
            endReached = True
            stream.Close
        End If
        InputReader_noMoreLines = endReached
    End If
End Function

Public Function InputReader_readLine() As String
    If Not begun Then
        For i = 0 To skipLines
            stream.readLine
        Next i
        begun = True
    End If
    InputReader_readLine = stream.readLine
End Function
