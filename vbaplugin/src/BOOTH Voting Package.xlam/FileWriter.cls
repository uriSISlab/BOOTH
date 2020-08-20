VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements OutputWriter

Private filePath As String
Private fso As FileSystemObject
Private stream As TextStream
Private row_num As Long

Public Sub class_initialize()
    row_num = 1
End Sub

Public Sub setFilePath(fp As String)
    filePath = fp
    Set fso = New FileSystemObject
    Set stream = fso.CreateTextFile(fp)
End Sub

Public Sub OutputWriter_writeLineArr(lineArray() As String)
    ' TODO wrap item in quotes if it contains a comma so that output proper CSV
    stream.writeLine Join(lineArray, ", ")
    row_num = row_num + 1
End Sub

Public Sub OutputWriter_writeLine(ParamArray params() As Variant)
    Dim lineArr() As String
    Dim itemCount As Integer
    itemCount = UBound(params) - LBound(params) + 1
    ReDim lineArr(itemCount - 1)
    For i = 0 To itemCount - 1
        lineArr(i) = params(LBound(params) + i)
    Next i
    OutputWriter_writeLineArr lineArr
End Sub

Public Function OutputWriter_getRowNum() As Long
    getRowNum = row_num
End Function

Public Sub OutputWriter_done()
    stream.Close
End Sub


