VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements OutputWriter

Private sheet As Worksheet
Private row_num As Long

Public Sub class_initialize()
    row_num = 1
End Sub

Public Sub setOutputSheet(ws As Worksheet)
    Set sheet = ws
End Sub

Public Sub OutputWriter_writeLineArr(lineArray() As String)
    Dim rangeEnd As String
    rangeEnd = Util.getLetterFromNumber(Util.getStringArrayLength(lineArray)) & row_num
    sheet.Range("A" & row_num, rangeEnd) = lineArray
    row_num = row_num + 1
End Sub

Public Sub OutputWriter_writeLine(ParamArray var() As Variant)
    Dim lineArr() As String
    Dim itemCount As Integer
    itemCount = UBound(var) - LBound(var) + 1
    ReDim lineArr(itemCount - 1)
    For i = 0 To itemCount - 1
        lineArr(i) = var(LBound(var) + i)
    Next i
    OutputWriter_writeLineArr lineArr
End Sub

Public Function OutputWriter_getRowNum() As Long
    getRowNum = row_num
End Function

Public Sub formatPretty()
    Dim columns As Integer
    columns = sheet.UsedRange.columns.count
    sheet.Range("A1", Util.getLetterFromNumber(columns) & "1").Font.Bold = True
    sheet.UsedRange.columns.AutoFit
End Sub

Public Sub OutputWriter_done()
    formatPretty
End Sub
