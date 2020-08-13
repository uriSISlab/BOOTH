VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Enum OutputType
    ' Output to an excel worksheet
    sheet = 0
    ' Output to a text stream (usually a file)
    Stream = 1
End Enum

Private output_sheet As Worksheet
Private output_stream As TextStream
Private output_type As OutputType
Private row_num As Long

Public Sub class_initialize()
    row_num = 1
End Sub

Public Sub setOutputSheet(ws As Worksheet)
    Set output_sheet = ws
    output_type = OutputType.sheet
End Sub

Public Sub setOutputStream(ts As TextStream)
    Set output_stream = ts
    output_type = OutputType.Stream
End Sub

Public Sub writeLine(lineArray() As String)
    If output_type = OutputType.Stream Then
        ' TODO wrap item in quotes if it contains a comma so that output proper CSV
        output_stream.writeLine Join(lineArray, ", ")
    ElseIf output_type = OutputType.sheet Then
        Dim rangeEnd As String
        rangeEnd = Util.getLetterFromNumber(Util.getStringArrayLength(lineArray)) & row_num
        output_sheet.Range("A" & row_num, rangeEnd) = lineArray
        row_num = row_num + 1
    End If
End Sub
