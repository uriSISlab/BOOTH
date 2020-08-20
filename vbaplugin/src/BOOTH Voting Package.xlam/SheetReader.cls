VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements InputReader

Private sheet As Worksheet
Private separator As String
Private currentLine As Long
Private columns As Integer
Private begun As Boolean

Public Sub setSheetAndSeparator(sh As Worksheet, sp As String)
    Set sheet = sh
    separator = sp
    columns = sh.UsedRange.columns.count
    currentLine = 1
    begun = False
End Sub

Public Sub InputReader_setSkipLines(skipCount As Integer)
    currentLine = 1 + skipCount
End Sub

Public Function InputReader_noMoreLines() As Boolean
    InputReader_noMoreLines = currentLine > sheet.UsedRange.rows.count
End Function

Public Function InputReader_readLine() As String
    Dim line As String
    Dim j As Integer
    line = CStr(sheet.Range("A" & currentLine))
    For j = 2 To columns
        ' Join the row with the separator
        line = line & separator & sheet.Range(Util.getLetterFromNumber(j) & currentLine)
    Next j
    currentLine = currentLine + 1
    InputReader_readLine = line
End Function
