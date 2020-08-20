VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub setWriter(w As OutputWriter)
End Sub

Public Sub setFilename(fn As String)
End Sub

Public Sub readLine(line As String)
End Sub

Public Sub writeHeader()
End Sub

Public Function isThisLog(sheet As Worksheet) As Boolean
End Function

Public Function getSeparator() As String
End Function
