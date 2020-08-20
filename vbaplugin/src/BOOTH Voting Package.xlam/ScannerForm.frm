VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScannerForm 
   Caption         =   "Ballot Scanner Timer"
   ClientHeight    =   9410.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   17980
   OleObjectBlob   =   "ScannerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScannerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Clear1_Click()
TextBox1 = ""
End Sub

Private Sub SaveButton_Click()
    ActiveWorkbook.Save
End Sub

Private Sub StartScan1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 3) = time
StopScan1.Enabled = True
StartScan1.Enabled = False
Image1.BorderColor = &HFF00&
UndoLast1.Enabled = True
StartScan1.BackColor = &HFF00&
End Sub

Private Sub StartScan2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 6) = time
StopScan2.Enabled = True
StartScan2.Enabled = False
Image2.BorderColor = &HFF00&
UndoLast2.Enabled = True
StartScan2.BackColor = &HFF00&
End Sub

Private Sub StartScan3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 9) = time
StopScan3.Enabled = True
StartScan3.Enabled = False
Image3.BorderColor = &HFF00&
UndoLast3.Enabled = True
StartScan3.BackColor = &HFF00&
End Sub

Private Sub StartScan4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = time
StopScan4.Enabled = True
StartScan4.Enabled = False
Image4.BorderColor = &HFF00&
UndoLast4.Enabled = True
StartScan4.BackColor = &HFF00&
End Sub


Private Sub StopScan1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4) = time
ActiveSheet.Cells(nr, 5) = ActiveSheet.Cells(nr, 4) - ActiveSheet.Cells(nr, 3)
StopScan1.Enabled = False
StartScan1.Enabled = True
Image1.BorderColor = &H80000011
TextBox1 = ""
StartScan1.BackColor = &H8000000F
End Sub

Private Sub StopScan2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row
ActiveSheet.Cells(nr, 7) = time
ActiveSheet.Cells(nr, 8) = ActiveSheet.Cells(nr, 7) - ActiveSheet.Cells(nr, 6)
StopScan2.Enabled = False
StartScan2.Enabled = True
Image2.BorderColor = &H80000011
TextBox2 = ""
StartScan2.BackColor = &H8000000F
End Sub

Private Sub StopScan3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10) = time
ActiveSheet.Cells(nr, 11) = ActiveSheet.Cells(nr, 10) - ActiveSheet.Cells(nr, 9)
StopScan3.Enabled = False
StartScan3.Enabled = True
Image3.BorderColor = &H80000011
TextBox3 = ""
StartScan3.BackColor = &H8000000F
End Sub

Private Sub StopScan4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row
ActiveSheet.Cells(nr, 13) = time
ActiveSheet.Cells(nr, 14) = ActiveSheet.Cells(nr, 13) - ActiveSheet.Cells(nr, 12)
StopScan4.Enabled = False
StartScan4.Enabled = True
Image4.BorderColor = &H80000011
TextBox4 = ""
StartScan4.BackColor = &H8000000F
End Sub



Private Sub UndoLast1_Click()

If (ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4).Clear
ActiveSheet.Cells(nr, 3).Clear
ActiveSheet.Cells(nr, 5).Clear

TextBox1 = ""
StopScan1.Enabled = False
StartScan1.Enabled = True
Image1.BorderColor = &H80000011
StartScan1.BackColor = &H8000000F
UndoLast1.Enabled = False
Else: End If



End Sub

Private Sub UndoLast2_Click()

If (ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row
ActiveSheet.Cells(nr, 7).Clear
ActiveSheet.Cells(nr, 6).Clear
ActiveSheet.Cells(nr, 8).Clear

TextBox2 = ""
StopScan2.Enabled = False
StartScan2.Enabled = True
Image2.BorderColor = &H80000011
StartScan2.BackColor = &H8000000F
UndoLast2.Enabled = False
Else: End If

End Sub

Private Sub UndoLast3_Click()

If (ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10).Clear
ActiveSheet.Cells(nr, 9).Clear
ActiveSheet.Cells(nr, 11).Clear

TextBox3 = ""
StopScan3.Enabled = False
StartScan3.Enabled = True
Image3.BorderColor = &H80000011
StartScan3.BackColor = &H8000000F
UndoLast3.Enabled = False
Else: End If

End Sub

Private Sub UndoLast4_Click()

If (ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row
ActiveSheet.Cells(nr, 13).Clear
ActiveSheet.Cells(nr, 12).Clear
ActiveSheet.Cells(nr, 14).Clear

TextBox4 = ""
StopScan4.Enabled = False
StartScan4.Enabled = True
Image4.BorderColor = &H80000011
StartScan4.BackColor = &H8000000F
UndoLast4.Enabled = False
Else: End If

End Sub

Private Sub Clear2_Click()
TextBox2 = ""
End Sub

Private Sub Clear3_Click()
TextBox3 = ""
End Sub

Private Sub Clear4_Click()
TextBox4 = ""
End Sub

Private Sub Clear5_Click()
TextBox5 = ""
End Sub

Private Sub Clear6_Click()
TextBox6 = ""
End Sub

Private Sub StopScan5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16) = time
ActiveSheet.Cells(nr, 17) = ActiveSheet.Cells(nr, 16) - ActiveSheet.Cells(nr, 15)
StopScan5.Enabled = False
StartScan5.Enabled = True
Image5.BorderColor = &H80000011
TextBox5 = ""
StartScan5.BackColor = &H8000000F
End Sub
Private Sub StopScan6_Click()
nr = ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row
ActiveSheet.Cells(nr, 19) = time
ActiveSheet.Cells(nr, 20) = ActiveSheet.Cells(nr, 19) - ActiveSheet.Cells(nr, 18)
StopScan6.Enabled = False
StartScan6.Enabled = True
Image6.BorderColor = &H80000011
TextBox6 = ""
StartScan6.BackColor = &H8000000F
End Sub


Private Sub UndoLast5_Click()

If (ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16).Clear
ActiveSheet.Cells(nr, 15).Clear
ActiveSheet.Cells(nr, 17).Clear
TextBox5 = ""
StopScan5.Enabled = False
StartScan5.Enabled = True
Image5.BorderColor = &H80000011
StartScan5.BackColor = &H8000000F
UndoLast5.Enabled = False
Else: End If

End Sub
Private Sub UndoLast6_Click()

If (ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row
ActiveSheet.Cells(nr, 19).Clear
ActiveSheet.Cells(nr, 18).Clear
ActiveSheet.Cells(nr, 20).Clear

TextBox6 = ""
StopScan6.Enabled = False
StartScan6.Enabled = True
Image6.BorderColor = &H80000011
StartScan6.BackColor = &H8000000F
UndoLast6.Enabled = False
Else: End If

End Sub

Private Sub StartScan5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 15) = time
StopScan5.Enabled = True
StartScan5.Enabled = False
Image5.BorderColor = &HFF00&
UndoLast5.Enabled = True
StartScan5.BackColor = &HFF00&
End Sub
Private Sub StartScan6_Click()
nr = ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 18) = time
StopScan6.Enabled = True
StartScan6.Enabled = False
Image6.BorderColor = &HFF00&
UndoLast6.Enabled = True
StartScan6.BackColor = &HFF00&
End Sub

Private Sub SaveComment_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 21) = VotingBoothForm.CommentBox.Value
VotingBoothForm.CommentBox.Value = ""

End Sub


Private Sub UserForm_Initialize()
StopScan1.Enabled = False
StopScan2.Enabled = False
StopScan3.Enabled = False
StopScan4.Enabled = False
StopScan5.Enabled = False
StopScan6.Enabled = False
UndoLast1.Enabled = False
UndoLast2.Enabled = False
UndoLast3.Enabled = False
UndoLast4.Enabled = False
UndoLast5.Enabled = False
UndoLast6.Enabled = False
ActiveSheet.Cells(1, 1) = "Scanner1_Start"
ActiveSheet.Cells(1, 2) = "Scanner1_Stop"
ActiveSheet.Cells(1, 3) = "Scanner1_Duration"
ActiveSheet.Cells(1, 4) = "Scanner2_Start"
ActiveSheet.Cells(1, 5) = "Scanner2_Stop"
ActiveSheet.Cells(1, 6) = "Scanner2_Duration"
ActiveSheet.Cells(1, 7) = "Scanner3_Start"
ActiveSheet.Cells(1, 8) = "Scanner3_Stop"
ActiveSheet.Cells(1, 9) = "Scanner3_Duration"
ActiveSheet.Cells(1, 10) = "Scanner4_Start"
ActiveSheet.Cells(1, 11) = "Scanner4_Stop"
ActiveSheet.Cells(1, 12) = "Scanner4_Duration"
ActiveSheet.Cells(1, 13) = "Scanner5_Start"
ActiveSheet.Cells(1, 14) = "Scanner5_Stop"
ActiveSheet.Cells(1, 15) = "Scanner5_Duration"
ActiveSheet.Cells(1, 16) = "Scanner6_Start"
ActiveSheet.Cells(1, 17) = "Scanner6_Stop"
ActiveSheet.Cells(1, 18) = "Scanner6_Duration"
ActiveSheet.Cells(1, 19) = "Comments"
ActiveSheet.Range("C:C").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("F:F").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("I:I").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("L:L").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("O:O").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("R:R").NumberFormat = "hh:mm:ss"
ActiveSheet.columns("A:AA").AutoFit
ActiveSheet.Range("A1", "AA1").Font.Bold = True

End Sub
