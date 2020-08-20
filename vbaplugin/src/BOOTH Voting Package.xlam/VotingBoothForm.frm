VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VotingBoothForm 
   Caption         =   "Privacy Booth Timer"
   ClientHeight    =   9410.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   17980
   OleObjectBlob   =   "VotingBoothForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VotingBoothForm"
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

Private Sub StartBooth1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 3) = time
StopBooth1.Enabled = True
StartBooth1.Enabled = False
Image1.BorderColor = &HFF00&
UndoLast1.Enabled = True
StartBooth1.BackColor = &HFF00&
End Sub

Private Sub StartBooth2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 6) = time
StopBooth2.Enabled = True
StartBooth2.Enabled = False
Image2.BorderColor = &HFF00&
UndoLast2.Enabled = True
StartBooth2.BackColor = &HFF00&
End Sub



Private Sub StartBooth3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 9) = time
StopBooth3.Enabled = True
StartBooth3.Enabled = False
Image3.BorderColor = &HFF00&
UndoLast3.Enabled = True
StartBooth3.BackColor = &HFF00&
End Sub

Private Sub StartBooth4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = time
StopBooth4.Enabled = True
StartBooth4.Enabled = False
Image4.BorderColor = &HFF00&
UndoLast4.Enabled = True
StartBooth4.BackColor = &HFF00&
End Sub


Private Sub StopBooth1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4) = time
ActiveSheet.Cells(nr, 5) = ActiveSheet.Cells(nr, 4) - ActiveSheet.Cells(nr, 3)
StopBooth1.Enabled = False
StartBooth1.Enabled = True
Image1.BorderColor = &H80000011
TextBox1 = ""
StartBooth1.BackColor = &H8000000F
End Sub

Private Sub StopBooth2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row
ActiveSheet.Cells(nr, 7) = time
ActiveSheet.Cells(nr, 8) = ActiveSheet.Cells(nr, 7) - ActiveSheet.Cells(nr, 6)
StopBooth2.Enabled = False
StartBooth2.Enabled = True
Image2.BorderColor = &H80000011
TextBox2 = ""
StartBooth2.BackColor = &H8000000F
End Sub

Private Sub StopBooth3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10) = time
ActiveSheet.Cells(nr, 11) = ActiveSheet.Cells(nr, 10) - ActiveSheet.Cells(nr, 9)
StopBooth3.Enabled = False
StartBooth3.Enabled = True
Image3.BorderColor = &H80000011
TextBox3 = ""
StartBooth3.BackColor = &H8000000F
End Sub

Private Sub StopBooth4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row
ActiveSheet.Cells(nr, 13) = time
ActiveSheet.Cells(nr, 14) = ActiveSheet.Cells(nr, 13) - ActiveSheet.Cells(nr, 12)
StopBooth4.Enabled = False
StartBooth4.Enabled = True
Image4.BorderColor = &H80000011
TextBox4 = ""
StartBooth4.BackColor = &H8000000F
End Sub

Private Sub UndoLast1_Click()

If (ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4).Clear
ActiveSheet.Cells(nr, 3).Clear
ActiveSheet.Cells(nr, 5).Clear
TextBox1 = ""
StopBooth1.Enabled = False
StartBooth1.Enabled = True
Image1.BorderColor = &H80000011
StartBooth1.BackColor = &H8000000F
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
StopBooth2.Enabled = False
StartBooth2.Enabled = True
Image2.BorderColor = &H80000011
StartBooth2.BackColor = &H8000000F
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
StopBooth3.Enabled = False
StartBooth3.Enabled = True
Image3.BorderColor = &H80000011
StartBooth3.BackColor = &H8000000F
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
StopBooth4.Enabled = False
StartBooth4.Enabled = True
Image4.BorderColor = &H80000011
StartBooth4.BackColor = &H8000000F
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

Private Sub StopBooth5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16) = time
ActiveSheet.Cells(nr, 17) = ActiveSheet.Cells(nr, 16) - ActiveSheet.Cells(nr, 15)
StopBooth5.Enabled = False
StartBooth5.Enabled = True
Image5.BorderColor = &H80000011
TextBox5 = ""
StartBooth5.BackColor = &H8000000F
End Sub
Private Sub StopBooth6_Click()
nr = ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row
ActiveSheet.Cells(nr, 19) = time
ActiveSheet.Cells(nr, 20) = ActiveSheet.Cells(nr, 19) - ActiveSheet.Cells(nr, 18)
StopBooth6.Enabled = False
StartBooth6.Enabled = True
Image6.BorderColor = &H80000011
TextBox6 = ""
StartBooth6.BackColor = &H8000000F
End Sub


Private Sub UndoLast5_Click()

If (ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16).Clear
ActiveSheet.Cells(nr, 15).Clear
ActiveSheet.Cells(nr, 17).Clear
TextBox5 = ""
StopBooth5.Enabled = False
StartBooth5.Enabled = True
Image5.BorderColor = &H80000011
StartBooth5.BackColor = &H8000000F
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
StopBooth6.Enabled = False
StartBooth6.Enabled = True
Image6.BorderColor = &H80000011
StartBooth6.BackColor = &H8000000F
UndoLast6.Enabled = False
Else: End If

End Sub

Private Sub StartBooth5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 15) = time
StopBooth5.Enabled = True
StartBooth5.Enabled = False
Image5.BorderColor = &HFF00&
UndoLast5.Enabled = True
StartBooth5.BackColor = &HFF00&
End Sub
Private Sub StartBooth6_Click()
nr = ActiveSheet.Cells(Rows.count, 18).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 18) = time
StopBooth6.Enabled = True
StartBooth6.Enabled = False
Image6.BorderColor = &HFF00&
UndoLast6.Enabled = True
StartBooth6.BackColor = &HFF00&
End Sub
Private Sub SaveComment_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 21) = VotingBoothForm.CommentBox.Value
VotingBoothForm.CommentBox.Value = ""

End Sub

Private Sub UserForm_Initialize()
StopBooth1.Enabled = False
StopBooth2.Enabled = False
StopBooth3.Enabled = False
StopBooth4.Enabled = False
StopBooth5.Enabled = False
StopBooth6.Enabled = False
UndoLast1.Enabled = False
UndoLast2.Enabled = False
UndoLast3.Enabled = False
UndoLast4.Enabled = False
UndoLast5.Enabled = False
UndoLast6.Enabled = False
ActiveSheet.Cells(1, 1) = "VotingBooth1_Start"
ActiveSheet.Cells(1, 2) = "VotingBooth1_Stop"
ActiveSheet.Cells(1, 3) = "VotingBooth1_Duration"
ActiveSheet.Cells(1, 4) = "VotingBooth2_Start"
ActiveSheet.Cells(1, 5) = "VotingBooth2_Stop"
ActiveSheet.Cells(1, 6) = "VotingBooth2_Duration"
ActiveSheet.Cells(1, 7) = "VotingBooth3_Start"
ActiveSheet.Cells(1, 8) = "VotingBooth3_Stop"
ActiveSheet.Cells(1, 9) = "VotingBooth3_Duration"
ActiveSheet.Cells(1, 10) = "VotingBooth4_Start"
ActiveSheet.Cells(1, 11) = "VotingBooth4_Stop"
ActiveSheet.Cells(1, 12) = "VotingBooth4_Duration"
ActiveSheet.Cells(1, 13) = "VotingBooth5_Start"
ActiveSheet.Cells(1, 14) = "VotingBooth5_Stop"
ActiveSheet.Cells(1, 15) = "VotingBooth5_Duration"
ActiveSheet.Cells(1, 16) = "VotingBooth6_Start"
ActiveSheet.Cells(1, 17) = "VotingBooth6_Stop"
ActiveSheet.Cells(1, 18) = "VotingBooth6_Duration"
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
