VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BMDForm 
   Caption         =   "BMD Timer"
   ClientHeight    =   10770
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   18660
   OleObjectBlob   =   "BMDForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BMDForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub SaveButton_Click()
    ActiveWorkbook.Save
End Sub

Private Sub StartBMD1_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 1) = time
StopBMD1.Enabled = True
StartBMD1.Enabled = False
UndoLast1.Enabled = True
Image1.BorderColor = &HFF00&
StartBMD1.BackColor = &HFF00&
Help1.Enabled = True
End Sub

Private Sub StartBMD2_Click()
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 5) = time
StopBMD2.Enabled = True
StartBMD2.Enabled = False
UndoLast2.Enabled = True
Image2.BorderColor = &HFF00&
StartBMD2.BackColor = &HFF00&
Help2.Enabled = True
End Sub

Private Sub StartBMD3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 9) = time
StopBMD3.Enabled = True
StartBMD3.Enabled = False
UndoLast3.Enabled = True
Image3.BorderColor = &HFF00&
StartBMD3.BackColor = &HFF00&
Help3.Enabled = True
End Sub

Private Sub StartBMD4_Click()
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 13) = time
StopBMD4.Enabled = True
StartBMD4.Enabled = False
UndoLast4.Enabled = True
Image4.BorderColor = &HFF00&
StartBMD4.BackColor = &HFF00&
Help4.Enabled = True
End Sub


Private Sub StartBMD5_Click()
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 17) = time
StopBMD5.Enabled = True
StartBMD5.Enabled = False
UndoLast5.Enabled = True
Image5.BorderColor = &HFF00&
StartBMD5.BackColor = &HFF00&
Help5.Enabled = True
End Sub

Private Sub StartBMD6_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 21) = time
StopBMD6.Enabled = True
StartBMD6.Enabled = False
UndoLast6.Enabled = True
Image6.BorderColor = &HFF00&
StartBMD6.BackColor = &HFF00&
Help6.Enabled = True
End Sub

Private Sub StopBMD1_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
ActiveSheet.Cells(nr, 2) = time
ActiveSheet.Cells(nr, 3) = ActiveSheet.Cells(nr, 2) - ActiveSheet.Cells(nr, 1)
StopBMD1.Enabled = False
StartBMD1.Enabled = True
Image1.BorderColor = &H80000011
TextBox1 = ""
StartBMD1.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 4) = "") Then
ActiveSheet.Cells(nr, 4) = 0
Else: End If

Help1.Enabled = False

End Sub

Private Sub StopBMD2_Click()
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row
ActiveSheet.Cells(nr, 6) = time
ActiveSheet.Cells(nr, 7) = ActiveSheet.Cells(nr, 6) - ActiveSheet.Cells(nr, 5)
StopBMD2.Enabled = False
StartBMD2.Enabled = True
Image2.BorderColor = &H80000011
TextBox2 = ""
StartBMD2.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 8) = "") Then
ActiveSheet.Cells(nr, 8) = 0
Else: End If

Help2.Enabled = False

End Sub

Private Sub StopBMD3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10) = time
ActiveSheet.Cells(nr, 11) = ActiveSheet.Cells(nr, 10) - ActiveSheet.Cells(nr, 9)
StopBMD3.Enabled = False
StartBMD3.Enabled = True
Image3.BorderColor = &H80000011
TextBox3 = ""
StartBMD3.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 12) = "") Then
ActiveSheet.Cells(nr, 12) = 0
Else: End If

Help3.Enabled = False

End Sub

Private Sub StopBMD4_Click()
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row
ActiveSheet.Cells(nr, 14) = time
ActiveSheet.Cells(nr, 15) = ActiveSheet.Cells(nr, 14) - ActiveSheet.Cells(nr, 13)
StopBMD4.Enabled = False
StartBMD4.Enabled = True
Image4.BorderColor = &H80000011
TextBox4 = ""
StartBMD4.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 16) = "") Then
ActiveSheet.Cells(nr, 16) = 0
Else: End If

Help4.Enabled = False

End Sub

Private Sub StopBMD5_Click()
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row
ActiveSheet.Cells(nr, 18) = time
ActiveSheet.Cells(nr, 19) = ActiveSheet.Cells(nr, 18) - ActiveSheet.Cells(nr, 17)
StopBMD5.Enabled = False
StartBMD5.Enabled = True
Image5.BorderColor = &H80000011
TextBox5 = ""
StartBMD5.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 20) = "") Then
ActiveSheet.Cells(nr, 20) = 0
Else: End If

Help5.Enabled = False

End Sub

Private Sub StopBMD6_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row
ActiveSheet.Cells(nr, 22) = time
ActiveSheet.Cells(nr, 23) = ActiveSheet.Cells(nr, 22) - ActiveSheet.Cells(nr, 21)
StopBMD6.Enabled = False
StartBMD6.Enabled = True
Image6.BorderColor = &H80000011
TextBox6 = ""
StartBMD6.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 24) = "") Then
ActiveSheet.Cells(nr, 24) = 0
Else: End If

Help6.Enabled = False

End Sub

Private Sub Clear1_Click()
TextBox1 = ""
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


Private Sub UndoLast1_Click()

If (ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
ActiveSheet.Cells(nr, 2).Clear
ActiveSheet.Cells(nr, 1).Clear
ActiveSheet.Cells(nr, 3).Clear
ActiveSheet.Cells(nr, 4).Clear
StopBMD1.Enabled = False
StartBMD1.Enabled = True
UndoLast1.Enabled = False
Help1.Enabled = False
Image1.BorderColor = &H80000011
StartBMD1.BackColor = &H8000000F
TextBox1 = ""
Else: End If



End Sub

Private Sub UndoLast2_Click()

If (ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row
ActiveSheet.Cells(nr, 5).Clear
ActiveSheet.Cells(nr, 7).Clear
ActiveSheet.Cells(nr, 6).Clear
ActiveSheet.Cells(nr, 8).Clear
StopBMD2.Enabled = False
StartBMD2.Enabled = True
UndoLast2.Enabled = False
Help2.Enabled = False
Image2.BorderColor = &H80000011
StartBMD2.BackColor = &H8000000F
TextBox2 = ""
Else: End If

End Sub

Private Sub UndoLast3_Click()

If (ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 9).Clear
ActiveSheet.Cells(nr, 10).Clear
ActiveSheet.Cells(nr, 11).Clear
ActiveSheet.Cells(nr, 12).Clear
StopBMD3.Enabled = False
StartBMD3.Enabled = True
UndoLast3.Enabled = False
Help3.Enabled = False
Image3.BorderColor = &H80000011
StartBMD3.BackColor = &H8000000F
TextBox3 = ""
Else: End If

End Sub

Private Sub UndoLast4_Click()

If (ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row
ActiveSheet.Cells(nr, 13).Clear
ActiveSheet.Cells(nr, 14).Clear
ActiveSheet.Cells(nr, 15).Clear
ActiveSheet.Cells(nr, 16).Clear
StopBMD4.Enabled = False
StartBMD4.Enabled = True
UndoLast4.Enabled = False
Help4.Enabled = False
Image4.BorderColor = &H80000011
StartBMD4.BackColor = &H8000000F
TextBox4 = ""
Else: End If

End Sub

Private Sub UndoLast5_Click()
If (ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row
ActiveSheet.Cells(nr, 17).Clear
ActiveSheet.Cells(nr, 18).Clear
ActiveSheet.Cells(nr, 19).Clear
ActiveSheet.Cells(nr, 20).Clear
StopBMD5.Enabled = False
StartBMD5.Enabled = True
UndoLast5.Enabled = False
Help5.Enabled = False
Image5.BorderColor = &H80000011
StartBMD5.BackColor = &H8000000F
TextBox5 = ""
Else: End If

End Sub


Private Sub UndoLast6_Click()
If (ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row
ActiveSheet.Cells(nr, 21).Clear
ActiveSheet.Cells(nr, 22).Clear
ActiveSheet.Cells(nr, 23).Clear
ActiveSheet.Cells(nr, 24).Clear
StopBMD6.Enabled = False
StartBMD6.Enabled = True
UndoLast6.Enabled = False
Help6.Enabled = False
Image6.BorderColor = &H80000011
StartBMD6.BackColor = &H8000000F
TextBox6 = ""
Else: End If

End Sub

Private Sub Clear5_Click()
TextBox5 = ""
End Sub

Private Sub Help1_Click()
nr = ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 4) = "Helped"
Help1.Enabled = False
End Sub
Private Sub Help2_Click()
nr = ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 8) = "Helped"
Help2.Enabled = False
End Sub
Private Sub Help3_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = "Helped"
Help3.Enabled = False
End Sub
Private Sub Help4_Click()
nr = ActiveSheet.Cells(Rows.count, 16).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 16) = "Helped"
Help4.Enabled = False
End Sub
Private Sub Help5_Click()
nr = ActiveSheet.Cells(Rows.count, 20).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 20) = "Helped"
Help5.Enabled = False
End Sub
Private Sub Help6_Click()
nr = ActiveSheet.Cells(Rows.count, 24).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 24) = "Helped"
Help6.Enabled = False
End Sub
Private Sub Clear6_Click()
TextBox6 = ""
End Sub


Private Sub StoreComment_Click()
nr = ActiveSheet.Cells(Rows.count, 25).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 25) = CommentBox.Value
CommentBox.Value = ""
End Sub

Private Sub UserForm_Initialize()
StopBMD1.Enabled = False
StopBMD2.Enabled = False
StopBMD3.Enabled = False
StopBMD4.Enabled = False
StopBMD5.Enabled = False
StopBMD6.Enabled = False
UndoLast1.Enabled = False
UndoLast2.Enabled = False
UndoLast3.Enabled = False
UndoLast4.Enabled = False
UndoLast5.Enabled = False
UndoLast6.Enabled = False
Help1.Enabled = False
Help2.Enabled = False
Help3.Enabled = False
Help4.Enabled = False
Help5.Enabled = False
Help6.Enabled = False
ActiveSheet.Cells(1, 1) = "BMD1_Start"
ActiveSheet.Cells(1, 2) = "BMD1_Stop"
ActiveSheet.Cells(1, 3) = "BMD1_Duration"
ActiveSheet.Cells(1, 4) = "BMD1_Help"
ActiveSheet.Cells(1, 5) = "BMD2_Start"
ActiveSheet.Cells(1, 6) = "BMD2_Stop"
ActiveSheet.Cells(1, 7) = "BMD2_Duration"
ActiveSheet.Cells(1, 8) = "BMD2_Help"
ActiveSheet.Cells(1, 9) = "BMD3_Start"
ActiveSheet.Cells(1, 10) = "BMD3_Stop"
ActiveSheet.Cells(1, 11) = "BMD3_Duration"
ActiveSheet.Cells(1, 12) = "BMD3_Help"
ActiveSheet.Cells(1, 13) = "BMD4_Start"
ActiveSheet.Cells(1, 14) = "BMD4_Stop"
ActiveSheet.Cells(1, 15) = "BMD4_Duration"
ActiveSheet.Cells(1, 16) = "BMD4_Help"
ActiveSheet.Cells(1, 17) = "BMD5_Start"
ActiveSheet.Cells(1, 18) = "BMD5_Stop"
ActiveSheet.Cells(1, 19) = "BMD5_Duration"
ActiveSheet.Cells(1, 20) = "BMD5_Help"
ActiveSheet.Cells(1, 21) = "BMD6_Start"
ActiveSheet.Cells(1, 22) = "BMD6_Stop"
ActiveSheet.Cells(1, 23) = "BMD6_Duration"
ActiveSheet.Cells(1, 24) = "BMD6_Help"
ActiveSheet.Cells(1, 25) = "Comments"
ActiveSheet.Range("C:C").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("G:G").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("K:K").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("O:O").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("S:S").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("W:W").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("D:D").NumberFormat = Text
ActiveSheet.Range("H:H").NumberFormat = Text
ActiveSheet.Range("L:L").NumberFormat = Text
ActiveSheet.Range("P:P").NumberFormat = Text
ActiveSheet.Range("T:T").NumberFormat = Text
ActiveSheet.Range("X:X").NumberFormat = Text
ActiveSheet.columns("A:AA").AutoFit
ActiveSheet.Range("A1", "AA1").Font.Bold = True

End Sub
