VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckInForm 
   Caption         =   "Check In Timer"
   ClientHeight    =   10860
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   17980
   OleObjectBlob   =   "CheckInForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckInForm"
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

Private Sub StartPad1_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 1) = time
StopPad1.Enabled = True
StartPad1.Enabled = False
Image1.BorderColor = &HFF00&
UndoLast1.Enabled = True
VBM1.Enabled = True
StartProv1.Enabled = True
EndProv1.Enabled = True
StartPad1.BackColor = &HFF00&
End Sub

Private Sub StartPad2_Click()
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 5) = time
StopPad2.Enabled = True
StartPad2.Enabled = False
Image2.BorderColor = &HFF00&
UndoLast2.Enabled = True
VBM2.Enabled = True
StartProv2.Enabled = True
StartPad2.BackColor = &HFF00&
EndProv2.Enabled = True
End Sub

Private Sub StartPad3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 9) = time
StopPad3.Enabled = True
StartPad3.Enabled = False
Image3.BorderColor = &HFF00&
UndoLast3.Enabled = True
VBM3.Enabled = True
StartProv3.Enabled = True
EndProv3.Enabled = True
StartPad3.BackColor = &HFF00&
End Sub

Private Sub StartPad4_Click()
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 13) = time
StopPad4.Enabled = True
StartPad4.Enabled = False
Image4.BorderColor = &HFF00&
UndoLast4.Enabled = True
VBM4.Enabled = True
StartProv4.Enabled = True
EndProv4.Enabled = True
StartPad4.BackColor = &HFF00&
End Sub


Private Sub StopPad1_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
ActiveSheet.Cells(nr, 2) = time
ActiveSheet.Cells(nr, 3) = ActiveSheet.Cells(nr, 2) - ActiveSheet.Cells(nr, 1)
StopPad1.Enabled = False
StartPad1.Enabled = True
Image1.BorderColor = &H80000011
TextBox1 = ""
VBM1.Enabled = False
StartProv1.Enabled = False
EndProv1.Enabled = False
StartPad1.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 4) = "") Then
ActiveSheet.Cells(nr, 4) = "Normal"
Else: End If
End Sub

Private Sub StopPad2_Click()
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row
ActiveSheet.Cells(nr, 6) = time
ActiveSheet.Cells(nr, 7) = ActiveSheet.Cells(nr, 6) - ActiveSheet.Cells(nr, 5)
StopPad2.Enabled = False
StartPad2.Enabled = True
Image2.BorderColor = &H80000011
TextBox2 = ""
VBM2.Enabled = False
StartProv2.Enabled = False
EndProv2.Enabled = False
StartPad2.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 8) = "") Then
ActiveSheet.Cells(nr, 8) = "Normal"
Else: End If
End Sub

Private Sub StopPad3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10) = time
ActiveSheet.Cells(nr, 11) = ActiveSheet.Cells(nr, 10) - ActiveSheet.Cells(nr, 9)
StopPad3.Enabled = False
StartPad3.Enabled = True
Image3.BorderColor = &H80000011
TextBox3 = ""
VBM3.Enabled = False
StartProv3.Enabled = False
EndProv3.Enabled = False
StartPad3.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 12) = "") Then
ActiveSheet.Cells(nr, 12) = "Normal"
Else: End If
End Sub

Private Sub StopPad4_Click()
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row
ActiveSheet.Cells(nr, 14) = time
ActiveSheet.Cells(nr, 15) = ActiveSheet.Cells(nr, 14) - ActiveSheet.Cells(nr, 13)
StopPad4.Enabled = False
StartPad4.Enabled = True
Image4.BorderColor = &H80000011
TextBox4 = ""
VBM4.Enabled = False
StartProv4.Enabled = False
EndProv4.Enabled = False
StartPad4.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 16) = "") Then
ActiveSheet.Cells(nr, 16) = "Normal"
Else: End If
End Sub



Private Sub UndoLast1_Click()

If (ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
ActiveSheet.Cells(nr, 2).Clear
ActiveSheet.Cells(nr, 1).Clear
ActiveSheet.Cells(nr, 3).Clear
ActiveSheet.Cells(nr, 4).Clear
TextBox1 = ""
StopPad1.Enabled = False
StartPad1.Enabled = True
Image1.BorderColor = &H80000011
StartPad1.BackColor = &H8000000F
UndoLast1.Enabled = False
VBM1.Enabled = False
StartProv1.Enabled = False
EndProv1.Enabled = False
Else: End If



End Sub

Private Sub UndoLast2_Click()

If (ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row
ActiveSheet.Cells(nr, 5).Clear
ActiveSheet.Cells(nr, 8).Clear
ActiveSheet.Cells(nr, 6).Clear
ActiveSheet.Cells(nr, 7).Clear
TextBox2 = ""
StopPad2.Enabled = False
StartPad2.Enabled = True
Image2.BorderColor = &H80000011
StartPad2.BackColor = &H8000000F
UndoLast2.Enabled = False
VBM2.Enabled = False
StartProv2.Enabled = False
EndProv2.Enabled = False
Else: End If

End Sub

Private Sub UndoLast3_Click()

If (ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 11).Clear
ActiveSheet.Cells(nr, 10).Clear
ActiveSheet.Cells(nr, 9).Clear
ActiveSheet.Cells(nr, 12).Clear
TextBox3 = ""
StopPad3.Enabled = False
StartPad3.Enabled = True
Image3.BorderColor = &H80000011
StartPad3.BackColor = &H8000000F
UndoLast3.Enabled = False
VBM3.Enabled = False
StartProv3.Enabled = False
EndProv3.Enabled = False
Else: End If

End Sub

Private Sub UndoLast4_Click()

If (ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 13).End(xlUp).Row
ActiveSheet.Cells(nr, 13).Clear
ActiveSheet.Cells(nr, 14).Clear
ActiveSheet.Cells(nr, 15).Clear
ActiveSheet.Cells(nr, 16).Clear
TextBox4 = ""
StopPad4.Enabled = False
StartPad4.Enabled = True
Image4.BorderColor = &H80000011
StartPad4.BackColor = &H8000000F
UndoLast4.Enabled = False
VBM4.Enabled = False
StartProv4.Enabled = False
EndProv4.Enabled = False
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

Private Sub StopPad5_Click()
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row
ActiveSheet.Cells(nr, 18) = time
ActiveSheet.Cells(nr, 19) = ActiveSheet.Cells(nr, 18) - ActiveSheet.Cells(nr, 17)
StopPad5.Enabled = False
StartPad5.Enabled = True
Image5.BorderColor = &H80000011
TextBox5 = ""
VBM5.Enabled = False
StartProv5.Enabled = False
EndProv5.Enabled = False
StartPad5.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 20) = "") Then
ActiveSheet.Cells(nr, 20) = "Normal"
Else: End If
End Sub
Private Sub StopPad6_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row
ActiveSheet.Cells(nr, 22) = time
ActiveSheet.Cells(nr, 23) = ActiveSheet.Cells(nr, 22) - ActiveSheet.Cells(nr, 21)
StopPad6.Enabled = False
StartPad6.Enabled = True
Image6.BorderColor = &H80000011
TextBox6 = ""
VBM6.Enabled = False
StartProv6.Enabled = False
EndProv6.Enabled = False
StartPad6.BackColor = &H8000000F
If (ActiveSheet.Cells(nr, 24) = "") Then
ActiveSheet.Cells(nr, 24) = "Normal"
Else: End If
End Sub


Private Sub UndoLast5_Click()

If (ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row
ActiveSheet.Cells(nr, 17).Clear
ActiveSheet.Cells(nr, 18).Clear
ActiveSheet.Cells(nr, 19).Clear
ActiveSheet.Cells(nr, 20).Clear
TextBox5 = ""
StopPad5.Enabled = False
StartPad5.Enabled = True
Image5.BorderColor = &H80000011
StartPad5.BackColor = &H8000000F
UndoLast5.Enabled = False
VBM5.Enabled = False
StartProv5.Enabled = False
EndProv5.Enabled = False
Else: End If

End Sub
Private Sub UndoLast6_Click()

If (ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row
ActiveSheet.Cells(nr, 21).Clear
ActiveSheet.Cells(nr, 22).Clear
ActiveSheet.Cells(nr, 23).Clear
ActiveSheet.Cells(nr, 24).Clear
TextBox6 = ""
StopPad6.Enabled = False
StartPad6.Enabled = True
Image6.BorderColor = &H80000011
StartPad6.BackColor = &H8000000F
UndoLast6.Enabled = False
VBM6.Enabled = False
StartProv6.Enabled = False
EndProv6.Enabled = False
Else: End If

End Sub

Private Sub StartPad5_Click()
nr = ActiveSheet.Cells(Rows.count, 17).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 17) = time
StopPad5.Enabled = True
StartPad5.Enabled = False
Image5.BorderColor = &HFF00&
UndoLast5.Enabled = True
VBM5.Enabled = True
StartProv5.Enabled = True
EndProv5.Enabled = True
StartPad5.BackColor = &HFF00&
End Sub
Private Sub StartPad6_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 21) = time
StopPad6.Enabled = True
StartPad6.Enabled = False
Image6.BorderColor = &HFF00&
UndoLast6.Enabled = True
VBM6.Enabled = True
StartProv6.Enabled = True
EndProv6.Enabled = True
StartPad6.BackColor = &HFF00&
End Sub

Private Sub SaveComment_Click()
nr = ActiveSheet.Cells(Rows.count, 25).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 25) = CheckInForm.CommentBox.Value
CheckInForm.CommentBox.Value = ""

End Sub

Private Sub UserForm_Initialize()
StopPad1.Enabled = False
StopPad2.Enabled = False
StopPad3.Enabled = False
StopPad4.Enabled = False
StopPad5.Enabled = False
StopPad6.Enabled = False
UndoLast1.Enabled = False
UndoLast2.Enabled = False
UndoLast3.Enabled = False
UndoLast4.Enabled = False
UndoLast5.Enabled = False
UndoLast6.Enabled = False
VBM1.Enabled = False
VBM2.Enabled = False
VBM3.Enabled = False
VBM4.Enabled = False
VBM5.Enabled = False
VBM6.Enabled = False
StartProv1.Enabled = False
StartProv2.Enabled = False
StartProv3.Enabled = False
StartProv4.Enabled = False
StartProv5.Enabled = False
StartProv6.Enabled = False
EndProv1.Enabled = False
EndProv2.Enabled = False
EndProv3.Enabled = False
EndProv4.Enabled = False
EndProv5.Enabled = False
EndProv6.Enabled = False
ActiveSheet.Cells(1, 1) = "CheckIn1_Start"
ActiveSheet.Cells(1, 2) = "CheckIn1_Stop"
ActiveSheet.Cells(1, 3) = "CheckIn1_Duration"
ActiveSheet.Cells(1, 4) = "CheckIn1_Type"
ActiveSheet.Cells(1, 5) = "CheckIn2_Start"
ActiveSheet.Cells(1, 6) = "CheckIn2_Stop"
ActiveSheet.Cells(1, 7) = "CheckIn2_Duration"
ActiveSheet.Cells(1, 8) = "CheckIn2_Type"
ActiveSheet.Cells(1, 9) = "CheckIn3_Start"
ActiveSheet.Cells(1, 10) = "CheckIn3_Stop"
ActiveSheet.Cells(1, 11) = "CheckIn3_Duration"
ActiveSheet.Cells(1, 12) = "CheckIn3_Type"
ActiveSheet.Cells(1, 13) = "CheckIn4_Start"
ActiveSheet.Cells(1, 14) = "CheckIn4_Stop"
ActiveSheet.Cells(1, 15) = "CheckIn4_Duration"
ActiveSheet.Cells(1, 16) = "CheckIn4_Type"
ActiveSheet.Cells(1, 17) = "CheckIn5_Start"
ActiveSheet.Cells(1, 18) = "CheckIn5_Stop"
ActiveSheet.Cells(1, 19) = "CheckIn5_Duration"
ActiveSheet.Cells(1, 20) = "CheckIn5_Type"
ActiveSheet.Cells(1, 21) = "CheckIn6_Start"
ActiveSheet.Cells(1, 22) = "CheckIn6_Stop"
ActiveSheet.Cells(1, 23) = "CheckIn6_Duration"
ActiveSheet.Cells(1, 24) = "CheckIn6_Type"
ActiveSheet.Cells(1, 25) = "Comments"
ActiveSheet.Range("C:C").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("G:G").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("K:K").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("O:O").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("S:S").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("W:W").NumberFormat = "hh:mm:ss"
ActiveSheet.columns("A:AC").AutoFit
ActiveSheet.Range("A1", "AC1").Font.Bold = True


End Sub

Private Sub VBM1_Click()
nr = ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 4) = "VBM"
VBM1.Enabled = False
StartProv1.Enabled = False
EndProv1.Enabled = False
End Sub
Private Sub VBM2_Click()
nr = ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 8) = "VBM"
VBM2.Enabled = False
StartProv2.Enabled = False
EndProv2.Enabled = False
End Sub
Private Sub VBM3_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = "VBM"
VBM3.Enabled = False
StartProv3.Enabled = False
EndProv3.Enabled = False
End Sub
Private Sub VBM4_Click()
nr = ActiveSheet.Cells(Rows.count, 16).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 16) = "VBM"
VBM4.Enabled = False
StartProv4.Enabled = False
EndProv4.Enabled = False
End Sub
Private Sub VBM5_Click()
nr = ActiveSheet.Cells(Rows.count, 20).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 20) = "VBM"
VBM5.Enabled = False
StartProv5.Enabled = False
EndProv5.Enabled = False
End Sub
Private Sub VBM6_Click()
nr = ActiveSheet.Cells(Rows.count, 24).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 24) = "VBM"
VBM6.Enabled = False
StartProv6.Enabled = False
EndProv6.Enabled = False
End Sub
Private Sub StartProv1_Click()
nr = ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 4) = "Given Provisional"
VBM1.Enabled = False
StartProv1.Enabled = False
EndProv1.Enabled = False
End Sub
Private Sub StartProv2_Click()
nr = ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 8) = "Given Provisional"
VBM2.Enabled = False
StartProv2.Enabled = False
EndProv2.Enabled = False
End Sub
Private Sub StartProv3_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = "Given Provisional"
VBM3.Enabled = False
StartProv3.Enabled = False
EndProv3.Enabled = False
End Sub
Private Sub StartProv4_Click()
nr = ActiveSheet.Cells(Rows.count, 16).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 16) = "Given Provisional"
VBM4.Enabled = False
StartProv4.Enabled = False
EndProv4.Enabled = False
End Sub
Private Sub StartProv5_Click()
nr = ActiveSheet.Cells(Rows.count, 20).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 20) = "Given Provisional"
VBM5.Enabled = False
StartProv5.Enabled = False
EndProv5.Enabled = False
End Sub
Private Sub StartProv6_Click()
nr = ActiveSheet.Cells(Rows.count, 24).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 24) = "Given Provisional"
VBM6.Enabled = False
StartProv6.Enabled = False
EndProv6.Enabled = False
End Sub
Private Sub EndProv1_Click()
nr = ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 4) = "Returned Provisional"
VBM1.Enabled = False
StartProv1.Enabled = False
EndProv1.Enabled = False
End Sub
Private Sub EndProv2_Click()
nr = ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 8) = "Returned Provisional"
VBM2.Enabled = False
StartProv2.Enabled = False
EndProv2.Enabled = False
End Sub
Private Sub EndProv3_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = "Returned Provisional"
VBM3.Enabled = False
StartProv3.Enabled = False
EndProv3.Enabled = False
End Sub
Private Sub EndProv4_Click()
nr = ActiveSheet.Cells(Rows.count, 16).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 16) = "Returned Provisional"
VBM4.Enabled = False
StartProv4.Enabled = False
EndProv4.Enabled = False
End Sub
Private Sub EndProv5_Click()
nr = ActiveSheet.Cells(Rows.count, 20).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 20) = "Returned Provisional"
VBM5.Enabled = False
StartProv5.Enabled = False
EndProv5.Enabled = False
End Sub
Private Sub EndProv6_Click()
nr = ActiveSheet.Cells(Rows.count, 24).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 26) = "Returned Provisional"
VBM6.Enabled = False
StartProv6.Enabled = False
EndProv6.Enabled = False
End Sub
