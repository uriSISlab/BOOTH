VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ThroughputArrivalForm 
   Caption         =   "Throughput and Arrival Timer"
   ClientHeight    =   10640
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   17980
   OleObjectBlob   =   "ThroughputArrivalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ThroughputArrivalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Arrival_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 1) = time
ArriveCount.Caption = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row - 1
UndoLastArrive.Enabled = True
End Sub

Private Sub Clear1_Click()
TextBox1 = ""
End Sub


Private Sub SaveButton_Click()
    ActiveWorkbook.Save
End Sub

Private Sub ThroughputStart1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 3) = time
ThroughputStop1.Enabled = True
ThroughputStart1.Enabled = False
Image1.BorderColor = &HFF00&
UndoLast1.Enabled = True
ThroughputStart1.BackColor = &HFF00&
End Sub

Private Sub ThroughputStart2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 6) = time
ThroughputStop2.Enabled = True
ThroughputStart2.Enabled = False
Image2.BorderColor = &HFF00&
UndoLast2.Enabled = True
ThroughputStart2.BackColor = &HFF00&
End Sub

Private Sub ThroughputStart3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 9) = time
ThroughputStop3.Enabled = True
ThroughputStart3.Enabled = False
Image3.BorderColor = &HFF00&
UndoLast3.Enabled = True
ThroughputStart3.BackColor = &HFF00&
End Sub

Private Sub ThroughputStart4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 12) = time
ThroughputStop4.Enabled = True
ThroughputStart4.Enabled = False
Image4.BorderColor = &HFF00&
UndoLast4.Enabled = True
ThroughputStart4.BackColor = &HFF00&
End Sub


Private Sub ThroughputStop1_Click()
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4) = time
ActiveSheet.Cells(nr, 5) = ActiveSheet.Cells(nr, 4) - ActiveSheet.Cells(nr, 3)
ThroughputStop1.Enabled = False
ThroughputStart1.Enabled = True
Image1.BorderColor = &H80000011
TextBox1 = ""
ThroughputStart1.BackColor = &H8000000F
End Sub

Private Sub ThroughputStop2_Click()
nr = ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row
ActiveSheet.Cells(nr, 7) = time
ActiveSheet.Cells(nr, 8) = ActiveSheet.Cells(nr, 7) - ActiveSheet.Cells(nr, 6)
ThroughputStop2.Enabled = False
ThroughputStart2.Enabled = True
Image2.BorderColor = &H80000011
TextBox2 = ""
ThroughputStart2.BackColor = &H8000000F
End Sub

Private Sub ThroughputStop3_Click()
nr = ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row
ActiveSheet.Cells(nr, 10) = time
ActiveSheet.Cells(nr, 11) = ActiveSheet.Cells(nr, 10) - ActiveSheet.Cells(nr, 9)
ThroughputStop3.Enabled = False
ThroughputStart3.Enabled = True
Image3.BorderColor = &H80000011
TextBox3 = ""
ThroughputStart3.BackColor = &H8000000F
End Sub

Private Sub ThroughputStop4_Click()
nr = ActiveSheet.Cells(Rows.count, 12).End(xlUp).Row
ActiveSheet.Cells(nr, 13) = time
ActiveSheet.Cells(nr, 14) = ActiveSheet.Cells(nr, 13) - ActiveSheet.Cells(nr, 12)
ThroughputStop4.Enabled = False
ThroughputStart4.Enabled = True
Image4.BorderColor = &H80000011
TextBox4 = ""
ThroughputStart4.BackColor = &H8000000F
End Sub



Private Sub UndoLast1_Click()

If (ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row
ActiveSheet.Cells(nr, 4).Clear
ActiveSheet.Cells(nr, 3).Clear
ActiveSheet.Cells(nr, 5).Clear
TextBox1 = ""
ThroughputStop1.Enabled = False
ThroughputStart1.Enabled = True
Image1.BorderColor = &H80000011
ThroughputStart1.BackColor = &H8000000F
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
ThroughputStop2.Enabled = False
ThroughputStart2.Enabled = True
Image2.BorderColor = &H80000011
ThroughputStart2.BackColor = &H8000000F
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
ThroughputStop3.Enabled = False
ThroughputStart3.Enabled = True
Image3.BorderColor = &H80000011
ThroughputStart3.BackColor = &H8000000F
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
ThroughputStop4.Enabled = False
ThroughputStart4.Enabled = True
Image4.BorderColor = &H80000011
ThroughputStart4.BackColor = &H8000000F
UndoLast4.Enabled = False
Else: End If

End Sub

Private Sub UndoLastArrive_Click()
If (ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
ActiveSheet.Cells(nr, 1).Clear
ActiveSheet.Cells(nr, 2).Clear
UndoLastArrive.Enabled = False
ArriveCount.Caption = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row - 1
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

Private Sub ThroughputStop5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16) = time
ActiveSheet.Cells(nr, 17) = ActiveSheet.Cells(nr, 16) - ActiveSheet.Cells(nr, 15)
ThroughputStop5.Enabled = False
ThroughputStart5.Enabled = True
Image5.BorderColor = &H80000011
TextBox5 = ""
ThroughputStart5.BackColor = &H8000000F
End Sub
Private Sub UndoLast5_Click()

If (ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row > 1) Then
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row
ActiveSheet.Cells(nr, 16).Clear
ActiveSheet.Cells(nr, 15).Clear
ActiveSheet.Cells(nr, 17).Clear
TextBox5 = ""
ThroughputStop5.Enabled = False
ThroughputStart5.Enabled = True
Image5.BorderColor = &H80000011
ThroughputStart5.BackColor = &H8000000F
UndoLast5.Enabled = False
Else: End If

End Sub
Private Sub ThroughputStart5_Click()
nr = ActiveSheet.Cells(Rows.count, 15).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 15) = time
ThroughputStop5.Enabled = True
ThroughputStart5.Enabled = False
Image5.BorderColor = &HFF00&
UndoLast5.Enabled = True
ThroughputStart5.BackColor = &HFF00&
End Sub

Private Sub SaveComment_Click()
nr = ActiveSheet.Cells(Rows.count, 21).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 21) = ThroughputArrivalForm.CommentBox.Value
ThroughputArrivalForm.CommentBox.Value = ""

End Sub

Private Sub UserForm_Initialize()
ThroughputStop1.Enabled = False
ThroughputStop2.Enabled = False
ThroughputStop3.Enabled = False
ThroughputStop4.Enabled = False
ThroughputStop5.Enabled = False
UndoLast1.Enabled = False
UndoLast2.Enabled = False
UndoLast3.Enabled = False
UndoLast4.Enabled = False
UndoLast5.Enabled = False
UndoLastArrive.Enabled = False
ActiveSheet.Cells(1, 1) = "Arrival_Time"
ActiveSheet.Cells(1, 2) = "Arrival_Type"
ActiveSheet.Cells(1, 3) = "Throughput1_Start"
ActiveSheet.Cells(1, 4) = "Throughput1_Stop"
ActiveSheet.Cells(1, 5) = "Throughput1_Duration"
ActiveSheet.Cells(1, 6) = "Throughput2_Start"
ActiveSheet.Cells(1, 7) = "Throughput2_Stop"
ActiveSheet.Cells(1, 8) = "Throughput2_Duration"
ActiveSheet.Cells(1, 9) = "Throughput3_Start"
ActiveSheet.Cells(1, 10) = "Throughput3_Stop"
ActiveSheet.Cells(1, 11) = "Throughput3_Duration"
ActiveSheet.Cells(1, 12) = "Throughput4_Start"
ActiveSheet.Cells(1, 13) = "Throughput4_Stop"
ActiveSheet.Cells(1, 14) = "Throughput4_Duration"
ActiveSheet.Cells(1, 15) = "Throughput5_Start"
ActiveSheet.Cells(1, 16) = "Throughput5_Stop"
ActiveSheet.Cells(1, 17) = "Throughput5_Duration"
ActiveSheet.Cells(1, 18) = "Throughput6_Start"
ActiveSheet.Cells(1, 19) = "Throughput6_Stop"
ActiveSheet.Cells(1, 20) = "Throughput6_Duration"
ActiveSheet.Cells(1, 21) = "Comments"
ActiveSheet.Range("E:E").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("H:H").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("K:K").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("N:N").NumberFormat = "hh:mm:ss"
ActiveSheet.Range("Q:Q").NumberFormat = "hh:mm:ss"
ActiveSheet.columns("A:AC").AutoFit
ActiveSheet.Range("A1", "AC1").Font.Bold = True


ArriveCount.Caption = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row - 1
End Sub


Private Sub VBMArrive_Click()
nr = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
ActiveSheet.Cells(nr, 1) = time
ActiveSheet.Cells(nr, 2) = "VBM"
ArriveCount.Caption = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row - 1
UndoLastArrive.Enabled = True
End Sub
