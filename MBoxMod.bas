Attribute VB_Name = "MBoxMod"
Public mbReturn%

Public Enum MBoxButtons
 mbOkOnly
 mbAgreeOnly
 mbOKNoWay
 mbOKCancel
 mbYesNo
 mbYesNoCancel
 mbExitNoWay
 mbExitNoWayCancel
 mbSaveDontSave
 mbLoadDontLoad
 mbPrintDontPrint
 mbEnterLeave
 mbEnterLeaveCancel
 mbIAgreeDontAgree
 End Enum

Public Enum MBoxStyle
mbFlat
mbAutumn
mbSpring
mbSummer
mbWinter
mbPurple
mbPink
mbBlue
mbYellow
mbBrown
End Enum

Public Enum IconValue
 mbNoIcon
 mbQuestion
 mbInfo
 mbNoEntry
 mbExclamation
 mbSave
 mbOpen
 mbPrint
 mbCritical
 mbTrash
End Enum

Public Function Msbox(Message As Variant, Optional Style As MBoxStyle, Optional Buttons As MBoxButtons, Optional MBoxIcon As IconValue, Optional Title As Variant)
Dim xx%, yy%, R1, R2, G1, G2, B1, B2, Rs, Gs, Bs, Rx, Gx, Bx
Dim LCol1&, LCol2&, Border1&, Border2&

On Error Resume Next
If IsMissing(Title) Then Title = App.Title
'***  set default
With MBox
    MBox.ScaleMode = 3
    MBox.Width = 4605
    MBox.Height = 2190
    .Label3.Top = 2
    .Pic1.Move 3, 3, MBox.ScaleWidth - 6, 18
    .Label2.Caption = ""
    .Label2.Move 9, 27, 249, 16
    
'*** set style-colors
If Style = mbFlat Then
MBox.BackColor = &HC0C0C0
R1 = &H80: R2 = &H80
G1 = &H80: G2 = &H80
B1 = &H80: B2 = &H80
LCol1 = &HFFFFFF
Border1 = &HE0E0E0
Border2 = &H606060
GoTo SkipColors
End If

If Style = mbAutumn Then
MBox.BackColor = &HA0C0C0
R1 = &HA0: R2 = &H60
G1 = &HA0: G2 = &H60
B1 = &H80: B2 = &H40
LCol1 = &HC0F0F0
LCol2 = &H3F0F0F
Border1 = &HB0D0D0
Border2 = &H608080
End If
If Style = mbSpring Then
MBox.BackColor = &H80C0A0
R1 = &H80: R2 = &H20
G1 = &HA0: G2 = &H40
B1 = &H60: B2 = &H0
LCol1 = &H80F0C0
LCol2 = &H601040
Border1 = &H90D0B0
Border2 = &H406060
End If
If Style = mbSummer Then
MBox.BackColor = &H2060E0
R1 = &HD0: R2 = &H40
G1 = &H60: G2 = &H0
B1 = &H40: B2 = &H0
LCol1 = &H40A0F0
LCol2 = &H103020
Border1 = &H7090D0
Border2 = &H102040
End If
If Style = mbWinter Then
MBox.BackColor = &HE0A0A0
R1 = &HF0: R2 = &H60
G1 = &HF0: G2 = &H60
B1 = &HF0: B2 = &H80
LCol1 = &H804040
LCol2 = &H403000
Border1 = &HF08080
Border2 = &H802040
End If
If Style = mbPurple Then
MBox.BackColor = &HA090A0
R1 = &HA0: R2 = &H40
G1 = &H90: G2 = &H0
B1 = &HA0: B2 = &H40
LCol1 = &HC0C0C0
LCol2 = &H604010
Border1 = &H908090
Border2 = &H402040
End If
If Style = mbPink Then
MBox.BackColor = &HC0C0F0
R1 = &HD0: R2 = &H80
G1 = &HA0: G2 = &H70
B1 = &HA0: B2 = &H70
LCol1 = &H202080
LCol2 = &H40
Border1 = &HC0C0F0
Border2 = &H8080A0
End If
If Style = mbBlue Then
MBox.BackColor = &HD05060
R2 = &H70: R1 = &H30
G2 = &H70: G1 = &H0
B2 = &HF0: B1 = &H20
LCol1 = &HFFFF00
LCol2 = &HC0C080
Border1 = &HD05050
Border2 = &H802020
End If
If Style = mbYellow Then
MBox.BackColor = &H60E0E0
R1 = &HF0: R2 = &HA0
G1 = &HF0: G2 = &HA0
B1 = &H80: B2 = &H20
LCol1 = &H206000
LCol2 = &H106040
Border1 = &H80FFFF
Border2 = &H208080
End If
If Style = mbBrown Then
MBox.BackColor = &H1060A0
R1 = &HF0: R2 = &H80
G1 = &HA0: G2 = &H30
B1 = &H60: B2 = &H20
LCol1 = &H20F0A0
LCol2 = &H10E040
Border1 = &H2080F0
Border2 = &H104080
End If
'*** set colors
SkipColors:
.Label3.ForeColor = LCol1
.Label1(0).ForeColor = LCol1
.Label1(1).ForeColor = LCol1
.Label1(2).ForeColor = LCol1
.Label2.ForeColor = LCol2

'*** set gradient
Rx = R1: Gx = G1: Bx = B1
Rs = (R1 - R2) / (.Pic1.ScaleHeight - 1)
Gs = (G1 - G2) / (.Pic1.ScaleHeight - 1)
Bs = (B1 - B2) / (.Pic1.ScaleHeight - 1)
    For xx = 0 To .Pic1.Height - 1
    .Pic1.Line (0, xx)-(.Pic1.Width, xx), RGB(Rx, Gx, Bx)
    For yy = 0 To 2
        .But1(yy).Line (0, xx)-(.But1(yy).Width, xx), RGB(Rx, Gx, Bx)
    Next yy
    Rx = Rx - Rs
    Gx = Gx - Gs
    Bx = Bx - Bs
Next xx

'*** setborders
MBox.Line (0, 0)-(MBox.ScaleWidth - 1, MBox.ScaleHeight - 1), Border1, B
MBox.Line (1, 1)-(MBox.ScaleWidth - 2, MBox.ScaleHeight - 2), Border2, B
For xx = 0 To 2
.But1(xx).Line (0, 0)-(.But1(xx).Width - 1, .But1(xx).Height - 1), Border1, B
.But1(xx).Line (0, .But1(xx).Height - 1)-(.But1(xx).Width, .But1(xx).Height - 1), Border2
.But1(xx).Line (.But1(xx).Width - 1, 0)-(.But1(xx).Width - 1, .But1(xx).Height - 1), Border2
Next xx

'*** set buttons
For xx = 0 To 2
.Label1(xx).Caption = ""
.Label1(xx).Move 0, 2
Next xx

If Buttons = mbOkOnly Then
   .Label1(0).Caption = "OK"
End If
If Buttons = mbAgreeOnly Then
.Label1(0).Caption = "Agree !"
End If
If Buttons = mbOKNoWay Then
   .Label1(0).Caption = "OK"
   .Label1(1).Caption = "No way !"
End If
If Buttons = mbOKCancel Then
   .Label1(0).Caption = "OK"
   .Label1(1).Caption = "Cancel"
End If
If Buttons = mbYesNo Then
   .Label1(0).Caption = "Yes"
   .Label1(1).Caption = "No"
End If
If Buttons = mbYesNoCancel Then
   .Label1(0).Caption = "Yes"
   .Label1(1).Caption = "No"
   .Label1(2).Caption = "Cancel"
End If
If Buttons = mbExitNoWay Then
   .Label1(0).Caption = "Exit"
   .Label1(1).Caption = "No way !"
End If
If Buttons = mbExitNoWayCancel Then
   .Label1(0).Caption = "Exit"
   .Label1(1).Caption = "No way !"
   .Label1(2).Caption = "Cancel"
End If
If Buttons = mbSaveDontSave Then
   .Label1(0).Caption = "Save"
   .Label1(1).Caption = "Don't Save"
End If
If Buttons = mbLoadDontLoad Then
   .Label1(0).Caption = "Load"
   .Label1(1).Caption = "Don't Load"
End If
If Buttons = mbPrintDontPrint Then
   .Label1(0).Caption = "Print"
   .Label1(1).Caption = "Don't Print"
End If
If Buttons = mbEnterLeave Then
   .Label1(0).Caption = "Enter"
   .Label1(1).Caption = "Leave"
End If
If Buttons = mbEnterLeaveCancel Then
   .Label1(0).Caption = "Enter"
   .Label1(1).Caption = "Leave"
   .Label1(2).Caption = "Cancel"
End If
If Buttons = mbIAgreeDontAgree Then
   .Label1(0).Caption = "Agree !"
   .Label1(1).Caption = "Don't Agree !"
End If

'***  position buttons
.But1(1).Visible = False
.But1(2).Visible = False
    If .Label1(1) = "" And .Label1(2) = "" Then
    .But1(0).Move (MBox.ScaleWidth / 2) - (.But1(0).Width / 2), 117
    End If
    If .Label1(1) <> "" And .Label1(2) = "" Then
    .But1(0).Move (MBox.ScaleWidth / 2) - (.But1(0).Width) - 6, 117
    .But1(1).Move (MBox.ScaleWidth / 2) + 6, 117
    .But1(1).Visible = True
    End If
    If .Label1(1) <> "" And .Label1(2) <> "" Then
    .But1(1).Move (MBox.ScaleWidth / 2) - (.But1(0).Width / 2), 117
    .But1(0).Move (MBox.ScaleWidth / 2) - (.But1(0).Width / 2) - .But1(0).Width - 6, 117
    .But1(2).Move (MBox.ScaleWidth / 2) + (.But1(0).Width / 2) + 6, 117
    .But1(1).Visible = True
    .But1(2).Visible = True
    End If
    
'*** set icon
.Image1.Picture = MBox.ImageList1.ListImages(MBoxIcon).Picture

'*** set text
.Label2.Caption = Message
If .Label2.Width > 249 Then Label2.Width = 249
If .Label2.Height > 78 Then .Label2.Height = 78
.Label2.Top = (MBox.ScaleHeight / 2) - (.Label2.Height / 2) - 5
MBox.Label3 = Title

End With
MBox.Show 1
End Function
