VERSION 5.00
Begin VB.Form TestFrm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "mbBrown"
      Height          =   375
      Index           =   9
      Left            =   45
      MaskColor       =   &H0000FFFF&
      TabIndex        =   10
      Top             =   1575
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbYellow"
      Height          =   375
      Index           =   8
      Left            =   2295
      MaskColor       =   &H0000FFFF&
      TabIndex        =   9
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbBlue"
      Height          =   375
      Index           =   7
      Left            =   1170
      TabIndex        =   8
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbPink"
      Height          =   375
      Index           =   6
      Left            =   45
      TabIndex        =   7
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbPurple"
      Height          =   375
      Index           =   5
      Left            =   2295
      TabIndex        =   5
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbWinter"
      Height          =   375
      Index           =   4
      Left            =   1170
      TabIndex        =   4
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbSummer"
      Height          =   375
      Index           =   3
      Left            =   45
      TabIndex        =   3
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbSpring"
      Height          =   375
      Index           =   2
      Left            =   2295
      TabIndex        =   2
      Top             =   225
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbAutumn"
      Height          =   375
      Index           =   1
      Left            =   1170
      TabIndex        =   1
      Top             =   225
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "mbFlat"
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   225
      Width           =   1050
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   315
      TabIndex        =   6
      Top             =   2385
      Width           =   3120
   End
End
Attribute VB_Name = "TestFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
Msbox "This is the flat messagebox, also shown as the default." & vbCr & "Just enter text, and nothing else..." & vbCr & vbCr & "Do you see ?"
Case 1
Msbox "This is the style 'mbAutumn'", mbAutumn, mbEnterLeave, mbNoEntry, "Show it to me !"
Case 2
Msbox "Here you can see the 'mbSpring' messagebox", mbSpring, mbYesNo, mbQuestion, "Show me what you got"
Case 3
Msbox "And 'mbSummer'" & vbCr & "There are up to 10 different styles", mbSummer, mbOkOnly, mbCritical, "Amazing !"
Case 4
Msbox "When it's cold outside... mbWinter ! ", mbWinter, mbYesNoCancel, mbPrint, "This is great !"
Case 5
Msbox "Purple...", mbPurple, mbLoadDontLoad, mbExclamation, "So many colors !"
Case 6
Msbox "And 'mbPink'", mbPink, mbExitNoWay, mbInfo, "I am convinced..."
Case 7
Msbox "A bit dark, but still good" & vbCr & "The 'mbBlue' messagebox'", mbBlue, mbIAgreeDontAgree, mbTrash, "Stop it ! That's too much !"
Case 8
Msbox "Yellow", mbYellow, mbPrintDontPrint, mbPrint, "Print ? What is there to print ?"
Case 9
Msbox "mmmm... yeah... not bad ! ", mbBrown, mbAgreeOnly, mbOpen, "A bit brown, don't you think ?"
End Select
Label1.Caption = "You clicked button " & mbReturn
End Sub

Private Sub Form_Load()
TestFrm.Move 0, 0
End Sub
