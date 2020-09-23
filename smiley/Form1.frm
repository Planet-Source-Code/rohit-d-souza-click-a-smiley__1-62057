VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   4680
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtscore 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdreset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1080
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   480
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   3600
      Picture         =   "Form1.frx":044E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2160
      Picture         =   "Form1.frx":0890
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "Form1.frx":0CD2
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2400
      Picture         =   "Form1.frx":1114
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":1556
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Your Score"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   -3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Starting game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdreset_Click()
Timer1.Enabled = False
reset
End Sub

Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
txtscore = 0
For i = 0 To Picture1.count - 1 'to initally make smileys invisible
Picture1(i).Left = -1000
Picture1(i).Top = -1000
Next
Label2.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub Picture1_Click(Index As Integer)
Dim count As Integer 'to keep account of no of smileys clicked
Select Case Index
Case 0
Picture1(0).Visible = False
count = txtscore + 1
txtscore = count
Case 1
Picture1(1).Visible = False
count = txtscore + 1
txtscore = count
Case 2
Picture1(2).Visible = False
count = txtscore + 1
txtscore = count
Case 3
Picture1(3).Visible = False
count = txtscore + 1
txtscore = count
Case 4
Picture1(4).Visible = False
count = txtscore + 1
txtscore = count
Case 5
Picture1(5).Visible = False
count = txtscore + 1
txtscore = count
End Select
End Sub

Private Sub Timer1_Timer()
Label2.Visible = False
Picture1(0).Left = Rnd() * 1200
Picture1(0).Top = Rnd() * 1200
Picture1(1).Left = Rnd(2) * 2000
Picture1(1).Top = Rnd(2) * 3000
Picture1(2).Left = Rnd(3) * 10000
Picture1(2).Top = Rnd(3) * 10000
Picture1(3).Left = Rnd(4) * 12000
Picture1(3).Top = Rnd(4) * 12000
Picture1(4).Left = Rnd(5) * 14000
Picture1(4).Top = Rnd(5) * 14000
Picture1(5).Left = Rnd(6) * 16000
Picture1(5).Top = Rnd(6) * 3000

End Sub

Private Sub txtscore_Change()
If txtscore.Text = 6 Then
Timer1.Enabled = False
If MsgBox("Good U Clicked On All The Smiley's do u want to play another game ", vbYesNo) = vbYes Then
reset
Unload Me
Form2.Show
Else
End
End If
End If
End Sub
Sub timestart(inter As Integer) 'to get interval according to level selected
Timer1.Interval = inter
End Sub
Sub reset()
Dim i As Integer
txtscore.Text = 0
Timer1.Enabled = True
For i = 0 To Picture1.count - 1
Picture1(i).Visible = True
Next
End Sub
