VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "About The Game"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Play Game"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Hard"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Meduim"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Easy"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Developed By Rohit D'souza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Select Difficulty Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Welcome To Click a Smiley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public interval1 As Integer
Private Sub Command1_Click()
If Option1.Value = True Then
interval1 = 1000
start
ElseIf Option2.Value = True Then
interval1 = 500
start
ElseIf Option3.Value = True Then
interval1 = 300
start
End If


End Sub
Sub start()
Form2.Hide
Form1.timestart interval1
Form1.Show
End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

