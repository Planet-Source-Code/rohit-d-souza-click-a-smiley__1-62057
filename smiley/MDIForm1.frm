VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Clik A Smiley"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnugamemenu 
      Caption         =   "Game Menu"
      Begin VB.Menu mnunewgame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuaboutsmiley 
         Caption         =   "About Smiley"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub mnuaboutsmiley_Click()
frmAbout.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End
End Sub

Private Sub mnuHowtoplay_Click()

End Sub

Private Sub mnunewgame_Click()
If Form2.Visible = True Then Unload Form2
If Form1.Visible = True Then Unload Form1
Form2.Show
End Sub

Private Sub mnyhelpaboutsmiley_Click()
frmAbout.Show
End Sub


