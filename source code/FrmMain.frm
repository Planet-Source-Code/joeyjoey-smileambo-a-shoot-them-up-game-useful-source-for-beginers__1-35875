VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Smileambo"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblNewgame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Smileambo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4020
   End
   Begin VB.Image ImgSmile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1725
      Left            =   0
      Picture         =   "FrmMain.frx":0CCA
      Top             =   240
      Width           =   4020
   End
   Begin VB.Label LblBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   4020
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstX As Integer
Dim FirstY As Integer



Private Sub LblAbout_Click()
Me.Hide 'hide main screen
Frmabout.Show 'show about box

End Sub

Private Sub LblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblAbout.BackColor = &HC0C0C0 ' change color

End Sub

Private Sub LblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblNewgame.BackColor = &H80000005 ' change color back to normal
LblQuit.BackColor = &H80000005 ' change color back to normal
LblAbout.BackColor = &H80000005 ' change color back to normal

End Sub

Private Sub LblNewgame_Click()
FrmGame.Show ' start a new game
Me.Hide ' hide main screen

End Sub

Private Sub LblNewgame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblNewgame.BackColor = &HC0C0C0 ' change color

End Sub

Private Sub LblQuit_Click()
Unload Me ' unload form from memory
End ' quit

End Sub

Private Sub LblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblQuit.BackColor = &HC0C0C0 ' change color

End Sub

Private Sub LblTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    FirstX = X ' set FirstX as X (mouse coordinate X)
    FirstY = Y ' set FirstY as Y (mouse coordinate Y)
End If

End Sub

Private Sub LblTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Me.Left = Me.Left + (X - FirstX) ' when lbltop is moved move form
    Me.Top = Me.Top + (Y - FirstY) ' when lbltop is moved move form
End If

End Sub
