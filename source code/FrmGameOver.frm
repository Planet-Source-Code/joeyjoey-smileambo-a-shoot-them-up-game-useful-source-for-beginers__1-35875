VERSION 5.00
Begin VB.Form FrmGameOver 
   BorderStyle     =   0  'None
   Caption         =   "Game Over"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   2940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Lblnext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Next level"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label LblPlayAgain 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Play Again"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Your dead ... mikkinu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image ImgNirvanaSmiley 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1080
      Left            =   720
      Picture         =   "FrmGameOver.frx":0000
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label LblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Game Over"
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
      TabIndex        =   2
      Top             =   0
      Width           =   2940
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   2940
   End
   Begin VB.Shape ShpBack 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   0
      Top             =   240
      Width           =   2940
   End
End
Attribute VB_Name = "FrmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstX As Integer
Dim FirstY As Integer

Private Sub Form_Load()
Unload FrmGame ' unload level1
Unload FrmGame2 ' unload level2

End Sub

Private Sub LblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblPlayAgain.BackColor = &H80000005 ' change color back to normal
LblQuit.BackColor = &H80000005 ' change color back to normal

End Sub

Private Sub Lblnext_Click()
Unload Me ' unload game over
FrmGame2.Show ' show level 2

End Sub

Private Sub Lblnext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lblnext.BackColor = &HC0C0C0 ' change color

End Sub

Private Sub LblPlayAgain_Click()
Unload Me ' unload game over
FrmGame.Show ' load level1

End Sub

Private Sub LblPlayAgain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblPlayAgain.BackColor = &HC0C0C0 ' change color

End Sub

Private Sub LblQuit_Click()
End ' quit game

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
