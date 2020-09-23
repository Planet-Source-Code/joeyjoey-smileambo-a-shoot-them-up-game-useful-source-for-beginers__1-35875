VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About..."
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   2940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmroJump 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   120
   End
   Begin VB.Image ImgCon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      Picture         =   "FrmAbout.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image ImgNirvanaSmiley 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1080
      Left            =   720
      Picture         =   "FrmAbout.frx":0414
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label LblNose 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lbldirection 
      Caption         =   "-50"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Lblo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Writen By Joseph B  rg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   320
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Smileambo v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   405
      Width           =   1455
   End
   Begin VB.Image ImgSmile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      Picture         =   "FrmAbout.frx":12CE
      Top             =   360
      Width           =   315
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Ok"
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
      Left            =   960
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
      TabIndex        =   2
      Top             =   1920
      Width           =   2940
   End
   Begin VB.Label LblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About..."
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
      TabIndex        =   1
      Top             =   0
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
Attribute VB_Name = "Frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstX As Integer
Dim FirstY As Integer

Private Sub ImgSmile_Click()
ImgSmile = ImgCon ' change image to ImgCon

End Sub


Private Sub Lbl1_Click()
If Lbl1.Caption = "Zm1le4mb0 vi.o" Then Lbl1.Caption = "Smileambo v1.0" Else Lbl1.Caption = "Zm1le4mb0 vi.o" 'change label

End Sub

Private Sub LblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblQuit.BackColor = &H80000005 ' change colors

End Sub

Private Sub LblNose_Click()
LblQuit.Caption = "Not &Ok" ' change caption

End Sub

Private Sub Lblo_Click()
TmroJump.Enabled = True ' start timer

End Sub

Private Sub LblQuit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then ' if right click then ...
    LblQuit.Caption = "ERROR : " & Round(Rnd * 90) ' change caption
    LblTop.Caption = "Windows..." ' change caption
End If

If Button = 1 Then ' if left click then ...
    FrmMain.Show ' show main screen
    Unload Me ' unload about box
End If

End Sub

Private Sub LblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblQuit.BackColor = &HC0C0C0
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


Private Sub TmroJump_Timer()
If Lblo.Top <= Lbl1.Top + Lbl1.Height Then lbldirection.Caption = Round(200 * Rnd)
If Lblo.Top >= LblBack.Top - Lblo.Height Then lbldirection.Caption = "-50"
Lblo.Top = Lblo.Top + lbldirection
' make the "o" bounce

End Sub
