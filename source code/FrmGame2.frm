VERSION 5.00
Begin VB.Form FrmGame2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrEnemyMove 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4440
      Top             =   3600
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   0
      Width           =   315
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
      TabIndex        =   9
      Top             =   0
      Width           =   6780
   End
   Begin VB.Image ImgChracter 
      Height          =   285
      Left            =   360
      Picture         =   "FrmGame2.frx":0000
      Top             =   1320
      Width           =   750
   End
   Begin VB.Image ImgBoxes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      Top             =   1440
      Width           =   285
   End
   Begin VB.Image ImgEnemy2 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   5880
      Top             =   1680
      Width           =   405
   End
   Begin VB.Image ImgEnemy 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   6120
      Top             =   1080
      Width           =   405
   End
   Begin VB.Image imgBullet2 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   6000
      Picture         =   "FrmGame2.frx":03DB
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2520
      Left            =   0
      Picture         =   "FrmGame2.frx":0505
      Top             =   285
      Width           =   6810
   End
   Begin VB.Image ImgBlood 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   6360
      Picture         =   "FrmGame2.frx":37491
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label Lblimg2top 
      Caption         =   "-50"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image ImgAngrieGatling 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   5040
      Picture         =   "FrmGame2.frx":37AE5
      Top             =   3600
      Width           =   795
   End
   Begin VB.Image ImgCenter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   3220
      Picture         =   "FrmGame2.frx":38111
      Stretch         =   -1  'True
      Top             =   2850
      Width           =   555
   End
   Begin VB.Image imgBullet 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   360
      Picture         =   "FrmGame2.frx":38263
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgCharnormal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Picture         =   "FrmGame2.frx":38425
      Top             =   3600
      Width           =   750
   End
   Begin VB.Image imgCharshooting 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      Picture         =   "FrmGame2.frx":38800
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label lblnumbullets 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "200"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Lblammo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ammo :"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblhealth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Health :"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblhitpoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label LblEnemy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5760
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image ImgBat 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2520
      Picture         =   "FrmGame2.frx":38E17
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   405
   End
   Begin VB.Image ImgMinibul 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   150
      Left            =   2160
      Picture         =   "FrmGame2.frx":392B5
      Top             =   3720
      Width           =   285
   End
   Begin VB.Image ImgAngrie 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3000
      Picture         =   "FrmGame2.frx":393DF
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image ImgConf 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      Picture         =   "FrmGame2.frx":395B0
      Top             =   3600
      Width           =   315
   End
   Begin VB.Label LblEnemyHel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label LblnumKills 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label LblKills 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kills :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image ImgPresent 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4080
      Picture         =   "FrmGame2.frx":396FE
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image ImgWink 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      Picture         =   "FrmGame2.frx":398CE
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   315
   End
   Begin VB.Shape shpScreen 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   0
      Top             =   240
      Width           =   6780
   End
   Begin VB.Label LblBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   6780
   End
End
Attribute VB_Name = "FrmGame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FirstX As Integer
Dim FirstY As Integer
Dim vleft As Boolean
Dim vright As Boolean
Dim vup As Boolean
Dim vdown As Boolean
Dim vcharMove As Boolean
Dim vmovespeed As Integer
Dim vRndmon As Integer
Dim vRndBul As Integer
Dim vrepulsion As Integer
Dim vshooting As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeySpace ' space bar
        vshooting = True
    Case vbKeyLeft ' left arrow key
        vleft = True
    Case vbKeyRight ' right arrow key
        vright = True
    Case vbKeyUp ' up arrow key
        vup = True
    Case vbKeyDown ' down arrow key
        vdown = True
End Select
         
If vleft = True And ImgChracter.Left > shpScreen.Left Then ' if left button is pressed and user is within screen then ...
    ImgChracter.Left = ImgChracter.Left - vmovespeed ' move the character left by vmovespeed twips
    onbox ' check if character is on a box
End If

If vright = True And ImgChracter.Left < shpScreen.Left - ImgChracter.Width + shpScreen.Width Then ' if right button is pressed and user is within screen then ...
    ImgChracter.Left = ImgChracter.Left + vmovespeed ' move the character right by vmovespeed twips
    onbox ' check if character is on a box
End If

If vup = True And ImgChracter.Top > shpScreen.Top Then ' if up button is pressed and user is within screen then ...
    ImgChracter.Top = ImgChracter.Top - vmovespeed ' move the character up by vmovespeed twips
    onbox ' check if character is on a box
End If

If vdown = True And ImgChracter.Top < shpScreen.Top - ImgChracter.Height + shpScreen.Height Then ' if down button is pressed and user is within screen then ...
    ImgChracter.Top = ImgChracter.Top + vmovespeed ' move the character down by vmovespeed twips
    onbox ' check if character is on a box
End If

If vshooting = True Then ' if space bar is pressed then ...
    shoot ' check bullet movement
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeySpace ' space bar
        vshooting = False ' reset shooting
    Case vbKeyLeft ' left arrow key
        vleft = False ' reset left movement
    Case vbKeyRight ' right arrow key
        vright = False ' reset right movement
    Case vbKeyUp ' up arrow key
        vup = False ' reset up movement
    Case vbKeyDown ' down arrow key
        vdown = False ' reset down movement
End Select

If vshooting = False Then ImgChracter = imgCharnormal ' change image to normal
          
End Sub

Private Sub Form_Load()
ImgEnemy2 = ImgAngrieGatling ' set picture as ImgAngrieGatling
bat ' start with bat as opponent
vleft = False ' not moving left
vright = False ' not moving right
vup = False ' not moving up
vdown = False ' not moving down
vcharMove = True ' character can move
vmovespeed = 100 ' character movement speed
vrepulsion = 100 ' who much will the character be pussed away from walls

End Sub

Private Sub LblQuit_Click()
FrmMain.Show ' show main screen
Unload Me ' unload level 2

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
LblQuit.BackColor = &H80000005 ' reset color

If Button = vbLeftButton Then ' if left mouse button click then ...
    Me.Left = Me.Left + (X - FirstX) ' when lbltop is moved move form
    Me.Top = Me.Top + (Y - FirstY) ' when lbltop is moved move form
End If

End Sub
Private Sub shoot()
If CInt(lblnumbullets.Caption) <= 0 Then GoTo noammo ' if ammo is 0 or less then goto the end of shoot
DoEvents
imgBullet.Top = ImgChracter.Top + 70 ' set bullets top to that of the character
imgBullet.Left = ImgChracter.Left ' set bullets left to that of the character
imgBullet.Visible = True ' show bullet
ImgChracter = imgCharshooting.Picture ' change character's image to shooting

Do
 imgBullet.Left = imgBullet.Left + 50 ' move bullet
 If imgBullet.Left <= ImgEnemy.Left + ImgEnemy.Width And imgBullet.Left >= ImgEnemy.Left - imgBullet.Width And imgBullet.Top >= ImgEnemy.Top And imgBullet.Top <= ImgEnemy.Top + ImgEnemy.Height Then enemyhlt ' if enemy was hit go to enemyhlt

Loop Until imgBullet.Left >= 7080 ' do this until bullet reaches the screen's end

imgBullet.Top = ImgChracter.Top + 70 'set bullets top to that of the character
imgBullet.Left = shpScreen.Left + shpScreen.Width ' set bullets to end of screen
lblnumbullets.Caption = CStr(CInt(lblnumbullets.Caption) - 1) ' decrese ammo
imgBullet.Visible = False ' hide bullet
noammo:

End Sub
Private Sub enemyhlt()
If CInt(LblEnemyHel.Caption) <= 0 Then ' if enemy's health is 0 or less then ...
    ImgBlood.Top = ImgEnemy.Top ' blood is shown on the same top of ImgEnemy
    ImgBlood.Left = ImgEnemy.Left  ' blood is shown on the same left of ImgEnemy
    ImgBlood.Visible = True ' show the blood
    ImgEnemy.Left = 6240 ' reset ImgEnemy's position
    ImgEnemy.Top = Round((2760 - 400) * Rnd) ' set a random top
    If ImgEnemy.Top <= shpScreen.Top Then ImgEnemy.Top = ImgEnemy.Top + 400 ' if ImgEnemy is out of the screen than add 400 to its posiotion
    dead ' go to dead
    GoTo enedead ' skip reduction of enemy's health
End If

LblEnemyHel.Caption = LblEnemyHel.Caption - 1 ' decrese enemy's health
enedead:

End Sub

Private Sub bat()
ImgEnemy = ImgBat ' change image to bat
LblEnemy = "Bat" ' change label to bat
LblEnemyHel = "26" ' change enemy's health to 26
enemymove ' move enemy

End Sub
Private Sub angrie()
ImgEnemy = ImgAngrie ' change image to angrie
LblEnemy = "Angrie" ' change label to angrie
LblEnemyHel = "65" ' change enemy's health to 65
enemymove ' move enemy

End Sub
Private Sub confu()
ImgEnemy = ImgConf ' change image to confu
LblEnemy = "Confu" ' change label to confu
LblEnemyHel = "52" ' change enemy's health to 52
enemymove ' move enemy

End Sub
Private Sub dead()
LblnumKills = LblnumKills + 1 ' increase number of kills
vRndmon = Round(Rnd * 3) ' randomize a number from 0 to 2
If vRndmon = 0 Then bat ' if 0 then go to bat
If vRndmon = 1 Then angrie ' if 1 then go to angrie
If vRndmon = 2 Then confu ' if 2 then go to confu
box ' drop box onto the screen
ImgEnemy.Visible = True ' make enemy visible
If LblnumKills = 30 Then levelwon 'kills needed to go to level 2

End Sub
Private Sub enemymove()
TmrEnemyMove.Enabled = True ' eneble enemy movement timer

End Sub

Private Sub TmrEnemyMove_Timer()
If ImgEnemy.Left <= shpScreen.Left - ImgEnemy.Width Then ImgEnemy.Left = 6120 ' if the enemy reaches the screen's left then reset position
If ImgEnemy.Left <= ImgChracter.Left + ImgChracter.Width And ImgEnemy.Left >= ImgChracter.Left - ImgEnemy.Width And ImgEnemy.Top >= ImgChracter.Top And ImgEnemy.Top <= ImgChracter.Top + ImgChracter.Height Then ' if enemy hits character then ...
    lblhitpoints = lblhitpoints - 1 ' derease character's health
    If lblhitpoints <= 0 Then 'if users health is 0 or less then ...
        Unload Me 'unload level 1
        FrmGameOver.Show ' show game over
    End If
    
    If lblhitpoints < 50 Then 'if users health is less than 50 then ...
        ImgCenter = ImgConf ' change image to confused
    End If
  
    If lblhitpoints < 20 Then 'if users health is less than 20 then ...
        ImgCenter = ImgAngrie ' change image to angry
    End If
  
    If lblhitpoints > 50 Then 'if users health is more than 50 then ...
        ImgCenter = ImgWink ' change image to normal
    End If
    
    LblEnemyHel = LblEnemyHel - 1 ' decrese enemy's health
    enemyhlt ' check if enemy was killed
End If
    
If LblEnemy = "Bat" Then ' if caption is bat then ...
    ImgEnemy.Left = ImgEnemy.Left - 50 ' move by 50 twips to the left
End If

If LblEnemy = "Confu" Then  ' if caption is confu then ...
    ImgEnemy.Left = ImgEnemy.Left - 50  ' move by 50 twips to the left
    ImgEnemy.Top = (Sin(ImgEnemy.Left) * 1000) + 1350 ' move in a sine wave
End If

If LblEnemy = "Angrie" Then ' if caption is angrie then ...
    ImgEnemy.Left = ImgEnemy.Left - 100 ' move by 100 twips to the left
    ImgEnemy.Top = ImgChracter.Top ' make ImgEnemy's top the same as ImgChracter.Top
End If

    
    
If ImgEnemy2.Top < 240 Then Lblimg2top = "+50" ' if image's top is the top of the screen set label to +50
If ImgEnemy2.Top > shpScreen.Height - ImgEnemy2.Height Then Lblimg2top = "-50" ' if image's top + height is the bottom of the screen set label to -50
ImgEnemy2.Top = ImgEnemy2.Top + Lblimg2top ' cahnge top by Lblimg2top
If lblhitpoints <= 70 Then GoTo noaction ' if characetr's helath is less than 70 then skip shooting
If ImgEnemy2.Top >= ImgChracter.Top And ImgEnemy2.Top <= ImgChracter.Top + ImgChracter.Height Then enemyshoot ' if enemy's top is like the caharacter's top then shoot
noaction:

ImgBlood.Visible = False ' hide the blood

End Sub

Private Sub box()
ImgBoxes.Left = (shpScreen.Width * Rnd) ' radomize box's left
ImgBoxes.Top = (shpScreen.Height * Rnd) ' radomize box's top
ImgBoxes = ImgPresent ' change image to that of a present
ImgBoxes.Visible = True 'make box visible

End Sub

Private Sub onbox()
If ImgChracter.Left <= ImgBoxes.Left + ImgBoxes.Width And ImgChracter.Left >= ImgBoxes.Left - ImgChracter.Width And ImgChracter.Top >= ImgBoxes.Top And ImgChracter.Top <= ImgBoxes.Top + ImgBoxes.Height Then ' if character is on a box then ...
    vRndBul = Round(4 * Rnd) ' set vRndBul a number 0 to 3
    If vRndBul = 0 Then
        lblnumbullets.Caption = lblnumbullets.Caption + 100 + Round(50 * Rnd) ' increse ammo by 100 + a random number from 0 to 49
    End If
    
    If vRndBul = 1 Then
        lblnumbullets.Caption = lblnumbullets.Caption + 100 - Round(50 * Rnd) ' increse ammo by 100 - a random number from 0 to 49
    End If
    
    If vRndBul = 2 Then
        lblnumbullets.Caption = lblnumbullets.Caption + 50 ' increse ammo by 50
    End If

    If vRndBul = 3 Then
       lblhitpoints.Caption = lblhitpoints.Caption + 50 ' increse health by 50
    End If

    ImgBoxes.Left = (shpScreen.Width * Rnd) ' randomize box's left
    ImgBoxes.Top = (shpScreen.Height * Rnd) ' randomize box's top
    ImgBoxes.Visible = False ' hide box
End If

End Sub
Private Sub levelwon()
Unload Me ' unload level 1
FrmGameOver.LblTop = "Level 2 Complete" ' change cation to Level 1 Complete
FrmGameOver.Lblnext.Visible = False ' hide label
FrmGameOver.Show 'show game over

End Sub

Private Sub enemyshoot()
If CInt(lblnumbullets.Caption) <= 0 Then GoTo noammo ' if ammo is 0 or less then goto the end of enemyshoot
DoEvents
imgBullet2.Top = ImgEnemy2.Top + 70 ' set bullets top to that of the enemy
imgBullet2.Left = ImgEnemy.Left ' set bullets left to that of the enemy
imgBullet2.Visible = True ' show bullet

Do
    imgBullet2.Left = imgBullet2.Left - 50  ' move bullet
    If imgBullet2.Left <= ImgChracter.Left + ImgChracter.Width And ImgChracter.Left >= ImgChracter.Left - imgBullet2.Width And imgBullet2.Top >= ImgChracter.Top And imgBullet2.Top <= ImgChracter.Top + ImgChracter.Height Then lblhitpoints = lblhitpoints - 0.05 ' if character was hit decrese his health
    If lblhitpoints <= 0 Then ' if charcter's health is 0 or less then ...
        TmrEnemyMove.Enabled = False ' disable timer
        Unload Me ' unload level 2
        FrmGameOver.Show ' show game over
    End If
    If lblhitpoints <= 50 Then ImgEnemy2.Visible = False
Loop Until imgBullet2.Left <= 0 ' move bullet untill it reaches to screen
imgBullet2.Top = ImgChracter.Top + 70 ' set bullet's top to that of the charcter
imgBullet.Left = ImgEnemy.Left ' set bullet's left to that of the enemy
imgBullet2.Visible = False ' hide bullet
noammo:
End Sub
