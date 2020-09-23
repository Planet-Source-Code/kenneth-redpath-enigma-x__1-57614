VERSION 5.00
Begin VB.Form frmNewGame 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Game"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   Icon            =   "frmNewGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicYou 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "frmNewGame.frx":0E42
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picYou3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "frmNewGame.frx":21D3
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picYou2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "frmNewGame.frx":3583
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin EnigmaX.isButton cmdOk 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Style           =   9
      Caption         =   "OK"
      IconAlign       =   1
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin EnigmaX.isButton cmdCancel 
      Height          =   375
      Left            =   3990
      TabIndex        =   8
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Cancel"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9DAC9&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   44
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   810
      Width           =   2655
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9DAC9&
      Height          =   285
      Left            =   2160
      MaxLength       =   37
      TabIndex        =   1
      Top             =   450
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9DAC9&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   37
      TabIndex        =   0
      Top             =   90
      Width           =   2655
   End
   Begin EnigmaX.isButton cmdNext 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Style           =   9
      Caption         =   "->"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin EnigmaX.isButton cmdPrev 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Style           =   9
      Caption         =   "<-"
      IconAlign       =   1
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    If cmdNext.Enabled = True Then
        If PicYou.Visible = True Then
            PicYou.Visible = False
            picYou2.Visible = True
            picYou3.Visible = False
            cmdPrev.Enabled = True
        ElseIf picYou2.Visible = True Then
            PicYou.Visible = False
            picYou2.Visible = False
            picYou3.Visible = True
            cmdNext.Enabled = False
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    'Checks to see that the command button is not disabled. _
        This is because the button control will still do the click _
        Event even if the button is disabled.
    If cmdOk.Enabled = True Then
        'Sets the UserName, Email and Password to strings so that they _
        can be used in other parts of the game
        s_UserName = txtName.Text 'The Users Name
        s_Email = txtEmail.Text 'Their Email
        s_Password = txtPassword.Text 'Their Password
        
        
        'This is the default computer profile, that all users will start of with :-)
        MotherBoard = "ECS P6FX1-A Pentium Pro ATX Motherboard" 'An Updated version of my motherboard
        CPU = "Intel Pentium 200Mhz" 'I wish i had this (yeah :-) NOT )
        Memory = "4MB" 'I've got 768MB
        MemorySize = "4000000" 'in bytes
        Modem = "Standard 300bps Modem" 'I'm only on dial-up :-( (as yet this does nothing, i am unsure of what i would do with it (Maybe slow downloads down)
        HDDSize = "200000000"
        HDDName = "Western Digital Caviar (200Mb)"
        HDDSerialNo = "AC200-32LA"
        
        frmGame.Show 'Shows the Main Games Form
        Unload frmMain
        Unload Me 'Unloads this Form
    End If
End Sub

Private Sub cmdPrev_Click()
    If cmdPrev.Enabled = True Then
        If picYou3.Visible = True Then
            PicYou.Visible = False
            picYou2.Visible = True
            picYou3.Visible = False
            cmdNext.Enabled = True
        ElseIf picYou2.Visible = True Then
            PicYou.Visible = True
            picYou2.Visible = False
            picYou3.Visible = False
            cmdNext.Enabled = True
            cmdPrev.Enabled = False
        End If
    End If
End Sub

Private Sub txtEmail_Change()
    If txtEmail.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtName.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtPassword.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtName_Change()
    If txtName.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtEmail.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtPassword.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtPassword_Change()
    If txtPassword.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtEmail.Text = "" Then
        cmdOk.Enabled = False
    ElseIf txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
