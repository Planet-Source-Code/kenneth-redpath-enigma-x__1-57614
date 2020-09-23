VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6375
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tbOptions 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Sounds"
      TabPicture(1)   =   "frmOptions.frx":0E5E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDialingSound"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame fraDialingSound 
         Caption         =   "Play the Dialling Sound"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5415
         Begin VB.CheckBox chkPlayDial 
            Caption         =   "Disable playing the Dialling Sound"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   5175
         End
      End
   End
   Begin EnigmaX.isButton cmdOk 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Style           =   9
      Caption         =   "OK"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin EnigmaX.isButton cmdCancel 
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Cancel"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   6360
      Y1              =   5610
      Y2              =   5610
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   6360
      Y1              =   5625
      Y2              =   5625
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me 'Unloads me
End Sub

Private Sub cmdOk_Click()
    
    If chkPlayDial.Value = 1 Then 'Dont Play the Sound
        CreateKey "HKCU\Software\Kenneth Redpath\Enigma X\Sounds\dial.mp3", "False"
    Else 'Play it
        CreateKey "HKCU\Software\Kenneth Redpath\Enigma X\Sounds\dial.mp3", "True"
    End If
    
    
    
    
    
    
    
    
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    'Check if they want to play the dial.mp3 file
    Result = ReadKey("HKCU\Software\Kenneth Redpath\Enigma X\Sounds\dial.mp3")
    If Result = "False" Then 'Check the check box (They dont want to play it)
        chkPlayDial.Value = 1
    Else
        chkPlayDial.Value = 0
    End If
        
End Sub
