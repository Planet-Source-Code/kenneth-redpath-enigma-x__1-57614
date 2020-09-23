VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmGame 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enigma X"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12495
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmGame.frx":0E42
   ScaleHeight     =   7455
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dialog 
      Left            =   1080
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMissionValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6720
      Width           =   3615
   End
   Begin RichTextLib.RichTextBox txtMissionInfo 
      Height          =   2295
      Left            =   7440
      TabIndex        =   22
      Top             =   4320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmGame.frx":3FD4
   End
   Begin VB.TextBox txtMissionName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   21
      Top             =   3960
      Width           =   4935
   End
   Begin EnigmaX.isButton cmdSoftware 
      Height          =   495
      Left            =   7440
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":4082
      Style           =   9
      Caption         =   "Software"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdViewHomeComp 
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":409E
      Style           =   9
      Caption         =   "Home Computer"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.PictureBox PicYou 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7560
      Picture         =   "frmGame.frx":40BA
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame fraUserInformation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Character Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.Label lblTotalExp 
         BackStyle       =   0  'Transparent
         Caption         =   "/ 1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label lblUsersExp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblEmail1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblExp 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
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
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$2000"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   7095
   End
   Begin RichTextLib.RichTextBox txtConsole 
      Height          =   4290
      Left            =   795
      TabIndex        =   1
      Top             =   930
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7567
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmGame.frx":544B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EnigmaX.isButton cmdHardware 
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":54CE
      Style           =   9
      Caption         =   "Hardware"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdMissionsObj 
      Height          =   495
      Left            =   9120
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":54EA
      Style           =   9
      Caption         =   "Missions"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdHelp 
      Height          =   495
      Left            =   10800
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":5506
      Style           =   9
      Caption         =   "Help"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdExit 
      Height          =   495
      Left            =   10800
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":5522
      Style           =   9
      Caption         =   "Exit"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdSaveGame 
      Height          =   495
      Left            =   9120
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":553E
      Style           =   9
      Caption         =   "Save Game"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdLoadGame 
      Height          =   495
      Left            =   9120
      TabIndex        =   20
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":555A
      Style           =   9
      Caption         =   "Load Game"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdCancelMission 
      Height          =   300
      Left            =   10800
      TabIndex        =   24
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      Icon            =   "frmGame.frx":5576
      Style           =   9
      Caption         =   "Cancel Mission"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdOptions 
      Height          =   495
      Left            =   7440
      TabIndex        =   27
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":5592
      Style           =   9
      Caption         =   "Options"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdTaskMonitor 
      Height          =   495
      Left            =   10800
      TabIndex        =   28
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":55AE
      Style           =   9
      Caption         =   "Task Monitor"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdEmail 
      Height          =   495
      Left            =   9120
      TabIndex        =   29
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":55CA
      Style           =   9
      Caption         =   "Email"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdMemory 
      Height          =   495
      Left            =   10800
      TabIndex        =   30
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmGame.frx":55E6
      Style           =   9
      Caption         =   "Memory"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin MediaPlayerCtl.MediaPlayer PlayMp3 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   0   'False
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblMissionValue 
      BackStyle       =   0  'Transparent
      Caption         =   "Mission Value:"
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
      Left            =   7440
      TabIndex        =   25
      Top             =   6765
      Width           =   1335
   End
   Begin VB.Menu mnuGameFile 
      Caption         =   "File"
      Begin VB.Menu mnuGameExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuGameTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuGameOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Makes sure all the variables are declared

Private Sub cmdEmail_Click() 'They Clicked the (Email) button
    frmEmail.Show vbModal, Me 'Show the Email form
End Sub

Private Sub cmdExit_Click() 'They Clicked the (Exit) button
    'Change this so if accidently pressed it comes up are you sure you want to exit.??
    End 'Ends the Program
End Sub

Private Sub cmdHardware_Click() 'They Clicked the (Hardware) button
    frmHardware.Show vbModal, Me 'Shows the Hardware Form where they can buy/sell hardware
End Sub

Private Sub cmdHelp_Click() 'They Clicked the (Help) button
    'This is where i will think to the help
    MsgBox "TODO: ADD LINK TO HELP"
End Sub

Private Sub cmdLoadGame_Click() 'They Clicked the (Load Game) button
    MsgBox "TODO: Insert the Code for LOAD GAME. :-("
End Sub

Private Sub cmdMemory_Click() 'They Clicked the (Memory) button
    frmMemory.Show vbModal, Me 'Shows the Memory Form
End Sub

Private Sub cmdMissionsObj_Click() 'They Clicked the (Mission) button
    frmMissions.Show vbModal, Me 'Shows the Missions Form
End Sub

Private Sub cmdOptions_Click() 'They Clicked the (Options) button
    frmOptions.Show vbModal, Me 'Shows the Options form where they can change settings
End Sub

Private Sub cmdSaveGame_Click() 'They Clicked the (Save Game) button
    MsgBox "TODO: Create code for Save Game. :-O" & vbCrLf & _
            "This should be disabled until setup of server " & _
            "is Completed"
            
    dialog.ShowOpen
    Open dialog.FileName For Output As #1
    Print #1, Encrypt(s_UserName & vbCrLf & s_Email & vbCrLf & s_Password)
    Close #1
    
    
    Dim TextLine As String
    Open dialog.FileName For Input As #1
            Input #1, TextLine
            MsgBox Decrypt("2753106213210310731062643107310624629931113109213210254253252254253")
    Close #1
End Sub

Private Sub cmdSoftware_Click() 'They Clicked the (Software) button
    frmSoftware.Show vbModal, Me 'Shows the Software Form where they can buy/sell software
End Sub

Private Sub cmdTaskMonitor_Click() 'They Clicked the (Task Monitor) button
    frmTaskMonitor.Show vbModal, Me 'Shows the Task Monitor form
End Sub

Private Sub cmdViewHomeComp_Click() 'They Clicked the (Home Computer) button
    frmComputer.Show vbModal, Me 'Shows the computer form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End 'Ends the Program
End Sub

Private Sub mnuGameOptions_Click() 'They Clicked the (Options) button
    frmOptions.Show vbModal, Me 'Show the Options form
End Sub

Private Sub txtConsole_Change()
    txtConsole.SelStart = Len(txtConsole.Text)
End Sub

Private Sub txtConsole_GotFocus()
    'Basically this means, if the user clicks on the console text box _
     it will automatically go to the User's text box. Ready for the commands _
     to be entered.
    txtUser.SetFocus 'Sets the focus to the User Input Text Box
End Sub

Private Sub Form_Load()
    Unload frmMain 'Unloads the form that we came from -- This is just in case that when this is _
                    connecting it might show the initial form. So this is just in case.
                    
    Me.Show 'Shows this form

'Character Information
    lblUserName.Caption = s_UserName 'This shows the user's name in the Name label
    lblEmail.Caption = s_Email 'This shows the user's email in the Email label

    txtUser.SelStart = 0 'This is to make sure that it is at the start of the User Input text box.
    txtConsole.SelStart = 0 'This is to make sure that it is at the start of the Console text box.
    txtConsole.SelColor = &H808080 'Sets the Colour of the text in the console _
                                    to Grey (sort of like the old DOS Colour)

If b_GameIsLoaded = False Then 'Has a game been loaded or is it a new game?
    Connecting = True 'This tells the program that it is connecting. So _
                       if the user enters code it does nothing
    Disconnected = False 'This tells the program that it is not disconnected duh! :-)
    ConHomeComputer 'Connects to Home Computer
    
Else 'A Game has been loaded
    
    'Need to figure a way to decrypt a file and load user name pass etc.
    
    'This will change depending on the mission and the settings
    ConHomeComputer 'Connects to Home Computer
    
End If 'The IF is if a game has been loaded

End Sub 'Ends the Sub :(

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer) 'When user Presses a key
If txtUser.Text = "" Then 'Is there anything in the User's Text Box? :-O
Else 'There was something in there
    
    If Connecting Or Disconnecting Or DoingSomething = True Then 'It is Connecting, Disconnecting or Doing something
    
    Else 'It is not Connecting, Disconnecting or Doing something
    
        Text = txtUser.Text 'Save what the user has typed to a string

        If KeyCode <> 13 Then Exit Sub 'If the user didn't press Enter dont continue
            txtUser.Text = ""
            txtUser.SelStart = 0
    
            Select Case Server
                Case "Server X10 - Home Computer"
                    KeysHome 'Call KeysHome from (KeyPressHome.bas)
                    Exit Sub 'Three Guesses :-)
                    
                    
            End Select 'Server Select
    
        If Disconnected = False Then 'Have we been disconnected
            'These are shown when their was an unrecognized command typed in
            txtConsole.Text = txtConsole.Text & vbCrLf & "'" & Text & "' is not recognized as an internal or " & _
                                                                 "external command, operable program or batch file." & vbCrLf
            txtConsole.Text = txtConsole.Text & vbCrLf & "If you Need help on Commands type command or Press F1" & vbCrLf
            txtConsole.Text = txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Else 'The Computer is disconnected
            txtConsole.Text = txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
        End If
        Exit Sub
    End If 'Ends IF the program is Connecting, Disconnecting or Doing Something
    Exit Sub

End If 'Ends IF there is no text entered into the User text box

End Sub 'Ends the txtUser_KeyDown Sub :-(
