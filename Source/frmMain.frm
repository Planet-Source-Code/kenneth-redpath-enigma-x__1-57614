VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Enigma X"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   4500
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog LoadGame 
      Left            =   1320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin EnigmaX.isButton cmdNewGame 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmMain.frx":20B3
      Style           =   9
      Caption         =   "New Game"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   4210752
      HighlightColor  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdLoadGame 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmMain.frx":20CF
      Style           =   9
      Caption         =   "Load Game"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   4210752
      HighlightColor  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdCredits 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmMain.frx":20EB
      Style           =   9
      Caption         =   "Credits"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   4210752
      HighlightColor  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdHelp 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmMain.frx":2107
      Style           =   9
      Caption         =   "Help"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   4210752
      HighlightColor  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdExit 
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Icon            =   "frmMain.frx":2123
      Style           =   9
      Caption         =   "Exit"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   4210752
      HighlightColor  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.Image imgMain 
      Height          =   3855
      Left            =   120
      Picture         =   "frmMain.frx":213F
      Top             =   540
      Width           =   3060
   End
   Begin VB.Image imgStopDrag1 
      Height          =   465
      Left            =   2985
      Top             =   60
      Width           =   1575
   End
   Begin VB.Image imgStopDrag2 
      Height          =   3945
      Left            =   0
      Top             =   480
      Width           =   4875
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   4680
      Top             =   105
      Width           =   210
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Makes sure all the variables are declared

'*******Handles the Drag of the Form when Mouse is Down*******
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
           (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
           lParam As Long) As Long
'*******************End of Drag of Form Code*******************

Private Sub cmdCredits_Click() 'They Clicked the (Credits) button
    frmCredits.Show vbModal, Me 'Show the Credits Form
End Sub

Private Sub cmdExit_Click() 'They Clicked the (Exit) button
    End 'Exits the Game
End Sub

Private Sub cmdHelp_Click() 'They Clicked the (Help) button
    frmHelp.Show vbModal, Me
End Sub

Private Sub cmdLoadGame_Click() 'They Clicked the (Load Game) button
    On Error GoTo Cancel 'If there is an error goto Cancel:
    
    With LoadGame 'LoadGame is the Common Dialog Control's Name
        .CancelError = True 'If cancel is pressed flag an error
        .Flags = cdlOFNHideReadOnly 'Hide the Read Only files
        .InitDir = App.Path & "\Saved Games" 'Start of in the Saved Games folder
        .DialogTitle = "Load a Enigma X Saved Game" 'Sets the Title
        .Filter = "Enigma X Saved Games (*.exg)|*.exg" 'Sets what sort of file's it's looking for
        .ShowOpen 'Shows the Dialog box
        s_LoadGameFileName = .FileName 'Assigns the File name of the Loaded file to the Variable s_LoadGameFileName
        
        'Enable when doing the load game file crap
        b_GameIsLoaded = True 'This tells us that there has been a game loaded
        NewServer = False 'This tells us that they have already set-up the Server
        
        'This is just temporary until i figure a way to store values in files.
        'These values will be RIPPED(lol) from the Saved Game file.
        s_UserName = "Kj" 'The User's Name
        s_Email = "kj@kj.com" 'The User's Email
        s_Password = "65465" 'The User's Password, Not currently useful but you never know.
        
        'Do they want to play the dial.mp3 file?
        PlayFile = False
        
        'these will also be taken from save game
        MotherBoard = "ECS P6FX1-A Pentium Pro ATX Motherboard" 'An Updated version of my motherboard
        CPU = "Intel Pentium 200Mhz" 'I wish i had this (yeah :-) NOT )
        Memory = "4MB" 'I've got 768MB -- Gotta start you with somethin good :-)
        MemorySize = "4000000" 'in bytes
        Modem = "Standard 300bps Modem" 'I'm only on dial-up :-( (as yet this does nothing, i am unsure of what i would do with it (Maybe slow downloads down)
        HDDSize = "200000000"
        HDDName = "Western Digital Caviar (200Mb)"
        HDDSerialNo = "AC200-32LA"
        
    End With
    
    frmGame.Show 'Shows the Game form
    Unload Me 'Hide this form
    Exit Sub 'Exits the sub
    
Cancel: 'If the User pressed cancel or there was an error loading the file do this
    Me.Show 'Shows this form
End Sub

Private Sub cmdNewGame_Click() 'They Clicked the (New Game) button
    
    'It is a new Server so set-up everything
    NewServer = True
    PlayFile = True
    
    frmNewGame.Show vbModal, Me

    b_GameIsLoaded = False 'A game hasn't been loaded so tell the program
End Sub

'This is when the user clicks on the title so it will move the form
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then 'If they used the Left Mouse Button
        Dim ReturnVal As Long
        X = ReleaseCapture() 'Where did they release the button

        'Move the window to that point and keep it there
        ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub imgClose_Click() 'They Clicked on the (X) Close image
    End 'End the Game
End Sub
