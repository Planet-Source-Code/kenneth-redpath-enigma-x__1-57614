VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credits"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin EnigmaX.isButton cmdClose 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Close"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin VB.Frame fraCredits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Credits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtCredits 
         BackColor       =   &H00E0E0E0&
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmCredits.frx":0E42
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label lblCredits 
         BackStyle       =   0  'Transparent
         Caption         =   "These are the people who, without them this would have never have gotten this far."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click() 'They Clicked on the Close button
    Unload Me 'Unloads Me
End Sub
