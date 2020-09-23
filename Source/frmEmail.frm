VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin EnigmaX.isButton cmdClose 
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Close"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin VB.TextBox txtMessage 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmEmail.frx":0E42
      Top             =   3120
      Width           =   7695
   End
   Begin MSComctlLib.ListView lstEmail 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "On Form Load do a Check maybe have progress bar etc.?? :-O"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "In the List Box have maybe if it has attachments, From, Subject, Recieved.. Maybe?? :-)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   7695
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
