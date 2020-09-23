VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View.exe - "
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtView 
      Height          =   2775
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4895
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmImage.frx":0E42
   End
   Begin VB.Image imgImage 
      Height          =   2250
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

