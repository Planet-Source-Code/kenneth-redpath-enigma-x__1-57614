VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMemory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmMemory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imgMem 
      Left            =   120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemory.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemory.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemory.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemory.frx":1F10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin EnigmaX.isButton cmdClose 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3240
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
   Begin VB.Frame fraMemory 
      Caption         =   "Memory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin EnigmaX.XP_ProgressBar pgsMem 
         Height          =   2295
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   4048
         Color           =   11034163
      End
      Begin MSComctlLib.ListView lstMem 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4683
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "imgMem"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3528
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   2117
            ImageIndex      =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   6174
            ImageIndex      =   3
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Memory Usage"
            Object.Width           =   3175
            ImageIndex      =   4
         EndProperty
      End
      Begin VB.Label lblFree 
         BackStyle       =   0  'Transparent
         Caption         =   "Used"
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
         Left            =   5160
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblMem 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
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
         Left            =   5160
         TabIndex        =   3
         Top             =   2640
         Width           =   495
      End
   End
   Begin EnigmaX.isButton cmdDelete 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Delete"
      IconAlign       =   1
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
   Begin EnigmaX.isButton cmdLoad 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Style           =   9
      Caption         =   "Load"
      IconAlign       =   1
      Enabled         =   0   'False
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
   End
End
Attribute VB_Name = "frmMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    If cmdLoad.Enabled = False Then
    MsgBox "TODO: Add the Code to Load the Different Software."
    End If
End Sub

Private Sub Form_Load()
    'Dim li
    'Loads the Software -- Might have to change depending on load game, buy software etc
    
    'Set li = lstMem.ListItems.Add(, , "View") 'The View.exe program
    'li.ListSubItems.Add , , "v1.0" 'Version
    'li.ListSubItems.Add , , "View certain types of files. (*.txt, *.jpg etc.)" 'Description
    'li.ListSubItems.Add , , "45,056 bytes" 'Size

    'Dim dsd
    'dsd = (5691776 + 932768 + 225312 + 45056)
    'pgsMem.Value = GetPercentage(dsd, MemorySize)
    
    'Dim MemPerUsed
    'MemPerUsed = GetPercentage(dsd, MemorySize)
    
    'lblMem.Caption = Format(MemPerUsed, "###0") & " %"
    
End Sub

Public Function GetPercentage(ByVal Value, ByVal Total) As Single
    Value = Value * 100 'Multiply by 100
    Total = Value / Total 'Then devide by the total to Get your percentage
    GetPercentage = Total 'Return the percentage
End Function
