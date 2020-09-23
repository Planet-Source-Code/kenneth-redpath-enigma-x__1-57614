VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmComputer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your Computer"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9135
   Icon            =   "frmComputer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSoftware 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Software:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2880
      TabIndex        =   15
      Top             =   1560
      Width           =   6135
      Begin EnigmaX.isButton cmdDelete 
         Height          =   300
         Left            =   5160
         TabIndex        =   17
         Top             =   2280
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Icon            =   "frmComputer.frx":0E42
         Style           =   9
         Caption         =   "Delete"
         IconAlign       =   1
         iNonThemeStyle  =   0
         Enabled         =   0   'False
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
      End
      Begin MSComctlLib.ListView lstSoftware 
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgSoftware"
         ColHdrIcons     =   "imgSoftware"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4410
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   2117
            ImageIndex      =   4
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7056
            ImageIndex      =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   2646
            ImageIndex      =   3
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgSoftware 
      Left            =   120
      Top             =   3960
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
            Picture         =   "frmComputer.frx":0E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputer.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputer.frx":1992
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputer.frx":1F2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin EnigmaX.isButton cmdClose 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmComputer.frx":24C6
      Style           =   9
      Caption         =   "Close"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.Frame fraSpecs 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hardware:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin VB.Label lblModem 
         BackStyle       =   0  'Transparent
         Caption         =   "Modem"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lblMo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modem:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblHDD 
         BackStyle       =   0  'Transparent
         Caption         =   "Hard Drive"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label lblHardDrive 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hard Drive:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblMemory 
         BackStyle       =   0  'Transparent
         Caption         =   "Memory"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblMem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblCpu 
         BackStyle       =   0  'Transparent
         Caption         =   "Cpu"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblCpuInf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CPU:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblMotherb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motherboard"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblMotherboard 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motherboard:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox imgComp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   0
      Picture         =   "frmComputer.frx":24E2
      ScaleHeight     =   3930
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2835
   End
   Begin EnigmaX.isButton cmdBuyHard 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmComputer.frx":45C1
      Style           =   9
      Caption         =   "Buy Hardware"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin EnigmaX.isButton cmdBuySoftware 
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmComputer.frx":45DD
      Style           =   9
      Caption         =   "Buy Software"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
End
Attribute VB_Name = "frmComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Makes sure all the variables are declared

Private Sub cmdBuyHard_Click() 'They Clicked the (Buy Hardware) button on the Form
    frmHardware.Show vbModal, frmGame 'Shows the Hardware window
End Sub

Private Sub cmdBuySoftware_Click() 'They Clicked the (Buy Software) button on the Form
    frmSoftware.Show vbModal, frmGame 'Shows the Software window
End Sub

Private Sub cmdClose_Click() 'They Clicked the (Close) Button
    Unload Me 'Unloads this form
End Sub

Private Sub Form_Load()

    'Loads the Current Computer Specs.
    lblMotherb.Caption = MotherBoard 'Motherboard Duh! :-)
    lblCpu.Caption = CPU 'Cpu
    lblMemory.Caption = Memory 'Memory
    lblHDD.Caption = HDDName 'Hard Drive Name
    lblModem.Caption = Modem 'Modem


    'This is for reference --- Probably
    'Dim li
    'Loads the Software -- Might have to change depending on load game, buy software etc
    
    'Set li = lstSoftware.ListItems.Add(, , "View", 1) 'The View.exe program
    'li.ListSubItems.Add , , "v1.0" 'Version
    'li.ListSubItems.Add , , "View certain types of files. (*.txt, *.jpg etc.)" 'Description
    'li.ListSubItems.Add , , "15,360 bytes" 'Size

End Sub
