VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmExplore 
   Caption         =   "Project Explorer"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   14085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   12720
      TabIndex        =   3
      ToolTipText     =   "Close screen"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   7455
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   8295
   End
   Begin ComctlLib.ProgressBar progBar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.TreeView treeCode 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   13150
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":0E42
            Key             =   "event"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":1194
            Key             =   "project"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":14E6
            Key             =   "class"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":1838
            Key             =   "method"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":1B8A
            Key             =   "variable"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":1EDC
            Key             =   "wizard"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Public Sub CleanScreen()
    On Error Resume Next
    txtCode.Text = ""
    treeCode.Nodes.Clear
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    CleanScreen
    Err.Clear
End Sub
