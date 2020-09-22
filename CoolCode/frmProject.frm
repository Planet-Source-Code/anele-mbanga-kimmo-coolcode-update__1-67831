VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProject 
   Caption         =   "Project Details"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   12360
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCode 
      Height          =   7935
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmProject.frx":0E42
      Top             =   120
      Width           =   6255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProject.frx":0E48
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProject.frx":129A
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treeProject 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   13996
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
    On Error Resume Next
    treeProject.Height = Me.ScaleHeight - 240
    txtCode.Height = Me.ScaleHeight - 240
    txtCode.Width = Me.ScaleWidth - treeProject.Width - 360
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdGoto_Click()
    On Error Resume Next
    'Dim lstSel As Node
    'Dim xLine() As String
    'Dim varLoc As Long
    'Set lstSel = lstProject.SelectedItem
    'If TypeName(lstSel) = "Nothing" Then Exit Sub
    'xLine = LstViewGetRow(lstProject, lstSel.Index)
    'varLoc = Val(xLine(6))
    'Set VbCp = VBInst.ActiveVBProject.VBComponents(xLine(2))
    'If TypeName(VbCp) = "Nothing" Then Exit Sub
    'VbCp.Activate
    'VbCp.CodeModule.CodePane.Show
    'VbCp.CodeModule.CodePane.TopLine = varLoc
    'VbCp.CodeModule.CodePane.SetSelection varLoc, 1, varLoc, -1
    Err.Clear
End Sub
'Private Sub lstProject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    On Error Resume Next
'    LstViewSwapSort lstProject, ColumnHeader
'    Err.Clear
'End Sub
Private Sub treeProject_BeforeLabelEdit(Cancel As Integer)
    On Error Resume Next
    Err.Clear
End Sub
Private Sub treeProject_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    txtCode.Text = Node.Tag
    Err.Clear
End Sub
