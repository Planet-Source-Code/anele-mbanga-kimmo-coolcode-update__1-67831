VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLinesComments 
   Caption         =   "Lines Starting With Comments"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLinesComments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ToExcel 
      Caption         =   "To Excel"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      ToolTipText     =   "Close screen"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Goto"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      ToolTipText     =   "Goto line selected on variable list"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      ToolTipText     =   "Remove all lines starting with comments"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      ToolTipText     =   "Close screen"
      Top             =   6240
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstVariables 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code In Line"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuToExcel 
         Caption         =   "To Excel"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "Go To"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmLinesComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    On Error Resume Next
    lstVariables.ListItems.Clear
    Unload Me
    Err.Clear
End Sub
Private Sub cmdGoto_Click()
    On Error Resume Next
    Dim sNode As Object
    Dim spLine() As String
    Dim varLoc As Long
    'Dim varName As String
    Set sNode = lstVariables.SelectedItem
    If TypeName(sNode) = "Nothing" Then Exit Sub
    spLine = LstViewGetRow(lstVariables, sNode.Index)
    varLoc = Val(spLine(3))
    Set VbCp = VBInst.ActiveVBProject.VBComponents(spLine(1))
    If TypeName(VbCp) = "Nothing" Then Exit Sub
    VbCp.CodeModule.CodePane.Show
    VbCp.CodeModule.CodePane.TopLine = varLoc
    VbCp.CodeModule.CodePane.SetSelection varLoc, 0, varLoc, -1
    Err.Clear
End Sub
Private Sub lstVariables_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    LstViewSwapSort lstVariables, ColumnHeader
    Err.Clear
End Sub
Private Sub mnuClose_Click()
    On Error Resume Next
    cmdClose_Click
    Err.Clear
End Sub
Private Sub ToExcel_Click()
    On Error Resume Next
    Dim xFile As String
    xFile = StringGetFileToken(VBInst.ActiveVBProject.Filename, "p") & "\Lines Starting With Comments.xls"
    LstViewToWorkSheetAsIs lstVariables, xFile, "Lines Starting With Comments", "Kimmo - CoolCode"
    Err.Clear
End Sub
