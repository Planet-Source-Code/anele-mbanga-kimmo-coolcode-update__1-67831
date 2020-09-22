VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVariables 
   Caption         =   "Variable Declarations"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVariables.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9360
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ToExcel 
      Caption         =   "To Excel"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Close screen"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstVariables 
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Right click to access menu"
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parent Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member Scope"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Member Starting"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Member Ending"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Member Calls"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Goto"
      Height          =   12825
      Left            =   1.29510e5
      TabIndex        =   2
      ToolTipText     =   "Goto line selected on variable list"
      Top             =   585
      Width           =   0
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Comment"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "Comment all dead variables"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      ToolTipText     =   "Close screen"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuToExcel 
         Caption         =   "To Excel"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Comment"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmVariables"
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
    Dim sNode As Variant
    Dim spLine() As String
    Dim varLoc As Long
    Dim xMem As String
    Dim varEnd As Long
    Set sNode = lstVariables.SelectedItem
    If TypeName(sNode) = "Nothing" Then Exit Sub
    spLine = LstViewGetRow(lstVariables, sNode.Index)
    varLoc = Val(spLine(5))
    varEnd = Val(spLine(6))
    xMem = spLine(1)
    xMem = MvField(xMem, 1, ".")
    Set VbCp = VBInst.ActiveVBProject.VBComponents(xMem)
    If TypeName(VbCp) = "Nothing" Then Exit Sub
    VbCp.Activate
    VbCp.CodeModule.CodePane.Show
    VbCp.CodeModule.CodePane.TopLine = varLoc
    VbCp.CodeModule.CodePane.SetSelection varLoc, 1, varEnd, -1
    Err.Clear
End Sub
Public Sub cmdRemove_Click()
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim spLine() As String
    Dim compName As String
    Dim varName As String
    Dim varLoc As Long
    Dim varEnd As Long
    nTot = lstVariables.ListItems.Count
    For nCnt = 1 To nTot
        spLine = LstViewGetRow(lstVariables, nCnt)
        lstVariables.ListItems(nCnt).EnsureVisible
        compName = spLine(1)
        compName = MvField(compName, 1, ".")
        varName = spLine(2)
        varLoc = Val(spLine(5))
        varEnd = Val(spLine(6))
        If spLine(7) = "0" Then
            If LCase$(spLine(3)) = "variable" Or LCase$(spLine(3)) = "constant" Then
                Set VbCp = VBInst.ActiveVBProject.VBComponents(compName)
                If TypeName(VbCp) = "Nothing" Then GoTo NextVariable
                VbCp.Activate
                VbCp.CodeModule.CodePane.Show
                VbCp.CodeModule.CodePane.TopLine = varLoc
                Code_CommentCode VbCp, varLoc, varEnd, True
            End If
        End If
NextVariable:
        DoEvents
        Err.Clear
    Next
    lstVariables.ListItems.Clear
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    lstVariables.Top = 120
    lstVariables.Left = 120
    lstVariables.Width = Me.ScaleWidth - 120
    lstVariables.Height = Me.ScaleHeight - 240
    Err.Clear
End Sub
Private Sub lstVariables_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    LstViewSwapSort lstVariables, ColumnHeader
    Err.Clear
End Sub
Public Function Code_CommentCode(VbCp As VBIDE.VBComponent, StartLine As Long, EndLine As Long, Optional bOperation As Boolean = True)
    On Error Resume Next
    Dim xCnt As Long
    Dim xLine As String
    For xCnt = StartLine To EndLine
        xLine = Trim$(VbCp.CodeModule.lines(xCnt, 1))
        Select Case bOperation
        Case True
            Do Until Left$(xLine, 1) <> "'"
                xLine = Mid$(xLine, 2)
            Loop
            xLine = "'" & xLine
        Case False
            Do Until Left$(xLine, 1) <> "'"
                xLine = Mid$(xLine, 2)
            Loop
        End Select
        VbCp.CodeModule.ReplaceLine xCnt, xLine
        DoEvents
        Err.Clear
    Next
    Err.Clear
End Function
Private Sub lstVariables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then PopupMenu mnuFile
    Err.Clear
End Sub
Private Sub mnuClose_Click()
    On Error Resume Next
    cmdClose_Click
    Err.Clear
End Sub
Private Sub mnuComment_Click()
    On Error Resume Next
    cmdRemove_Click
    Err.Clear
End Sub
Private Sub mnuToExcel_Click()
    On Error Resume Next
    ToExcel_Click
    Err.Clear
End Sub
Private Sub ToExcel_Click()
    On Error Resume Next
    Dim xFile As String
    xFile = StringGetFileToken(VBInst.ActiveVBProject.Filename, "p") & "\Member & Variable Declarations.xls"
    LstViewToWorkSheetAsIs lstVariables, xFile, "Member & Variable Declarations", "Kimmo - CoolCode"
    Err.Clear
End Sub
