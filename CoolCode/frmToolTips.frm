VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmToolTips 
   Caption         =   "Form Controls Tooltips"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5700
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToolTips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2520
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      ToolTipText     =   "Close screen"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Control Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtToolTips 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         ToolTipText     =   "Type in your tooltip here"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tooltip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   $"frmToolTips.frx":0E42
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblCaptionControl 
         AutoSize        =   -1  'True
         Caption         =   "lblCaptionControl"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblNameControl 
         AutoSize        =   -1  'True
         Caption         =   "lblNameControl"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.ListBox list 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "List of controls having tooltips"
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controls that can have Tooltips:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2310
   End
End
Attribute VB_Name = "frmToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcmpCurrentForm As VBComponent      'current form
Dim mcolCtls        As VBControls       'form's controls
Dim controlarray() As VBControl              'link from list's listdata to the control name
Private Function IsPropertyAvailable(ObjControl As Variant) As Boolean
    On Error Resume Next
    Dim xProperty As Property
    Set xProperty = ObjControl
    If TypeName(xProperty) = "Nothing" Then
        IsPropertyAvailable = False
    Else
        IsPropertyAvailable = True
    End If
    Err.Clear
End Function
Public Sub UpdateAll()
    On Error Resume Next
    Dim i As Integer
    Dim ctl As VBControl
    Dim sTmp As String
    ClearAll
    'load the component
    Set mcmpCurrentForm = VBInst.SelectedVBComponent
    'check to see if we have a valid component
    If mcmpCurrentForm Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    'make sure the active component is a form, user control or property page
    If (mcmpCurrentForm.Type <> vbext_ct_VBForm) And (mcmpCurrentForm.Type <> vbext_ct_UserControl) And (mcmpCurrentForm.Type <> vbext_ct_DocObject) And (mcmpCurrentForm.Type <> vbext_ct_PropPage) Then
        Err.Clear
        Exit Sub
    End If
    Set mcolCtls = mcmpCurrentForm.Designer.VBControls
    ReDim controlarray(0 To 0) As VBControl
    i = 0
    For Each ctl In mcmpCurrentForm.Designer.VBControls
        If IsPropertyAvailable(ctl.Properties!ToolTipText) = False Then GoTo SkipIt
        'try to get the tooltiptext
        'ti = ctl.Properties!ToolTipText
        'If Err Then
        '    'doesn't have a tabindex
        '    GoTo SkipIt
        'End If
        sTmp = ControlName(ctl)
        ReDim Preserve controlarray(0 To UBound(controlarray) + 1) As VBControl
        i = i + 1
        Set controlarray(UBound(controlarray)) = ctl ' add it to the list
        list.AddItem sTmp
        list.ItemData(list.NewIndex) = i
        list.Refresh
SkipIt:
        Err.Clear
    Next
    Err.Clear
End Sub
Private Function ControlName(ctl As VBIDE.VBControl) As String
    On Error Resume Next
    Dim sTmp As String
    Dim sCaption As String
    Dim i As Integer
    If IsPropertyAvailable(ctl.Properties!Name) = True Then sTmp = ctl.Properties!Name
    If IsPropertyAvailable(ctl.Properties!Caption) = True Then sCaption = ctl.Properties!Caption
    'will be null if there isn't one
    i = ctl.Properties!Index
    If i >= 0 Then
        sTmp = sTmp & "(" & i & ")"
    End If
    If Len(sCaption) > 0 Then
        ControlName = sTmp & " - '" & sCaption & "'"
    Else
        ControlName = sTmp
    End If
    Err.Clear
End Function
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    ClearAll
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 2 Then
        WindowState = 0
    End If
    Err.Clear
End Sub
Private Sub list_Click()
    On Error Resume Next
    Dim ctl As VBIDE.VBControl
    Set ctl = controlarray(list.ItemData(list.ListIndex))
    lblCaptionControl.Caption = vbNullString
    If IsPropertyAvailable(ctl.Properties!Name) = True Then lblNameControl.Caption = ctl.Properties!Name
    If IsPropertyAvailable(ctl.Properties!Caption) = True Then lblCaptionControl.Caption = ctl.Properties!Caption
    txtToolTips.Text = ctl.Properties!ToolTipText
    txtToolTips.Enabled = True
    Err.Clear
End Sub
Public Sub ClearAll()
    On Error Resume Next
    list.Clear
    lblNameControl.Caption = vbNullString
    lblCaptionControl.Caption = vbNullString
    txtToolTips.Text = vbNullString
    txtToolTips.Enabled = False
    Err.Clear
End Sub
Private Sub txtToolTips_Change()
    On Error Resume Next
    Dim ctl As VBIDE.VBControl
    Set ctl = controlarray(list.ItemData(list.ListIndex))
    If Err Then Exit Sub
    ctl.Properties!ToolTipText = txtToolTips.Text
    Err.Clear
End Sub
Private Sub txtToolTips_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        txtToolTips_Change
        If list.ListIndex < list.ListCount - 1 Then list.ListIndex = list.ListIndex + 1 Else list.ListIndex = 0
    End If
    Err.Clear
End Sub
