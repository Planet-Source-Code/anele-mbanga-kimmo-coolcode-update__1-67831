VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFonts 
   Caption         =   "Set Form Font"
   ClientHeight    =   6825
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
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFontSize 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txtFontName 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   4215
   End
   Begin VB.CommandButton selFont 
      Caption         =   "Set Font"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Close screen"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Close screen"
      Top             =   6360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dFont 
      Left            =   2160
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Close screen"
      Top             =   6360
      Width           =   1215
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
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "List of controls having tooltips"
      Top             =   360
      Width           =   5412
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   660
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   780
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controls that can have Font Property"
      Height          =   204
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2760
   End
End
Attribute VB_Name = "frmFonts"
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
    list.Clear
    txtFontName.Text = ""
    txtFontSize.Text = ""
    'load the component
    Set mcmpCurrentForm = VBInst.SelectedVBComponent
    'check to see if we have a valid component
    If TypeName(mcmpCurrentForm) = "Nothing" Then
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
        If IsPropertyAvailable(ctl.Properties!Font) = False Then GoTo SkipIt
        sTmp = ControlName(ctl)
        ReDim Preserve controlarray(0 To UBound(controlarray) + 1) As VBControl
        i = i + 1
        Set controlarray(UBound(controlarray)) = ctl
        'add it to the list
        list.AddItem sTmp
        list.ItemData(list.NewIndex) = i
        list.Refresh
SkipIt:
    Next
    Err.Clear
End Sub
Private Function ControlName(ctl As vbide.VBControl) As String
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
Private Sub cmdApply_Click()
    On Error GoTo ErrorReport
    If boolIsBlank(txtFontName, "font name") = True Then Exit Sub
    If boolIsBlank(txtFontSize, "font size") = True Then Exit Sub
    Dim ctl As VBControl
    Dim f As New StdFont
    Dim r As New StdFont
    Dim p As Property
    f.Name = txtFontName.Text
    f.Size = Val(txtFontSize.Text)
    For Each ctl In VBInst.SelectedVBComponent.Designer.VBControls
        r = ctl.Properties!Font
        For Each p In r.pr
        ctl.Properties!Font.Name = f.Name
        ctl.Properties!Font.Size = f.Size
        ctl.Refresh
        'ctl.Properties!FontName = txtFontName.Text
        'ctl.Properties!FontSize = Val(txtFontSize.Text)
    Next
    Err.Clear
    Exit Sub
ErrorReport:
    MsgBox Err.Number & vbCr & Err.Description, vbOKOnly + vbExclamation, Err.Source
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    list.Clear
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 2 Then
        WindowState = 0
    End If
    Err.Clear
End Sub
Private Sub selFont_Click()
    On Error Resume Next
    With dFont
        .flags = 0
        .FontName = "Tahoma"
        .FontSize = 8
        .flags = cdlCFBoth + cdlCFEffects
        .ShowFont
    End With
    txtFontName.Text = dFont.FontName
    txtFontSize.Text = dFont.FontSize
    Err.Clear
End Sub

