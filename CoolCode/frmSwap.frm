VERSION 5.00
Begin VB.Form frmSwap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swap Delimited Text"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSwap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7905
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox txtDelimiter 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "="
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Paste"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Copy table"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox txtSwap 
         Height          =   6735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   600
         Width           =   10455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Swap"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Copy table"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Done"
         Height          =   375
         Left            =   9240
         TabIndex        =   2
         ToolTipText     =   "Copy table"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Copy table"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delimiter"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdL_Click()
    On Error Resume Next
    Err.Clear
End Sub
Private Sub Command1_Click()
    On Error Resume Next
     Clipboard.Clear
    Clipboard.SetText txtSwap.Text
    Err.Clear
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    txtSwap.Text = Trim$(txtSwap.Text)
    txtSwap.Text = StringRemAllNL(txtSwap.Text)
    If Len(txtSwap.Text) = 0 Then Exit Sub
    Dim sNew() As String
    Dim xNew() As String
    Dim sCnt As Integer
    Dim Strr As String
    Dim SPos As Integer
    Dim EPos As Integer
    Dim dSiz As Integer
    Dim SMid As String
    Dim Delim As String
    Dim strTab As String
    strTab = String$(4, " ")
    Delim = txtDelimiter.Text
    StringParse sNew, txtSwap.Text, vbNewLine
    For sCnt = 1 To UBound(sNew)
        StringParse xNew, sNew(sCnt), Delim
        ReDim Preserve xNew(2)
        xNew(1) = StringClean(xNew(1))
        xNew(2) = StringClean(xNew(2))
        sNew(sCnt) = xNew(2) & " " & Delim & " " & xNew(1)
        sNew(sCnt) = Trim$(sNew(sCnt))
        Strr = sNew(sCnt)
        SPos = InStr(Strr, "'")
        EPos = InStr(Strr, Delim)
        If SPos > 0 Then
            Select Case SPos
            Case Is > EPos
                dSiz = SPos - EPos
                SMid = Mid$(Strr, SPos, dSiz)
                Mid$(Strr, SPos, dSiz) = ""
                sNew(sCnt) = Strr
            Case Else
                Strr = Left$(sNew(sCnt), SPos - 1) & Mid$(sNew(sCnt), EPos) & strTab & Mid$(sNew(sCnt), SPos, (EPos - SPos))
                sNew(sCnt) = Strr
            End Select
        End If
    Next
    Strr = MvFromArray(sNew, vbNewLine)
    txtSwap.Text = Strr
    Strr = vbNullString
    Err.Clear
End Sub
Private Sub Command4_Click()
    On Error Resume Next
        txtSwap.Text = Clipboard.GetText
    Err.Clear
End Sub
