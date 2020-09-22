VERSION 5.00
Begin VB.Form frmComment 
   Caption         =   "Add A Comment"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDate 
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
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtInitials 
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
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   8775
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   8775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
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
      Left            =   8520
      TabIndex        =   1
      Top             =   2640
      Width           =   1155
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Apply"
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
      Left            =   7320
      TabIndex        =   0
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developer"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
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
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 2 Then
        WindowState = 0
    End If
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    txtDate.Text = Format$(Now(), "dd/mm/yyyy hh:mm:ss ampm")
    txtComment.Text = vbNullString
    txtInitials.Text = GetSetting(App_Name, "Developer", "name", vbNullString)
    Err.Clear
End Sub
Private Sub OKButton_Click()
    On Error Resume Next
    '--------------------------------------------------------------------------------
    ' Description: This adds a comment at the specified position in your code
    ' Created by : Anele Mbanga
    ' Date-Time  : 25/01/2005 11:28:22 PM
    '--------------------------------------------------------------------------------
    If boolIsBlank(txtComment, "comment") = True Then Exit Sub
    If boolIsBlank(txtInitials, "developer name") = True Then Exit Sub
    If boolIsBlank(txtDate, "date") = True Then Exit Sub
    Dim m As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim z As String
    z = "'--------------------------------------------------------------------------------" & vbCrLf
    z = z & "' Description: " & MvComment(txtComment.Text, vbNewLine, 2) & vbCrLf
    z = z & "' Created by : " & txtInitials.Text & vbCrLf
    z = z & "' Date-Time  : " & txtDate.Text & vbCrLf
    z = z & "'--------------------------------------------------------------------------------"
    VBInst.ActiveCodePane.GetSelection m, n, X, Y
    VBInst.ActiveCodePane.CodeModule.InsertLines m, z
    SaveSetting App_Name, "Developer", "name", txtInitials.Text
    Unload Me
    Err.Clear
End Sub
