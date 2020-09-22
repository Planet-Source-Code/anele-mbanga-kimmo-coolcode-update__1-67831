VERSION 5.00
Object = "{9CCD14D6-ABE0-44BF-8F04-29E59D2CEA5E}#5.0#0"; "POLARZIPLIGHT.DLL"
Begin VB.Form frmProcess 
   Caption         =   "What to Process..."
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin POLARZIPLIGHTLibCtl.ZIPLight ZIPLight1 
      Left            =   240
      OleObjectBlob   =   "frmProcess.frx":0E42
      Top             =   2040
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear Errors First"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Apply the insertion"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Apply the insertion"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton optProcess 
      Caption         =   "All Components"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.OptionButton optProcess 
      Caption         =   "Current Component"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.OptionButton optProcess 
      Caption         =   "Current Procedure"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdApply_Click()
    On Error Resume Next
    clearError = chkClear.Value
    Unload Me
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    WhatToProcess = 4
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    WhatToProcess = 4
    Err.Clear
End Sub
Private Sub optProcess_Click(Index As Integer)
    On Error Resume Next
    If optProcess(Index).Value = True Then
        WhatToProcess = Index
    End If
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 2 Then
        WindowState = 0
    End If
    Err.Clear
End Sub
