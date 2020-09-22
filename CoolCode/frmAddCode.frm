VERSION 5.00
Begin VB.Form frmAddCode 
   Caption         =   "Add Source Code"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   10440
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFonts 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Component(s)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton optCurrent 
         Caption         =   "Current Component"
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
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Process current component"
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All Components"
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
         Left            =   2880
         TabIndex        =   10
         ToolTipText     =   "Process all components"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdApply 
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
      Left            =   7800
      TabIndex        =   8
      ToolTipText     =   "Apply the insertion"
      Top             =   2040
      Width           =   1215
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
      Left            =   9120
      TabIndex        =   7
      ToolTipText     =   "Close screen"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtInsertion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "The code to insert"
      Top             =   1560
      Width           =   8895
   End
   Begin VB.TextBox txtLineContaining 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "The line to search for"
      Top             =   1080
      Width           =   8895
   End
   Begin VB.Frame fra 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton optAfter 
         Caption         =   "Insert Code After"
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
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Insert the code after the line containing text"
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optBefore 
         Caption         =   "Insert Code Before"
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
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Insert the text before the line containing"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insertion"
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
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line Containing"
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
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ClearAll()
    On Error Resume Next
    optAfter.Value = False
    optAll.Value = False
    optBefore.Value = False
    optCurrent.Value = False
    txtInsertion.Text = vbNullString
    txtLineContaining.Text = vbNullString
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 2 Then
        WindowState = 0
    End If
    Err.Clear
End Sub
Private Sub cmdApply_Click()
    On Error Resume Next
    txtLineContaining.Text = Trim$(txtLineContaining.Text)
    txtInsertion.Text = Trim$(txtInsertion.Text)
    If boolIsBlank(Me.txtLineContaining, "line containing") = True Then Exit Sub
    If boolIsBlank(txtInsertion, "text to insert") = True Then Exit Sub
    SaveReg App_Name, "addcode", txtInsertion.Text, txtLineContaining.Text
    SaveReg App_Name, "location", IIf((optBefore.Value = True), "b", "a"), txtLineContaining.Text
    Screen.MousePointer = vbHourglass
    If optCurrent.Value = True Then
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        If optBefore.Value = True Then
            Code_InsertBeforeAfter VbCp, txtLineContaining.Text, txtInsertion.Text
        Else
            Code_InsertBeforeAfter VbCp, txtLineContaining.Text, txtInsertion.Text, "a"
        End If
    Else
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            If optBefore.Value = True Then
                Code_InsertBeforeAfter VbCp, txtLineContaining.Text, txtInsertion.Text
            Else
                Code_InsertBeforeAfter VbCp, txtLineContaining.Text, txtInsertion.Text, "a"
            End If
NextModule:
            DoEvents
            Err.Clear
        Next
    End If
EndProc:
    Screen.MousePointer = vbDefault
    Me.Hide
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
