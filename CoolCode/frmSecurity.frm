VERSION 5.00
Begin VB.Form frmSecurity 
   Caption         =   "Security Builder"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7440
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   435
      Left            =   6120
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtOutPut 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmSecurity.frx":0000
      Top             =   3240
      Width           =   8535
   End
   Begin VB.TextBox txtSelected 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmSecurity.frx":0006
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   510
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Text"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub

Private Sub cmdProcess_Click()
    On Error Resume Next
    txtOutPut.Text = ""
    If Len(txtSelected.Text) = 0 Then
        MsgBox "Please enter the text to process."
        Exit Sub
    End If
    txtOutPut.Text = SelectionOutput(txtSelected.Text)
    Err.Clear
        
End Sub

Private Function SelectionOutput(ByVal strValue As String) As String
    On Error Resume Next
    Dim strParent As String
    Dim strDescription As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLines() As String
    Dim rsLine As String
    Dim strOut As String
    Dim strKey As String
    
    strOut = ""
    strValue = Replace$(strValue, Quote, "")
    spLines = Split(strValue, vbNewLine)
    rsTot = UBound(spLines)
    For rsCnt = 0 To rsTot
        rsLine = Trim$(spLines(rsCnt))
        If InStr(1, rsLine, ".AddGroup") > 0 Then
            strParent = MvField(rsLine, 2, " ")
            strParent = MvField(strParent, 1, ",")
            strDescription = MvField(rsLine, 2, ",")
            strOut = strOut & "treeSecurity.Nodes.Add , ," & Quote & strParent & Quote & ", " & Quote & StringProperCase(strDescription) & Quote & ", " & Quote & "key" & Quote & ", " & Quote & "key" & Quote
        ElseIf InStr(1, rsLine, ".AddItem") > 0 Then
            strParent = MvField(rsLine, 2, " ")
            strParent = MvField(strParent, 1, ",")
            strKey = MvField(rsLine, 2, ",")
            strDescription = MvField(rsLine, 3, ",")
            strDescription = MvField(strDescription, 1, ",")
            strOut = strOut & vbNewLine & _
            "treeSecurity.Nodes.Add " & Quote & strParent & Quote & ", tvwChild, " & Quote & strKey & Quote & ", " & Quote & StringProperCase(strDescription) & Quote & ", " & Quote & "key" & Quote & ", " & Quote & "key" & Quote
        End If
    Next
    SelectionOutput = strOut
    Err.Clear
End Function



