VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WPF Ribbon Creator"
   ClientHeight    =   8160
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13185
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13996
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Convert Existing Code"
      TabPicture(0)   =   "frmConvert.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdConvertExisting"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdMenus"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSource"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTarget"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdClear"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPaste"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdButtons"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdResources"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCode"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCopy"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Create"
      TabPicture(1)   =   "frmConvert.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11640
         TabIndex        =   5
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCode 
         Caption         =   "Code"
         Height          =   375
         Left            =   10440
         TabIndex        =   10
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdResources 
         Caption         =   "Resources"
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdButtons 
         Caption         =   "Buttons"
         Height          =   375
         Left            =   8040
         TabIndex        =   8
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.TextBox txtTarget 
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   4080
         Width           =   12735
      End
      Begin VB.TextBox txtSource 
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   480
         Width           =   12735
      End
      Begin VB.CommandButton cmdMenus 
         Caption         =   "Menus"
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         ToolTipText     =   "Copy target to clipboard"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdConvertExisting 
         Caption         =   "Convert"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Code"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strResult As String
Private strCommands As String
Private strCode As String
Private strMenuCommands As String
Private strMenuCode As String
Private strMenuButtons As String

Private Sub cmdButtons_Click()
    On Error Resume Next
    txtTarget.Text = strResult
    Err.Clear
End Sub
Private Sub cmdClear_Click()
    On Error Resume Next
    txtSource.Text = ""
    txtTarget.Text = ""
    txtSource.SetFocus
    Err.Clear
End Sub
Private Sub cmdCode_Click()
    On Error Resume Next
    txtTarget.Text = strCode
    Err.Clear
End Sub
Private Sub cmdConvertExisting_Click()
    On Error Resume Next
    Dim StrSource As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim rsStr As String
    Dim thisLine() As String
    'Dim sID As String
    'Dim sCaption As String
    Dim sAction As String
    'Dim sCategory As String
    Dim colBottom As Collection
    Dim hasButton As Boolean
    Dim hasGroup As Boolean
    Dim strCommand As String
    Dim hasTab As Boolean
    StrSource = vbNewLine & Trim$(txtSource.Text)
    If Len(StrSource) = 0 Then
        MsgBox "Please specify the source code to convert first.", vbOKOnly + vbSystemModal, "Source Code Error"
        Err.Clear
        Exit Sub
    End If
    txtTarget.Text = ""
    strResult = ""
    strCommands = ""
    strCode = ""
    strMenuCommands = ""
    strMenuCode = ""
    strMenuButtons = ""
    
    spLine = Split(StrSource, vbNewLine)
    rsTot = UBound(spLine)
    Set colBottom = New Collection
    hasGroup = False
    hasButton = False
    For rsCnt = 1 To rsTot
        rsStr = Trim$(spLine(rsCnt))
        If Len(rsStr) = 0 Then GoTo NextCode
        If Left$(rsStr, 1) = "." Or Left$(rsStr, 1) = "'" Then
            ' line is within a with statement or is a comment
        Else
            thisLine = Split(rsStr, ",")
            sAction = Trim$(thisLine(0))
            ' search for addtab
            If InStr(1, sAction, "AddTab ") > 0 Then
                hasTab = True
                If hasButton = True Then
                    ' all the buttons in the category have been addded, then close the category group
                    strResult = strResult & "</r:RibbonGroup>" & vbNewLine
                    hasButton = False
                End If
                If hasGroup = True Then
                    strResult = strResult & "</r:RibbonTab>" & vbNewLine
                    hasGroup = False
                End If
                'source            0                1               2
                'PSR.AddTab "getting.started", "Getting Started", True
                strResult = strResult & "<r:RibbonTab Label=" & Trim$(thisLine(1)) & " FontFamily=" & Chr$(34) & "Tahoma" & Chr$(34) & ">" & vbNewLine
            ElseIf InStr(1, sAction, "AddCat ") > 0 Then
                hasGroup = True
                ' source 0                            1              2           3
                'PSR.AddCat "my.portfolios", "getting.started", "My Portfolios", False, ""
                If hasButton = True Then
                    ' all the buttons in the category have been addded, then close the category group
                    strResult = strResult & "</r:RibbonGroup>" & vbNewLine
                    hasButton = False
                End If
                strCommand = Trim$(Split(thisLine(0), " ")(1))
                strCommand = Replace$(strCommand, Chr$(34), "")
                strCommand = Replace$(strCommand, ".", "_")
                strResult = strResult & "<r:RibbonGroup Name=" & Replace$(Trim$(Split(thisLine(0), " ")(1)), ".", "_") & " HasDialogLauncher=" & Chr$(34) & Trim$(thisLine(3)) & Chr$(34) & ">" & vbNewLine
                strResult = strResult & vbTab & vbTab & "<r:RibbonGroup.Command>" & vbNewLine
                strResult = strResult & vbTab & vbTab & vbTab & "<r:RibbonCommand LabelTitle=" & Trim$(thisLine(2)) & " CanExecute=" & Chr$(34) & "OnCanExecute" & Chr$(34) & " Executed=" & Chr$(34) & "GroupClick_" & strCommand & Chr$(34) & " />" & vbNewLine
                strResult = strResult & vbTab & vbTab & "</r:RibbonGroup.Command>" & vbNewLine
                ' create code for groupclick
                strCode = strCode & "private void GroupClick_" & strCommand & "(object Sender, ExecutedRoutedEventArgs E)" & vbNewLine
                strCode = strCode & "{" & vbNewLine
                strCode = strCode & "// do nothing" & vbNewLine
                strCode = strCode & "}" & vbNewLine
            ElseIf InStr(1, sAction, "AddButton ") > 0 Then
                hasButton = True
                ' source         0                     1            2          3           4              5                 6
                'PSR.AddButton "portfolio_new1", "my.portfolios", "New", "newsomething", False, "Create a new portfolio", False
                strCommand = Trim$(Split(thisLine(0), " ")(1))
                strCommand = Replace$(strCommand, Chr$(34), "")
                strCommand = Replace$(strCommand, ".", "_")
                strResult = strResult & "<r:RibbonButton Name=" & Replace$(Trim$(Split(thisLine(0), " ")(1)), ".", "_") & _
                            " Command=" & Chr$(34) & "{StaticResource " & strCommand & "}" & Chr$(34) & " FontFamily=" & Chr$(34) & "Tahoma" & Chr$(34) & " FontSize=" & Chr$(34) & "11" & Chr$(34) & " />" & vbNewLine
                ' create commands for resources
                strCommands = strCommands & "<r:RibbonCommand x:Key=" & Chr$(34) & strCommand & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LabelTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipDescription=" & Trim$(thisLine(5)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "SmallImageSource=" & Chr$(34) & "Images\" & Replace$(Trim$(thisLine(3)), Chr$(34), "") & ".gif" & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LargeImageSource=" & Chr$(34) & "Images\" & Replace$(Trim$(thisLine(3)), Chr$(34), "") & ".gif" & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "CanExecute=" & Chr$(34) & "OnCanExecute" & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "Executed=" & Chr$(34) & "ButtonClick_" & strCommand & Chr$(34) & "/>" & vbNewLine
                ' create code for buttonclick
                strCode = strCode & "private void ButtonClick_" & strCommand & "(object sender, ExecutedRoutedEventArgs e)" & vbNewLine
                strCode = strCode & "{" & vbNewLine
                strCode = strCode & "// do nothing" & vbNewLine
                strCode = strCode & "}" & vbNewLine
                
                ' check if last line
                If rsCnt = rsTot Then
                    If hasGroup = True Then strResult = strResult & "</r:RibbonGroup>" & vbNewLine
                    If hasTab = True Then strResult = strResult & "</r:RibbonTab>" & vbNewLine
                End If
            ElseIf InStr(1, sAction, "AddComboBox ") > 0 Then
                hasButton = True
                ' adding a combo box
                '             0                               1             2                   3                  4            6
                'PSR.AddComboBox "fileserver_year", "fileserver_reports", "Year", "Year file was modified", "fileserver_year", 1000
                strCommand = Trim$(Split(thisLine(0), " ")(1))
                strCommand = Replace$(strCommand, Chr$(34), "")
                strCommand = Replace$(strCommand, ".", "_")
                strResult = strResult & "<r:RibbonComboBox Name=" & Replace$(Trim$(Split(thisLine(0), " ")(1)), ".", "_") & " Width=" & Chr(34) & thisLine(6) & Chr$(34) & _
                            " Command=" & Chr$(34) & "{StaticResource " & strCommand & "}" & Chr$(34) & " FontFamily=" & Chr$(34) & "Tahoma" & Chr$(34) & " FontSize=" & Chr$(34) & "11" & Chr$(34) & " />" & vbNewLine
                ' create commands for resources
                strCommands = strCommands & "<r:RibbonCommand x:Key=" & Chr$(34) & strCommand & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LabelTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipDescription=" & Trim$(thisLine(3)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "CanExecute=" & Chr$(34) & "OnCanExecute" & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "Executed=" & Chr$(34) & "ComboClick_" & strCommand & Chr$(34) & "/>" & vbNewLine
                ' create code for buttonclick
                strCode = strCode & "private void ComboClick_" & strCommand & "(object sender, ExecutedRoutedEventArgs e)" & vbNewLine
                strCode = strCode & "{" & vbNewLine
                strCode = strCode & "// do nothing" & vbNewLine
                strCode = strCode & "}" & vbNewLine
            
                If rsCnt = rsTot Then
                    If hasGroup = True Then strResult = strResult & "</r:RibbonGroup>" & vbNewLine
                    If hasTab = True Then strResult = strResult & "</r:RibbonTab>" & vbNewLine
                End If
            
            ElseIf InStr(1, sAction, "AddTextBox ") > 0 Then
                hasButton = True
                ' adding a textbox
                '              0                             1              2              3                4              5
                'PSR.AddTextBox "fileserver_find", "fileserver_manager", "Find", "Find in file names", "fileserver_find", 1500
                strCommand = Trim$(Split(thisLine(0), " ")(1))
                strCommand = Replace$(strCommand, Chr$(34), "")
                strCommand = Replace$(strCommand, ".", "_")
                strResult = strResult & "<r:RibbonTextBox Name=" & Replace$(Trim$(Split(thisLine(0), " ")(1)), ".", "_") & " Width=" & Chr(34) & thisLine(5) & Chr$(34) & _
                            " Command=" & Chr$(34) & "{StaticResource " & strCommand & "}" & Chr$(34) & " FontFamily=" & Chr$(34) & "Tahoma" & Chr$(34) & " FontSize=" & Chr$(34) & "11" & Chr$(34) & " />" & vbNewLine
                ' create commands for resources
                strCommands = strCommands & "<r:RibbonCommand x:Key=" & Chr$(34) & strCommand & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LabelTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipTitle=" & Trim$(thisLine(2)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipDescription=" & Trim$(thisLine(3)) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "CanExecute=" & Chr$(34) & "OnCanExecute" & Chr$(34) & vbNewLine
                strCommands = strCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "Executed=" & Chr$(34) & "TextBoxClick_" & strCommand & Chr$(34) & "/>" & vbNewLine
                ' create code for buttonclick
                strCode = strCode & "private void TextBoxClick_" & strCommand & "(object sender, ExecutedRoutedEventArgs e)" & vbNewLine
                strCode = strCode & "{" & vbNewLine
                strCode = strCode & "// do nothing" & vbNewLine
                strCode = strCode & "}" & vbNewLine
            
                If rsCnt = rsTot Then
                    If hasGroup = True Then strResult = strResult & "</r:RibbonGroup>" & vbNewLine
                    If hasTab = True Then strResult = strResult & "</r:RibbonTab>" & vbNewLine
                End If
            
            ElseIf InStr(1, sAction, "AddTopButton ") > 0 Then
                '                0                   1         2             3
                'PSR.AddTopButton "databases", "Databases", "table", "Database management"
                strCommand = Trim$(Split(thisLine(0), " ")(1))
                strCommand = Replace$(strCommand, Chr$(34), "")
                strCommand = Replace$(strCommand, ".", "_")
                
                ' menu buttons for the quick access toolbar
                strMenuButtons = strMenuButtons & "<r:RibbonButton Name=" & Replace$(Trim$(Split(thisLine(0), " ")(1)), ".", "_") & _
                                " Command=" & Chr$(34) & "{StaticResource " & strCommand & "}" & Chr$(34) & " />" & vbNewLine
                ' menu buttons for the resources section
                strMenuCommands = strMenuCommands & "<r:RibbonCommand x:Key=" & Chr$(34) & strCommand & Chr$(34) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LabelTitle=" & Trim$(thisLine(1)) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LabelDescription=" & Trim$(thisLine(3)) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipTitle=" & Trim$(thisLine(1)) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "ToolTipDescription=" & Trim$(thisLine(3)) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "SmallImageSource=" & Chr$(34) & "Images\" & Replace$(Trim$(thisLine(2)), Chr$(34), "") & ".gif" & Chr$(34) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "LargeImageSource=" & Chr$(34) & "Images\" & Replace$(Trim$(thisLine(2)), Chr$(34), "") & ".gif" & Chr$(34) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "CanExecute=" & Chr$(34) & "OnCanExecute" & Chr$(34) & vbNewLine
                strMenuCommands = strMenuCommands & vbTab & vbTab & vbTab & vbTab & vbTab & "Executed=" & Chr$(34) & "MenuClick_" & strCommand & Chr$(34) & "/>" & vbNewLine
                
                ' create code for menuclick
                strMenuCode = strMenuCode & "private void MenuClick_" & strCommand & "(object Sender, ExecutedRoutedEventArgs E)" & vbNewLine
                strMenuCode = strMenuCode & "{" & vbNewLine
                strMenuCode = strMenuCode & "// do nothing" & vbNewLine
                strMenuCode = strMenuCode & "}" & vbNewLine
            
            End If
        End If
NextCode:
        Err.Clear
    Next
    ' check if we have a tab that is not closed
    If InStr(1, strResult, "<r:RibbonTab Label=") > 0 And InStr(1, strResult, "</r:RibbonTab>") = 0 Then
        strResult = strResult & "</r:RibbonTab>" & vbNewLine
    End If
    ' check if we have a group that is not closed
    If InStr(1, strResult, "<r:RibbonGroup Name=") > 0 And InStr(1, strResult, "</r:RibbonGroup>") = 0 Then
        strResult = strResult & "</r:RibbonGroup>" & vbNewLine
    End If
    txtTarget.Text = strResult & vbNewLine & vbNewLine & _
    "--------------- RibbonCommand Code ------------------" & vbNewLine & vbNewLine & strCommands & vbNewLine & vbNewLine & _
    "--------------- Execution Code ----------------------" & vbNewLine & vbNewLine & strCode & vbNewLine & vbNewLine & _
    "--------------- Menu Buttons ------------------------" & vbNewLine & vbNewLine & strMenuButtons & vbNewLine & vbNewLine & _
    "--------------- Menu Commands -----------------------" & vbNewLine & vbNewLine & strMenuCommands & vbNewLine & vbNewLine & _
    "--------------- Menu Code ---------------------------" & vbNewLine & vbNewLine & strMenuCode
    
    cmdCopy.Enabled = True
    cmdButtons.Enabled = True
    cmdResources.Enabled = True
    cmdCode.Enabled = True
    cmdMenus.Enabled = True
    Err.Clear
End Sub
Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.SetText txtTarget.Text
    Err.Clear
End Sub

Private Sub cmdMenus_Click()
    On Error Resume Next
    txtTarget.Text = "--------------- Menu Buttons ------------------------" & vbNewLine & vbNewLine & strMenuButtons & vbNewLine & vbNewLine & _
    "--------------- Menu Commands -----------------------" & vbNewLine & vbNewLine & strMenuCommands & vbNewLine & vbNewLine & _
    "--------------- Menu Code ---------------------------" & vbNewLine & vbNewLine & strMenuCode
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    On Error Resume Next
    txtSource.Text = Clipboard.GetText
    Err.Clear
End Sub
Private Sub cmdResources_Click()
    On Error Resume Next
    txtTarget.Text = strCommands
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    cmdCopy.Enabled = False
    cmdConvertExisting.Enabled = False
    cmdButtons.Enabled = False
    cmdResources.Enabled = False
    cmdCode.Enabled = False
    cmdMenus.Enabled = False
    Err.Clear
End Sub
Private Sub txtSource_Change()
    On Error Resume Next
    cmdCopy.Enabled = False
    If Len(txtSource.Text) > 0 Then
        cmdConvertExisting.Enabled = True
    Else
        cmdConvertExisting.Enabled = False
    End If
    Err.Clear
End Sub
