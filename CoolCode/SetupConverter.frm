VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SetupConverter 
   Caption         =   "Setup Inno"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SetupConverter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download Inno Setup"
      Height          =   375
      Left            =   1560
      TabIndex        =   81
      Top             =   7680
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   0
      TabIndex        =   80
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdInno 
      Caption         =   "Inno"
      Height          =   375
      Left            =   13680
      TabIndex        =   79
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   12120
      TabIndex        =   78
      Top             =   7680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progBar 
      Height          =   735
      Left            =   4560
      TabIndex        =   22
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin TabDlg.SSTab tabSetup 
      Height          =   7575
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13361
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Setup"
      TabPicture(0)   =   "SetupConverter.frx":044A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "SetupConverter.frx":0466
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(18)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstViewRun"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lstFiles"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Inno Setup File"
      TabPicture(2)   =   "SetupConverter.frx":0482
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "txtInno"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.PictureBox fra 
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   -74880
         ScaleHeight     =   6975
         ScaleWidth      =   14775
         TabIndex        =   23
         Top             =   480
         Width           =   14775
         Begin VB.CheckBox Check20 
            Caption         =   "Wizard Image Stretch"
            Height          =   315
            Left            =   10920
            TabIndex        =   77
            Tag             =   "WizardImageStretch"
            ToolTipText     =   $"SetupConverter.frx":049E
            Top             =   6600
            Width           =   3615
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Window Visible"
            Height          =   315
            Left            =   10920
            TabIndex        =   76
            Tag             =   "WindowVisible"
            ToolTipText     =   "If set to yes, there will be a gradient background window displayed behind the wizard."
            Top             =   6360
            Width           =   3615
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Window Start Maximized"
            Height          =   315
            Left            =   10920
            TabIndex        =   75
            Tag             =   "WindowStartMaximized"
            ToolTipText     =   $"SetupConverter.frx":0547
            Top             =   6120
            Width           =   3615
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Show Tasks TreeLines"
            Height          =   315
            Left            =   10920
            TabIndex        =   74
            Tag             =   "ShowTasksTreeLines"
            ToolTipText     =   "When this directive is set to yes, Setup will show 'tree lines' between parent and sub tasks."
            Top             =   5880
            Width           =   3615
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Show Component Sizes"
            Height          =   315
            Left            =   10920
            TabIndex        =   73
            Tag             =   "ShowComponentSizes"
            ToolTipText     =   $"SetupConverter.frx":05D8
            Top             =   5640
            Width           =   3615
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Flat Components List"
            Height          =   315
            Left            =   10920
            TabIndex        =   72
            Tag             =   "FlatComponentsList"
            ToolTipText     =   $"SetupConverter.frx":0699
            Top             =   5400
            Width           =   3615
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Back Solid"
            Height          =   315
            Left            =   10920
            TabIndex        =   71
            Tag             =   "BackSolid"
            ToolTipText     =   $"SetupConverter.frx":0722
            Top             =   5160
            Width           =   3615
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Use Previous Setup Type"
            Height          =   315
            Left            =   10920
            TabIndex        =   70
            Tag             =   "UsePreviousSetupType"
            ToolTipText     =   $"SetupConverter.frx":07E2
            Top             =   4920
            Width           =   3615
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Use Previous Group"
            Height          =   315
            Left            =   10920
            TabIndex        =   69
            Tag             =   "UsePreviousGroup"
            ToolTipText     =   $"SetupConverter.frx":0905
            Top             =   4680
            Width           =   3615
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Use Previous App Dir"
            Height          =   315
            Left            =   10920
            TabIndex        =   68
            Tag             =   "UsePreviousAppDir"
            ToolTipText     =   $"SetupConverter.frx":0AB6
            Top             =   4440
            Width           =   3615
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Uninstallable"
            Height          =   315
            Left            =   10920
            TabIndex        =   67
            Tag             =   "Uninstallable"
            ToolTipText     =   $"SetupConverter.frx":0BC2
            Top             =   4200
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Setup Logging"
            Height          =   315
            Left            =   10920
            TabIndex        =   66
            Tag             =   "SetupLogging"
            ToolTipText     =   "If set to yes, Setup will always create a log file. Equivalent to passing /LOG on the command line."
            Top             =   3960
            Width           =   3615
         End
         Begin VB.ComboBox cboMinVersion 
            Height          =   315
            ItemData        =   "SetupConverter.frx":0CD8
            Left            =   1920
            List            =   "SetupConverter.frx":0D0C
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   6240
            Width           =   7815
         End
         Begin VB.CommandButton SmallImage 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   24
            Top             =   4800
            Width           =   855
         End
         Begin VB.CommandButton License 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   26
            Top             =   3360
            Width           =   855
         End
         Begin VB.CommandButton InforAfter 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   27
            Top             =   3000
            Width           =   855
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Disable Startup Prompt"
            Height          =   315
            Left            =   10920
            TabIndex        =   64
            Tag             =   "DisableStartupPrompt"
            ToolTipText     =   "When this is set to yes, Setup will not show the This will install... Do you wish to continue? prompt."
            Top             =   3720
            Width           =   3615
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Disable Ready Page"
            Height          =   315
            Left            =   10920
            TabIndex        =   63
            Tag             =   "DisableReadyPage"
            ToolTipText     =   "If this is set to yes, Setup will not show the Ready to Install wizard page."
            Top             =   3480
            Width           =   3615
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Disable Ready Memo"
            Height          =   315
            Left            =   10920
            TabIndex        =   62
            Tag             =   "DisableReadyMemo"
            ToolTipText     =   $"SetupConverter.frx":0E9A
            Top             =   3240
            Width           =   3615
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Disable Program Group Page"
            Height          =   315
            Left            =   10920
            TabIndex        =   61
            Tag             =   "DisableProgramGroupPage"
            ToolTipText     =   $"SetupConverter.frx":0F6B
            Top             =   3000
            Width           =   3615
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Disable Finished Page"
            Height          =   315
            Left            =   10920
            TabIndex        =   60
            Tag             =   "DisableFinishedPage"
            ToolTipText     =   $"SetupConverter.frx":104B
            Top             =   2760
            Width           =   3615
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Changes Environment"
            Height          =   315
            Left            =   10920
            TabIndex        =   59
            Tag             =   "ChangesEnvironment"
            ToolTipText     =   $"SetupConverter.frx":11BA
            Top             =   2520
            Width           =   3615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Changes Associations"
            Height          =   315
            Left            =   10920
            TabIndex        =   58
            Tag             =   "ChangesAssociations"
            ToolTipText     =   $"SetupConverter.frx":127C
            Top             =   2280
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Always Show Components List"
            Height          =   315
            Left            =   10920
            TabIndex        =   57
            Tag             =   "AlwaysShowComponentsList"
            ToolTipText     =   "Setup will always show the components list for customizable setups"
            Top             =   2040
            Width           =   3615
         End
         Begin VB.CommandButton Icon 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   54
            Top             =   5880
            Width           =   855
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   16
            Left            =   1920
            TabIndex        =   16
            Top             =   5880
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   2
            Tag             =   "AppPublisher"
            Text            =   "txtSetup"
            Top             =   840
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   1
            Tag             =   "AppName"
            Text            =   "txtSetup"
            Top             =   480
            Width           =   7815
         End
         Begin VB.CheckBox chkShowGroup 
            Caption         =   "Always Show Group On Ready Page"
            Height          =   315
            Left            =   10920
            TabIndex        =   36
            Tag             =   "AlwaysShowGroupOnReadyPage"
            ToolTipText     =   $"SetupConverter.frx":1333
            Top             =   840
            Width           =   3375
         End
         Begin VB.CheckBox chkShowDir 
            Caption         =   "Always Show Directory Ready Page"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10920
            TabIndex        =   35
            Tag             =   "AlwaysShowDirOnReadyPage"
            ToolTipText     =   $"SetupConverter.frx":1440
            Top             =   600
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.CheckBox chkUNC 
            Caption         =   "Allow UNC Path"
            Height          =   315
            Left            =   10920
            TabIndex        =   34
            Tag             =   "AllowUNCPath"
            ToolTipText     =   $"SetupConverter.frx":14CE
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkAllowRoot 
            Caption         =   "Allow Root Directory"
            Height          =   315
            Left            =   10920
            TabIndex        =   33
            Tag             =   "AllowRootDirectory"
            ToolTipText     =   $"SetupConverter.frx":1562
            Top             =   120
            Width           =   2055
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   0
            Tag             =   "OutputBaseFileName"
            Text            =   "txtSetup"
            Top             =   120
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   3
            Tag             =   "AppPublisherURL"
            Text            =   "txtSetup"
            Top             =   1200
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   4
            Tag             =   "AppVersion"
            Text            =   "txtSetup"
            Top             =   1560
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   5
            Tag             =   "AppVerName"
            Text            =   "txtSetup"
            Top             =   1920
            Width           =   7815
         End
         Begin VB.CheckBox chkCreateDir 
            Caption         =   "Create Application Directory"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10920
            TabIndex        =   32
            Tag             =   "CreateAppDir"
            ToolTipText     =   $"SetupConverter.frx":1600
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   6
            Tag             =   "DefaultDirName"
            Text            =   "txtSetup"
            Top             =   2280
            Width           =   7815
         End
         Begin VB.CheckBox chkWarning 
            Caption         =   "Display Directory Existence Warning"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10920
            TabIndex        =   31
            Tag             =   "DirExistsWarning"
            ToolTipText     =   $"SetupConverter.frx":168E
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   7
            Left            =   1920
            TabIndex        =   7
            Tag             =   "InfoBeforeFile"
            Text            =   "txtSetup"
            Top             =   2640
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   8
            Left            =   1920
            TabIndex        =   8
            Tag             =   "InfoAfterFile"
            Text            =   "txtSetup"
            Top             =   3000
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   9
            Left            =   1920
            TabIndex        =   9
            Tag             =   "LicenseFile"
            Text            =   "txtSetup"
            Top             =   3360
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   10
            Left            =   1920
            TabIndex        =   10
            Tag             =   "Password"
            Text            =   "txtSetup"
            Top             =   3720
            Width           =   7815
         End
         Begin VB.CheckBox chkRestart 
            Caption         =   "Restart If Needed"
            Height          =   315
            Left            =   10920
            TabIndex        =   30
            Tag             =   "RestartIfNeededByRun"
            ToolTipText     =   $"SetupConverter.frx":1723
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CheckBox chkUserInfor 
            Caption         =   "Show User Information Page"
            Height          =   315
            Left            =   10920
            TabIndex        =   29
            Tag             =   "UserInfoPage"
            ToolTipText     =   $"SetupConverter.frx":1826
            Top             =   1560
            Width           =   2775
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   11
            Left            =   1920
            TabIndex        =   11
            Tag             =   "AppCopyright"
            Text            =   "txtSetup"
            Top             =   4080
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   12
            Left            =   1920
            TabIndex        =   12
            Tag             =   "WizardImageFile"
            Text            =   "txtSetup"
            Top             =   4440
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   13
            Left            =   1920
            TabIndex        =   13
            Tag             =   "WizardSmallImageFile"
            Text            =   "txtSetup"
            Top             =   4800
            Width           =   7815
         End
         Begin VB.CommandButton InforBef 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   28
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton ImageFile 
            Caption         =   "Select"
            Height          =   375
            Left            =   9720
            TabIndex        =   25
            Top             =   4440
            Width           =   855
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   14
            Left            =   1920
            TabIndex        =   14
            Tag             =   "DefaultGroupName"
            Text            =   "txtSetup"
            Top             =   5160
            Width           =   7815
         End
         Begin VB.TextBox txtSetup 
            Height          =   375
            Index           =   15
            Left            =   1920
            TabIndex        =   15
            Text            =   "txtSetup"
            Top             =   5520
            Width           =   7815
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Windows Version"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   65
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon File"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   53
            Top             =   5880
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Publisher"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   645
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   405
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Setup File"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Publisher URL"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version Name"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Directory Name"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   46
            Top             =   2280
            Width           =   1110
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information Before "
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   45
            Top             =   2640
            Width           =   1410
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information After"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   3000
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "License File"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   3360
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   3720
            Width           =   690
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyrights"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   4080
            Width           =   780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wizard Image"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Top             =   4440
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wizard Small Image"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   4800
            Width           =   1395
         End
         Begin VB.Line Line1 
            X1              =   10680
            X2              =   10680
            Y1              =   120
            Y2              =   6960
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group Name"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   38
            Top             =   5160
            Width           =   885
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Executable"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   37
            Top             =   5520
            Width           =   795
         End
      End
      Begin RichTextLib.RichTextBox txtInno 
         Height          =   7095
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   12515
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"SetupConverter.frx":18B9
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   8705
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Destination"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Shared"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstViewRun 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   55
         Top             =   5640
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   3201
         View            =   3
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Destination"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Shared"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files to add to the Run section (Check the files to add)"
         Height          =   195
         Index           =   18
         Left            =   -74880
         TabIndex        =   56
         Top             =   5400
         Width           =   3915
      End
   End
   Begin VB.ComboBox cboFiles 
      Height          =   315
      Left            =   12600
      Sorted          =   -1  'True
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   6600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   12720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuVbSetup 
         Caption         =   "Open Vb Setup File"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
      End
      Begin VB.Menu mnuInnoSetup 
         Caption         =   "Inno Setup"
      End
      Begin VB.Menu de 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "SetupConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrSource As String
Public SetupFile As String
Private colFiles As Collection
Private lastPath As String
Private StrIssFile As String
Private Sub cmdConvert_Click()
    On Error Resume Next
    mnuConvert_Click
    Err.Clear
End Sub

Private Sub cmdDownload_Click()
    On Error Resume Next
    boolViewFile "http://inno-setup.en.softonic.com"
    Err.Clear
End Sub

Private Sub cmdInno_Click()
    On Error Resume Next
    mnuInnoSetup_Click
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    mnuExit_Click
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    fra.BorderStyle = 0
    'CleanAllControls Me
    ClearAll
    tabSetup.Tab = 0
    'LstBoxFromMV cboFiles, ReadReg("lastfiles")
    'ReloadMenu mnuProjects, cboFiles
    cboMinVersion.AddItem "Windows 95"
    cboMinVersion.AddItem "Windows 95 OSR 2 & OSR 2.1"
    cboMinVersion.AddItem "Windows 95 OSR 2.5"
    cboMinVersion.AddItem "Windows 98"
    cboMinVersion.AddItem "Windows 98 Second Edition"
    cboMinVersion.AddItem "Windows Me"
    cboMinVersion.AddItem "Windows NT 4.0"
    cboMinVersion.AddItem "Windows 2000"
    cboMinVersion.AddItem "Windows XP"
    cboMinVersion.AddItem "Windows XP 64-Bit Edition Version 2002 (Itanium)"
    cboMinVersion.AddItem "Windows Server 2003"
    cboMinVersion.AddItem "Windows XP x64 Edition (AMD64/EM64T)"
    cboMinVersion.AddItem "Windows XP 64-Bit Edition Version 2003 (Itanium)"
    cboMinVersion.AddItem "Windows Vista"
    cboMinVersion.AddItem "Windows Vista with Service Pack 1"
    cboMinVersion.AddItem "Windows Server 2008"
    cboMinVersion.Text = "Windows XP"
    Err.Clear
End Sub
Private Sub Icon_Click()
    On Error Resume Next
    txtSetup(16).Text = DialogOpen(CD, "Select Icon File", lastPath, "*.ico")
    If Len(txtSetup(16).Text) > 0 Then
        Dim spRec(1 To 4) As String
        spRec(1) = txtSetup(16).Text
        spRec(2) = "{app}"
        LstViewUpdate spRec, lstFiles, vbNullString
        'LstViewRemoveDuplicates lstFiles
        LstViewAutoResize lstFiles
    End If
    Err.Clear
End Sub
Private Sub ImageFile_Click()
    On Error Resume Next
    txtSetup(12).Text = DialogOpen(CD, "Select Image File", lastPath, "*.bmp")
    Err.Clear
End Sub
Private Sub InforAfter_Click()
    On Error Resume Next
    txtSetup(8).Text = DialogOpen(CD, "Select Information After File", lastPath, "*.txt")
    Err.Clear
End Sub
Private Sub InforBef_Click()
    On Error Resume Next
    txtSetup(7).Text = DialogOpen(CD, "Select Information Before File", lastPath, "*.txt")
    Err.Clear
End Sub
'Private Sub Last_Click()
'    On Error Resume Next
'    If File_Exists(App.Path & "\FileList.txt") = True Then
'        Screen.MousePointer = vbHourglass
'        'LstViewFromFile lstFiles, App.Path & "\FileList.txt"
'        LstViewAutoResize lstFiles
'        Screen.MousePointer = vbDefault
'    End If
'    Err.Clear
'End Sub
Private Sub License_Click()
    On Error Resume Next
    txtSetup(9).Text = DialogOpen(CD, "Select License File", lastPath, "*.txt")
    Err.Clear
End Sub
Private Sub mnuExit_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub mnuInnoSetup_Click()
    On Error Resume Next
    Call boolViewFile(StrIssFile)
    Err.Clear
End Sub
Private Sub mnuVbSetup_Click()
    On Error Resume Next
    ReadVbSetupFile True
    Err.Clear
End Sub
Private Sub mnuConvert_Click()
    On Error Resume Next
    Dim intResp As Integer
    If chkShowDir.Value = 0 Then
        intResp = MsgBox("It is not recommended to disable this option. Enable this option?", vbYesNo + vbApplicationModal, chkShowDir.Caption)
        If intResp = vbYes Then chkShowDir.Value = 1
    End If
    If chkWarning.Value = 0 Then
        intResp = MsgBox("It is not recommended to disable this option. Enable this option?", vbYesNo + vbApplicationModal, chkWarning.Caption)
        If intResp = vbYes Then chkWarning.Value = 1
    End If
    If chkCreateDir.Value = 0 Then
        intResp = MsgBox("It is not recommended to disable this option. Enable this option?", vbYesNo + vbApplicationModal, chkCreateDir.Caption)
        If intResp = vbYes Then chkCreateDir.Value = 1
    End If
    If Check10.Value = 0 Then
        intResp = MsgBox("It is not recommended to disable this option. Enable this option?", vbYesNo + vbApplicationModal, Check10.Caption)
        If intResp = vbYes Then Check10.Value = 1
    End If
    If Check8.Value = 0 Then
        intResp = MsgBox("It is not recommended to disable this option. Enable this option?", vbYesNo + vbApplicationModal, Check8.Caption)
        If intResp = vbYes Then Check8.Value = 1
    End If
    tabSetup.Tab = 2
    txtInno.Text = vbNullString
    StrIssFile = App.Path & "\" & txtSetup(1).Text
    If boolDirExists(StrIssFile) = False Then MkDir StrIssFile
    StrIssFile = StrIssFile & "\" & txtSetup(1).Text & ".iss"
    Screen.MousePointer = vbHourglass
    Dim lstDetails As Collection
    Dim lstRow() As String
    Dim lstLine As String
    Dim strFlags As String
    Dim frmTot As Long
    Dim frmCnt As Long
    Dim frmTag As String
    Dim frmType As String
    Dim frmLine As String
    Set lstDetails = New Collection
    lstDetails.Add "; Script generated by the Inno Setup Script Wizard."
    lstDetails.Add "; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!"
    lstDetails.Add vbNullString
    lstDetails.Add "[Setup]"
    ' somehow does not work, exlcude
    Select Case Trim$(cboMinVersion.Text)
    Case "Windows 95"
        'lstDetails.Add "MinVersion=4.0"
    Case "Windows 95 OSR 2 & OSR 2.1"
        'lstDetails.Add "MinVersion=4.0"
    Case "Windows 95 OSR 2.5"
        'lstDetails.Add "MinVersion=4.0"
    Case "Windows 98"
        'lstDetails.Add "MinVersion=4.1"
    Case "Windows 98 Second Edition"
        'lstDetails.Add "MinVersion=4.1"
    Case "Windows Me"
        'lstDetails.Add "MinVersion=4.9"
    Case "Windows NT 4.0"
        'lstDetails.Add "MinVersion=4.0"
    Case "Windows 2000"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows XP"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows XP 64-Bit Edition Version 2002 (Itanium)"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows Server 2003"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows XP x64 Edition (AMD64/EM64T)"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows XP 64-Bit Edition Version 2003 (Itanium)"
        'lstDetails.Add "MinVersion=5.0"
    Case "Windows Vista"
        'lstDetails.Add "MinVersion=6.0"
    Case "Windows Vista with Service Pack 1"
        'lstDetails.Add "MinVersion=6.0"
    Case "Windows Server 2008"
        'lstDetails.Add "MinVersion=6.0"
    End Select
    frmTot = Me.Controls.Count - 1
    For frmCnt = 0 To frmTot
        frmType = TypeName(Me.Controls(frmCnt))
        frmTag = Me.Controls(frmCnt).Tag
        If frmTag = vbNullString Then GoTo NextEntry
        Select Case frmType
        Case "TextBox"
            If Me.Controls(frmCnt).Text <> vbNullString Then
                frmLine = frmTag & "=" & Me.Controls(frmCnt).Text
                lstDetails.Add frmLine
            End If
        Case "CheckBox"
            If Me.Controls(frmCnt).Value = 1 Then
                frmLine = frmTag & "=yes"
            Else
                frmLine = frmTag & "=no"
            End If
            lstDetails.Add frmLine
        End Select
NextEntry:
        Err.Clear
    Next
    lstDetails.Add vbNullString
    lstDetails.Add "[Tasks]"
    lstDetails.Add "Name: " & Quote & "desktopicon" & Quote & "; Description: " & Quote & "Create a &desktop icon" & Quote & "; GroupDescription: " & Quote & "Additional icons:" & Quote
    lstDetails.Add vbNullString
    lstDetails.Add "[Files]"
    frmTot = lstFiles.ListItems.Count
    For frmCnt = 1 To frmTot
        lstRow = LstViewGetRow(lstFiles, frmCnt)
        strFlags = LCase$(Trim$(lstRow(3) & " " & lstRow(4)))
        lstRow(1) = UCase$(lstRow(1))
        Select Case UCase$(lstRow(3))
        Case "REGSERVER", "REGTYPELIB"
            strFlags = strFlags & " noregerror"
        End Select
        If InStr(1, lstRow(1), "STKIT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "COMCAT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "ASYCFILT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "OLEPRO") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "OLEAUT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "STDOLE") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "MSVBVM") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If InStr(1, lstRow(1), "MSVCRT") > 0 Then strFlags = "uninsneveruninstall " & strFlags
        If lstRow(4) = "SHAREDFILE" Then
            If InStr(1, strFlags, "uninsneveruninstall") = 0 Then strFlags = "uninsneveruninstall " & strFlags
        End If
        lstLine = "Source: " & Quote & lstRow(1) & Quote & "; DestDir: " & Quote & lstRow(2) & Quote & "; Flags: " & strFlags
        ' check if its unsafe file
        Select Case LCase$(File_Token(lstRow(1), "f"))
        Case "advapi32.dll", "comdlg32.dll", "gdi32.dll", "kernel32.dll", "riched32.dll", "shell32.dll", "user32.dll", "uxtheme.dll", "comctl32.dll"
        Case "shdocvw.dll", "shlwapi.dll", "urlmon.dll", "wininet.dll", "ctl3d32.dll", "comcat.dll"
        Case Else
            lstDetails.Add lstLine
        End Select
        Err.Clear
    Next
    lstDetails.Add vbNullString
    lstDetails.Add ";NOTE: Don't use " & Quote & "Flags: ignoreversion" & Quote & " on any shared system files"
    lstDetails.Add vbNullString
    lstDetails.Add "[Icons]"
    lstDetails.Add "Name: " & Quote & "{group}\" & txtSetup(14).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote
    lstDetails.Add "Name: " & Quote & "{group}\Uninstall " & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{uninstallexe}" & Quote
    If txtSetup(16).Text = vbNullString Then
        lstDetails.Add "Name: " & Quote & "{userdesktop}\" & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; Tasks: desktopicon"
    Else
        lstDetails.Add "Name: " & Quote & "{userdesktop}\" & txtSetup(1).Text & Quote & "; Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; IconFilename: " & Quote & "{app}\" & File_Token(txtSetup(16).Text, "f") & Quote & "; Tasks: desktopicon"
    End If
    lstDetails.Add vbNullString
    lstDetails.Add "[Run]"
    lstDetails.Add "Filename: " & Quote & "{app}\" & txtSetup(15).Text & Quote & "; Description: " & Quote & "Launch " & txtSetup(1).Text & Quote & "; Flags: nowait postinstall skipifsilent"
    For frmCnt = 1 To lstViewRun.ListItems.Count
        If lstViewRun.ListItems(frmCnt).Checked = True Then
            If InStr(1, lstViewRun.ListItems(frmCnt).Text, ".msi", vbTextCompare) > 0 Then
                lstDetails.Add "Filename: " & Quote & lstViewRun.ListItems(frmCnt).ListSubItems(1) & "\" & lstViewRun.ListItems(frmCnt).Text & Quote & ";Parameters: " & Quote & "/q" & Quote & "; Flags: shellexec"
            Else
                lstDetails.Add "Filename: " & Quote & lstViewRun.ListItems(frmCnt).ListSubItems(1) & "\" & lstViewRun.ListItems(frmCnt).Text & Quote & ";Parameters: " & Quote & "/q" & Quote
            End If
        End If
        Err.Clear
    Next
    txtInno.Text = MvFromCollection(lstDetails, vbNewLine)
    txtInno.SaveFile StrIssFile, rtfText
    Screen.MousePointer = vbDefault
    Beep
    mnuInnoSetup_Click
    Err.Clear
End Sub
'Private Sub SaveList_Click()
'    On Error Resume Next
'    Screen.MousePointer = vbHourglass
'    LstViewToFile progBar, lstFiles, App.Path & "\FileList.txt"
'    Screen.MousePointer = vbDefault
'    Err.Clear
'End Sub
Private Sub SmallImage_Click()
    On Error Resume Next
    txtSetup(13).Text = DialogOpen(CD, "Select Small Image File", lastPath, "*.bmp")
    Err.Clear
End Sub
Private Sub txtSetup_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    Select Case Index
    Case 2
        txtSetup(11).Text = txtSetup(2).Text
    End Select
    Err.Clear
End Sub
Public Sub ReadVbSetupFile(Optional Prompt As Boolean = False)
    On Error Resume Next
    Dim txtCnt As Integer
    lastPath = App.Path
    If Prompt = True Then
        SetupFile = DialogOpen(CD, "Select VB Setup File", lastPath, "*.lst")
        If Len(SetupFile) = 0 Then Exit Sub
    End If
    Caption = "Setup Inno - " & SetupFile
    lastPath = File_Token(SetupFile, "p")
    SaveReg "lastpath", lastPath
    For txtCnt = 0 To txtSetup.Count - 1
        txtSetup(txtCnt).Text = vbNullString
        Err.Clear
    Next
    lstFiles.ListItems.Clear
    txtInno.Text = vbNullString
    StrSource = lastPath & "\Support"
    If boolDirExists(StrSource) = False Then
        MsgBox "The support directory for the setup files does not exists." & vbCr & "This directory is usually created by the VB setup program.", vbOKOnly + vbExclamation + vbApplicationModal, StrSource & " Error"
        Err.Clear
        Exit Sub
    End If
    LstBoxUpdate cboFiles, SetupFile
    lstViewRun.ListItems.Clear
    Screen.MousePointer = vbHourglass
    Set colFiles = New Collection
    Dim intFile As String
    Dim StrLine As String
    Dim equalPos As Long
    Dim sHead As String
    Dim sRest As String
    Dim sFile As String
    Dim sOper As String
    Dim sDest As String
    Dim sShared As String
    Dim arrLine() As String
    Dim lstLine(1 To 4) As String
    Dim tmpFile As String
    intFile = FreeFile
    Open SetupFile For Input Access Read As #intFile
    Do Until EOF(intFile)
        ' read each line
        Line Input #intFile, StrLine
        StrLine = Trim$(StrLine)
        If Len(StrLine) = 0 Then GoTo NextLine
        ' prefixes are separated by an equal sign
        equalPos = InStr(1, StrLine, "=")
        Select Case equalPos
        Case 0
            GoTo NextLine
        Case Else
            sHead = Left$(MvField(StrLine, 1, "="), 4)
            sRest = MvField(StrLine, 2, "=")
            Select Case LCase$(sHead)
            Case "file"
                StrParse arrLine, sRest, ","
                ReDim Preserve arrLine(4)
                arrLine(2) = LCase$(arrLine(2))
                arrLine(3) = LCase$(arrLine(3))
                arrLine(4) = LCase$(arrLine(4))
                sFile = Mid$(arrLine(1), 2)     ' read file name
                sDest = arrLine(2)
                sOper = Replace(arrLine(3), "$(dllselfregister)", "Regserver")
                sOper = Replace(sOper, "$(tlbregister)", "Regtypelib")
                sShared = Replace(LCase$(arrLine(4)), "$(shared)", "Sharedfile")
                sDest = Replace(sDest, "$(winsyspathsysfile)", "{sys}")
                sDest = Replace(sDest, "$(apppath)", "{app}")
                sDest = Replace(sDest, "$(winsyspath)", "{sys}")
                sDest = Replace(sDest, "$(winpath)", "{win}")
                sDest = Replace(sDest, "$(msdaopath)", "{dao}")
                ' check rmchart.dll and redemption.dll
                tmpFile = File_Token(sFile, "f", "\")
                Select Case LCase$(tmpFile)
                Case "rmchart.dll", "redemption.dll"
                    sOper = "Regserver"
                End Select
                lstLine(1) = StrSource & "\" & sFile
                lstLine(2) = sDest
                lstLine(3) = sOper
                lstLine(4) = sShared
                LstViewUpdate lstLine, lstFiles, vbNullString
                Select Case InStr(1, sFile, ".exe", vbTextCompare)
                Case 0
                Case Else
                    lstLine(1) = File_Token(sFile, "f")
                    Call LstViewUpdate(lstLine, lstViewRun, vbNullString)
                End Select
                Select Case InStr(1, sFile, ".msi", vbTextCompare)
                Case 0
                Case Else
                    lstLine(1) = File_Token(sFile, "f")
                    Call LstViewUpdate(lstLine, lstViewRun, vbNullString)
                End Select
            Case "titl"
                txtSetup(1).Text = sRest
                txtSetup(6).Text = "{sd}\" & sRest
                txtSetup(5).Text = sRest
                txtSetup(0).Text = sRest
            Case "grou"
                txtSetup(14).Text = sRest
            Case "appe"
                txtSetup(15).Text = sRest
                txtSetup(16).Text = "{app}\" & sRest
            End Select
        End Select
NextLine:
    Loop
    Close #intFile
    LstViewAutoResize lstFiles
    LstViewAutoResize lstViewRun
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Public Sub LstViewToFile(progBar As Object, lstView As Object, ByVal strFile As String, Optional Delim As String = vbNullString)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim lstStr() As String
    Dim retStr As String
    Dim intFile As Integer
    Dim colStr As String
    intFile = FreeFile
    Open strFile For Output Access Write As #intFile
    retStr = vbNullString
    If Len(Delim) = 0 Then Delim = Chr$(253)
    lstTot = lstView.ListItems.Count
    colStr = LstViewColNames(lstView)
    colStr = Replace(colStr, ",", Delim)
    Print #intFile, colStr
    For lstCnt = 1 To lstTot
        lstStr = LstViewGetRow(lstView, lstCnt)
        retStr = MvFromArray(lstStr, Delim)
        Print #intFile, retStr
        Err.Clear
    Next
    Close #intFile
    Err.Clear
End Sub
Public Sub ClearAll()
    On Error Resume Next
    Dim varControl As Control
    For Each varControl In Me.Controls
        If TypeName(varControl) = "TextBox" Then
            varControl.Text = vbNullString
        ElseIf TypeName(varControl) = "ComboBox" Then
            varControl.Clear
        ElseIf TypeName(varControl) = "PictureBox" Then
            Set varControl.Picture = Nothing
        ElseIf TypeName(varControl) = "CheckBox" Then
            varControl.Value = 0
        ElseIf TypeName(varControl) = "RichTextBox" Then
            varControl.Text = vbNullString
        ElseIf TypeName(varControl) = "ListView" Then
            varControl.ListItems.Clear
        ElseIf TypeName(varControl) = "ListBox" Then
            varControl.Clear
        ElseIf TypeName(varControl) = "OptionButton" Then
            varControl.Value = False
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
