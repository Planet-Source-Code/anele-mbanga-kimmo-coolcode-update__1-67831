VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChangeFont 
   Caption         =   "Change Font"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstControls 
      Height          =   255
      Left            =   38340
      TabIndex        =   0
      Top             =   165
      Width           =   135
   End
   Begin MSComDlg.CommonDialog FontDB 
      Left            =   3720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChangeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mcmpCurrentForm As VBComponent      'current form
Dim mcolCtls        As VBControls       'form's controls
Dim controlarray() As VBControl              'link from list's listdata to the control name

