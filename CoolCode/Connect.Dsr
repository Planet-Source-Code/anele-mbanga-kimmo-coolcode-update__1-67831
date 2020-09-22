VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11610
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   18645
   _ExtentX        =   32888
   _ExtentY        =   20479
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Kimmo - Cool Source"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private clsCoolCode As clsCoolCodeAddIn
'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo Err_Handler:
    'save the vb instance
    Set VBInst = Application
    Running = True
    'create the toolbar
    Set VbInstCB = VBInst.CommandBars.Add(App_Name, msoBarTop, , False)
    Set VbInstCB1 = VBInst.CommandBars.Add(App_Name & "1", msoBarTop, , False)
    VbInstCB.Visible = GetSetting(App_Name, "Preferences", "Visible", True)
    VbInstCB.Top = 120
    VbInstCB.Left = 0
    VbInstCB1.Visible = GetSetting(App_Name, "Preferences", "Visible", True)
    VbInstCB1.Top = 240
    VbInstCB1.Left = 0
    Set clsCoolCode = New clsCoolCodeAddIn
    clsCoolCode.Initialize VbInstCB
    clsCoolCode.Initialize1 VbInstCB1
    Err.Clear
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".AddinInstance_OnConnection"
    Resume Next
    Err.Clear
End Sub
'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error GoTo Err_Handler:
    Running = False
    SaveSetting App_Name, "Preferences", "Visible", VbInstCB.Visible
    Unload frmAddCode
    Unload frmPg
    Unload frmToolTips
    Unload frmVariables
    Unload frmLinesComments
    Unload frmUPX
    Unload frmConvert
    VbInstCB.Visible = False
    VbInstCB1.Visible = False
    clsCoolCode.Destroy
    Set clsCoolCode = Nothing
    'destroy toolbar variable
    VbInstCB.Delete
    VbInstCB1.Delete
    Set VbInstCB = Nothing
    Set VbInstCB1 = Nothing
    Set VBInst = Nothing
    Set frmAddCode = Nothing
    Set frmPg = Nothing
    Set frmToolTips = Nothing
    Set frmVariables = Nothing
    Set frmLinesComments = Nothing
    Set frmConvert = Nothing
    Set frmUPX = Nothing
    Err.Clear
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".AddinInstance_OnDisconnection"
    Resume Next
    Err.Clear
End Sub

Private Function Compiled() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then
        Compiled = False
    Else
        Compiled = True
    End If
End Function

