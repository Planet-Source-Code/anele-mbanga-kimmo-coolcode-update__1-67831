VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoolCodeAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mcolCtls As VBIDE.VBControls       'form's controls
Private formControls As Collection
Private ccUpdateRs As Office.CommandBarButton
Private ccWPF As Office.CommandBarButton
Private WithEvents ccUpdateRsEvents As CommandBarEvents
Attribute ccUpdateRsEvents.VB_VarHelpID = -1
Private WithEvents ccWPFEvents As CommandBarEvents
Attribute ccWPFEvents.VB_VarHelpID = -1
Private ccChangeFont As Office.CommandBarButton
Private ccCompileCompress As Office.CommandBarButton
Private ccMySQL As Office.CommandBarButton
Private ccMySQL1 As Office.CommandBarButton
Private ccSwap As Office.CommandBarButton
Private ccCommentBlock As Office.CommandBarButton
Private ccUnCommentBlock As Office.CommandBarButton
Private ccConvertToUpperCase As Office.CommandBarButton
Private ccConvertToLowerCase As Office.CommandBarButton
Private ccConvertToProperCase As Office.CommandBarButton
Private WithEvents ccMySQL1Events As CommandBarEvents
Attribute ccMySQL1Events.VB_VarHelpID = -1
Private WithEvents ccMySQLEvents As CommandBarEvents
Attribute ccMySQLEvents.VB_VarHelpID = -1
Private WithEvents ccSwapEvents As CommandBarEvents
Attribute ccSwapEvents.VB_VarHelpID = -1
Private WithEvents ccCommentBlockEvents As CommandBarEvents
Attribute ccCommentBlockEvents.VB_VarHelpID = -1
Private WithEvents ccUnCommentBlockEvents As CommandBarEvents
Attribute ccUnCommentBlockEvents.VB_VarHelpID = -1
Private WithEvents ccChangeFontEvents As CommandBarEvents
Attribute ccChangeFontEvents.VB_VarHelpID = -1
Private WithEvents ccFontNameEvents As CommandBarEvents
Attribute ccFontNameEvents.VB_VarHelpID = -1
Private WithEvents ccFontSizeEvents As CommandBarEvents
Attribute ccFontSizeEvents.VB_VarHelpID = -1
Private WithEvents ccConvertToUpperCaseEvents As CommandBarEvents
Attribute ccConvertToUpperCaseEvents.VB_VarHelpID = -1
Private WithEvents ccConvertToLowerCaseEvents As CommandBarEvents
Attribute ccConvertToLowerCaseEvents.VB_VarHelpID = -1
Private WithEvents ccConvertToProperCaseEvents As CommandBarEvents
Attribute ccConvertToProperCaseEvents.VB_VarHelpID = -1
Private WithEvents ccCompileCompressEvents As CommandBarEvents
Attribute ccCompileCompressEvents.VB_VarHelpID = -1
Private ccClearText As Office.CommandBarButton
Private WithEvents ccClearTextEvents As CommandBarEvents
Attribute ccClearTextEvents.VB_VarHelpID = -1
Private ccInno As Office.CommandBarButton
Private ccFontName As Office.CommandBarControl
Private ccFontSize As Office.CommandBarControl
Private ccToolTips As Office.CommandBarButton
Private ccExplore As Office.CommandBarButton
Private ccComment As Office.CommandBarButton
Private ccRemove As Office.CommandBarButton
Private ccProcedures As Office.CommandBarButton
Private ccVariables As Office.CommandBarButton
Private ccCommentQuote As Office.CommandBarButton
Private ccAddCode As Office.CommandBarButton
Private ccClose As Office.CommandBarButton
Private ccAbout As Office.CommandBarButton
Private ccLineStartWithComments As Office.CommandBarButton
Private ccIntellisense As Office.CommandBarButton
Private ccValidDateMask As Office.CommandBarButton
Private ccRemoveError As Office.CommandBarButton
Private ccBackUp As Office.CommandBarButton
Private ccProperCase As Office.CommandBarButton
Private ccTextBoxHilite As Office.CommandBarButton
Private ccRemoveBlanks As Office.CommandBarButton
Private ccSort As Office.CommandBarButton
Private ccRemoveLike As Office.CommandBarButton
Private ccTrimLines As Office.CommandBarButton
Private ccRemoveLineContinuation As Office.CommandBarButton
Private ccBreakDeclarations As Office.CommandBarButton
Private ccSpeedStringFunctions As Office.CommandBarButton
Private ccMoveComments As Office.CommandBarButton
Private ccSpeedNextLoopEnd As Office.CommandBarButton
Private ccInsertDoEvents As Office.CommandBarButton
Private ccPassByVal As Office.CommandBarButton
Private ccPassByRef As Office.CommandBarButton
Private ccInitializeOptionalVariables As Office.CommandBarButton
Private ccRemoveEmptyProcedures As Office.CommandBarButton
Private ccNormalizeIf As Office.CommandBarButton
Private ccBreakComplexForLoops As Office.CommandBarButton
Private ccMoveDeclarations As Office.CommandBarButton
Private ccInsertOnErrorResumeNext As Office.CommandBarButton
Private ccFormat As Office.CommandBarButton
Private WithEvents ccInnoEvents As CommandBarEvents
Attribute ccInnoEvents.VB_VarHelpID = -1
Private WithEvents ccCommentEvents As CommandBarEvents
Attribute ccCommentEvents.VB_VarHelpID = -1
Private WithEvents ccLineStartWithCommentsEvents As CommandBarEvents
Attribute ccLineStartWithCommentsEvents.VB_VarHelpID = -1
Private WithEvents ccExploreEvents As CommandBarEvents
Attribute ccExploreEvents.VB_VarHelpID = -1
Private WithEvents ccVariablesEvents As CommandBarEvents
Attribute ccVariablesEvents.VB_VarHelpID = -1
Private WithEvents ccAddCodeEvents As CommandBarEvents
Attribute ccAddCodeEvents.VB_VarHelpID = -1
Private WithEvents ccToolTipsEvents As CommandBarEvents
Attribute ccToolTipsEvents.VB_VarHelpID = -1
Private WithEvents ccCloseEvents As CommandBarEvents
Attribute ccCloseEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveLikeEvents As CommandBarEvents
Attribute ccRemoveLikeEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveBlanksEvents As CommandBarEvents
Attribute ccRemoveBlanksEvents.VB_VarHelpID = -1
Private WithEvents ccTrimLinesEvents As CommandBarEvents
Attribute ccTrimLinesEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveLineContinuationEvents As CommandBarEvents
Attribute ccRemoveLineContinuationEvents.VB_VarHelpID = -1
Private WithEvents ccBreakDeclarationsEvents As CommandBarEvents
Attribute ccBreakDeclarationsEvents.VB_VarHelpID = -1
Private WithEvents ccSpeedStringFunctionsEvents As CommandBarEvents
Attribute ccSpeedStringFunctionsEvents.VB_VarHelpID = -1
Private WithEvents ccMoveCommentsEvents As CommandBarEvents
Attribute ccMoveCommentsEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveEvents As CommandBarEvents
Attribute ccRemoveEvents.VB_VarHelpID = -1
Private WithEvents ccCommentQuoteEvents As CommandBarEvents
Attribute ccCommentQuoteEvents.VB_VarHelpID = -1
Private WithEvents ccSpeedNextLoopEndEvents As CommandBarEvents
Attribute ccSpeedNextLoopEndEvents.VB_VarHelpID = -1
Private WithEvents ccInsertDoEventsEvents As CommandBarEvents
Attribute ccInsertDoEventsEvents.VB_VarHelpID = -1
Private WithEvents ccPassByValEvents As CommandBarEvents
Attribute ccPassByValEvents.VB_VarHelpID = -1
Private WithEvents ccPassByRefEvents As CommandBarEvents
Attribute ccPassByRefEvents.VB_VarHelpID = -1
Private WithEvents ccInitializeOptionalVariablesEvents As CommandBarEvents
Attribute ccInitializeOptionalVariablesEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveEmptyProceduresEvents As CommandBarEvents
Attribute ccRemoveEmptyProceduresEvents.VB_VarHelpID = -1
Private WithEvents ccNormalizeIfEvents As CommandBarEvents
Attribute ccNormalizeIfEvents.VB_VarHelpID = -1
Private WithEvents ccProceduresEvents As CommandBarEvents
Attribute ccProceduresEvents.VB_VarHelpID = -1
Private WithEvents ccBreakComplexForLoopsEvents As CommandBarEvents
Attribute ccBreakComplexForLoopsEvents.VB_VarHelpID = -1
Private WithEvents ccMoveDeclarationsEvents As CommandBarEvents
Attribute ccMoveDeclarationsEvents.VB_VarHelpID = -1
Private WithEvents ccInsertOnErrorResumeNextEvents As CommandBarEvents
Attribute ccInsertOnErrorResumeNextEvents.VB_VarHelpID = -1
Private WithEvents ccFormatEvents As CommandBarEvents
Attribute ccFormatEvents.VB_VarHelpID = -1
Private WithEvents ccSortEvents As CommandBarEvents
Attribute ccSortEvents.VB_VarHelpID = -1
Private WithEvents ccIntellisenseEvents As CommandBarEvents
Attribute ccIntellisenseEvents.VB_VarHelpID = -1
Private WithEvents ccTextBoxHiliteEvents As CommandBarEvents
Attribute ccTextBoxHiliteEvents.VB_VarHelpID = -1
Private WithEvents ccProperCaseEvents As CommandBarEvents
Attribute ccProperCaseEvents.VB_VarHelpID = -1
Private WithEvents ccBackUpEvents As CommandBarEvents
Attribute ccBackUpEvents.VB_VarHelpID = -1
Private WithEvents ccRemoveErrorEvents As CommandBarEvents
Attribute ccRemoveErrorEvents.VB_VarHelpID = -1
Private WithEvents ccAboutEvents As CommandBarEvents
Attribute ccAboutEvents.VB_VarHelpID = -1
Private WithEvents ccValidDateMaskEvents As CommandBarEvents
Attribute ccValidDateMaskEvents.VB_VarHelpID = -1
Private Sub ccAboutEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    MsgBox "Kimmo - Cool Source " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "This tool is used to improve code development in Visual Basic." & vbCrLf & vbCrLf & "Among other things, one can..." & vbCrLf & vbCrLf & _
    "Insert intellisense code, insert highlight code, insert propercase case," & vbCrLf & "backup project, remove blank lines, remove lines containing particular text," & vbCrLf & _
    "remove line continuation, sort procedures, trim lines, break multiple declarations," & vbCrLf & "speed string functions, move comments to line below, speed next loop ending," & vbCrLf & _
    "insert DoEvents on loops, pass string variables by value or by reference," & vbCrLf & "initialize optional variables, remove empty procedures," & vbCrLf & _
    "break if statements, break complex for loops, move declarations to beginning of procedures," & vbCrLf & "insert on error resume next and remove on error resume next, format source code." & vbCrLf & vbCrLf & _
    "Copyrights - " & App.LegalCopyright, vbOKOnly + vbInformation + vbApplicationModal, "About Kimmo - Cool Source"
    Err.Clear
End Sub
Private Sub ccAddCodeEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim iCurrentLine As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim lText As String
    Dim VbACp As VBIDE.CodePane
    Dim cLoc As String
    Set VbACp = VBInst.ActiveCodePane
    VbACp.GetSelection iCurrentLine, b, c, d
    'get line the mouse is on
    lText = VbACp.CodeModule.lines(iCurrentLine, 1)
    cLoc = ReadReg(App_Name, "location", Trim$(lText), "b")
    With frmAddCode
        .ClearAll
        .txtLineContaining.Text = Trim$(lText)
        If cLoc = "b" Then
            .optBefore.Value = True
        Else
            .optAfter.Value = True
        End If
        .optCurrent.Value = 1
        .txtInsertion.Text = ReadReg(App_Name, "addcode", .txtLineContaining.Text)
        .Show
    End With
EndProc:
    Set VbACp = Nothing
    Err.Clear
End Sub
Private Sub ccBackUpEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    BackupFirst False
    Err.Clear
End Sub
Private Sub ccBreakComplexForLoopsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    'BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_FixComplexLoops VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_FixComplexLoops VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccBreakDeclarationsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_BreakMultiDeclarations VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_BreakMultiDeclarations VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_BreakMultiDeclarations VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccChangeFontEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim strFont As String
    Dim StrSize As String
'Dim strQ As String
    Dim intA As Long
    Dim ctl As Object
    
    strFont = ccFontName.Text
    StrSize = ccFontSize.Text
    
    If Len(strFont) = 0 Or Len(StrSize) = 0 Then
        MsgBox "Please select the font name and font size first.", vbOKOnly + vbApplicationModal, "Font Error"
        Exit Sub
    End If
    If Val(StrSize) < 7 Then StrSize = 7
    intA = MsgBox("You are about to change the font property for " & ReadFormControls & " controls to " & strFont & ", " & StrSize & ", are you sure?", vbYesNo + vbQuestion + vbApplicationModal, "Conform Change")
    If intA = vbNo Then Exit Sub
    Set VbCp = VBInst.SelectedVBComponent
    For Each ctl In VbCp.Designer.VBControls
        If IsPropertyAvailable(ctl.Properties!Font) = False Then GoTo SkipIt
        ctl.Properties("Font")!Bold = False
        ctl.Properties("Font")!Italic = False
        'ctl.Properties("Font")!Strikethrough = False
        'ctl.Properties("Font")!Underline = True
        ctl.Properties("Font")!Name = strFont
        ctl.Properties("Font")!Size = Val(StrSize)
        ctl.Refresh
        'ctl.Properties("Autosize") = True
SkipIt:
        Err.Clear
    Next
    Err.Clear
End Sub
Private Sub ccClearTextEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim ctr As VBIDE.VBControl
    For Each ctr In VBInst.SelectedVBComponent.Designer.VBControls
        If ctr.ProgId = "VB.TextBox" Then
            ctr.Properties("Text") = vbNullString
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Private Sub ccCloseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim tpaWindows() As VBIDE.vbext_WindowType
    ReDim tpaWindows(2)
    tpaWindows(1) = vbext_wt_CodeWindow '# Windows like this one
    tpaWindows(2) = vbext_wt_Designer '# Windows with forms on them
    CloseWindows tpaWindows()
    Err.Clear
End Sub
Private Sub ccCommentBlockEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    CommentSelectedText
    Err.Clear
End Sub
Private Sub ccCommentEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmComment.Show
    Err.Clear
End Sub
Private Sub ccCommentQuoteEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    allCode = ProjectCode
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    frmLinesComments.lstVariables.ListItems.Clear
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        frmLinesComments.Show
        frmLinesComments.lstVariables.ListItems.Clear
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_CheckStartWithComments VbCp, "'" & Quote
    Case 3
        frmLinesComments.Show
        frmLinesComments.lstVariables.ListItems.Clear
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_CheckStartWithComments VbCp, "'" & Quote
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    If frmLinesComments.lstVariables.ListItems.Count > 0 Then
        frmLinesComments.Show
    Else
        MsgBox "There are no lines that start with comments and quote in the code.", vbOKOnly + vbInformation + vbApplicationModal, "No Comments"
    End If
    Err.Clear
End Sub
Private Function ResourceFile_SaveItemToDisk(ByVal iResourceNum As Integer, ByVal sResourceType As String, ByVal sDestFileName As String) As Long
    On Error Resume Next
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    On Error GoTo SaveResItemToDisk_err
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    iFileNumOut = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNumOut
    Put #iFileNumOut, , bytResourceData
    Close #iFileNumOut
    ResourceFile_SaveItemToDisk = 0
    Err.Clear
    Exit Function
SaveResItemToDisk_err:
    ResourceFile_SaveItemToDisk = Err.Number
    Err.Clear
End Function
Private Sub ccCompileCompressEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim sExePath As String
    Dim CurrentProject As VBIDE.VBProject
    Dim upxFile As String
    On Error GoTo ErrHandler
    'get reference to the currently open project
    Set CurrentProject = VBInst.ActiveVBProject
    If CurrentProject Is Nothing Then
        'the project seems not to be loaded
        MsgBox "Please, open a VB project.", vbExclamation
    Else
        ' extract the upx compiler
        upxFile = File_Token(CurrentProject.Filename, "p") & "\upx.exe"
        If File_Exists(upxFile) = False Then
            ResourceFile_SaveItemToDisk 101, "CUSTOM", upxFile
            Do Until File_Exists(upxFile) = True
                DoEvents
            Loop
        End If
        With CurrentProject
            Screen.MousePointer = vbHourglass
            .MakeCompiledFile
            sExePath = .BuildFileName 'get the path to the executable
            installPath = File_Token(sExePath, "p") & "\Package\Support"
            Screen.MousePointer = vbDefault
        End With
        frmUPX.SetProject sExePath 'set the path to the executable in the main form
        frmUPX.Show
        Set CurrentProject = Nothing
    End If
ExitSub:
    Err.Clear
    Exit Sub
ErrHandler:
    AddInErr Err
    Resume ExitSub
    Err.Clear
End Sub
Private Sub ccConvertToLowerCaseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    ConvertSelectedText vbLowerCase
    Err.Clear
End Sub
Private Sub ccConvertToProperCaseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    ConvertSelectedText vbProperCase
    Err.Clear
End Sub
Private Sub ccConvertToUpperCaseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    ConvertSelectedText vbUpperCase
    Err.Clear
End Sub
Private Sub ccExploreEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    allCode = ProjectCode
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    frmProject.Show
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        frmProject.txtCode.Text = ""
        frmProject.treeProject.Nodes.Clear
        'VbCp.CodeModule.CodePane.Show
        Code_Explore VbCp
        'Code_FixGetRow VbCp
    Case 3
        frmProject.txtCode.Text = ""
        frmProject.treeProject.Nodes.Clear
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            'VbCp.CodeModule.CodePane.Show
            Code_Explore VbCp
            'Code_FixGetRow VbCp
            'frmProject.Refresh
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    'Code_CheckUsage
    'LstViewAutoResize frmProject.lstProject
    'frmProject.Refresh
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccInnoEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    SetupConverter.SetupFile = StringGetFileToken(VBInst.ActiveVBProject.Filename, "p") & "\Package\SETUP.LST"
    If boolFileExists(SetupConverter.SetupFile) = False Then
        Call MyPrompt("The setup file does not exist in the package folder, please compile your project first!", "o", "e", "Setup File Error")
        Err.Clear
        Exit Sub
    End If
    SetupConverter.cmdInno.Enabled = False
    SetupConverter.Show
    SetupConverter.ReadVbSetupFile
    Err.Clear
End Sub
Private Sub ccFormatEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_MoveComment VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_MoveComment VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_MoveComment VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccInitializeOptionalVariablesEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_InitializeOptionalVariables VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_InitializeOptionalVariables VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_InitializeOptionalVariables VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccInsertDoEventsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_InsertDoEvents VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_InsertDoEvents VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_InsertDoEvents VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccInsertOnErrorResumeNextEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False, True
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_InsertOnError VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_InsertOnError VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccIntellisenseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_Intellisense VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_Intellisense VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccLineStartWithCommentsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    allCode = ProjectCode
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    frmLinesComments.lstVariables.ListItems.Clear
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        frmLinesComments.Show
        frmLinesComments.lstVariables.ListItems.Clear
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_CheckStartWithComments VbCp
    Case 3
        frmLinesComments.Show
        frmLinesComments.lstVariables.ListItems.Clear
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_CheckStartWithComments VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    If frmLinesComments.lstVariables.ListItems.Count > 0 Then
        frmLinesComments.Show
    Else
        MsgBox "There are no lines that start with comments in the code.", vbOKOnly + vbInformation + vbApplicationModal, "No Comments"
    End If
    Err.Clear
End Sub
Private Sub ccMoveCommentsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_MoveComment VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_MoveComment VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_MoveComment VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccMoveDeclarationsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_MoveDeclarations VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_MoveDeclarations VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccMySQLEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    MySQLSelectedText
    Err.Clear
End Sub
Private Sub ccMySQL1Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    MySQLSelectedText1
    Err.Clear
End Sub
Private Sub ccNormalizeIfEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_NormalizeIf VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_NormalizeIf VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccPassByRefEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_PassBy VbCp, "ByRef", procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_PassBy VbCp, "ByRef"
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_PassBy VbCp, "ByRef"
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccPassByValEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_PassBy VbCp, , procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_PassBy VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_PassBy VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccProceduresEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim cDead As Long
    ''BackupFirst
    aAll = WhatToProcess
    frmVariables.lstVariables.ListItems.Clear
    frmVariables.Show
    frmVariables.Caption = "Procedure Declarations"
    frmVariables.lstVariables.ColumnHeaders(3).Text = "Start Location"
    frmVariables.lstVariables.ColumnHeaders(4).Text = "End Location"
    cDead = 0
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        cDead = Code_CheckProcedures(VbCp)
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        cDead = 0
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            cDead = cDead + Code_CheckProcedures(VbCp)
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    If frmVariables.lstVariables.ListItems.Count > 0 Then
        frmVariables.Show
        frmVariables.cmdRemove.Enabled = IIf((cDead = 0), False, True)
    End If
    MsgBox cDead & " dead procedures have been found in your project / selected component!", vbOKOnly + vbExclamation + vbApplicationModal, "Dead Procedures"
    Err.Clear
End Sub
Private Sub ccProperCaseEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_Intellisense VbCp, "p"
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_Intellisense VbCp, "p"
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveBlanksEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_RemoveBlankLines VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_RemoveBlankLines VbCp, 1, -1
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_RemoveBlankLines VbCp, 1, -1
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveEmptyProceduresEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_RemoveEmpty VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_RemoveEmpty VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveErrorEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_RemoveLines VbCp, True, 1, -1, "Err.Clear", "On Error Resume Next", "'Err.Clear", "'On Error Resume Next"
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_RemoveLines VbCp, True, 1, -1, "Err.Clear", "On Error Resume Next", "'Err.Clear", "'On Error Resume Next"
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    allCode = ProjectCode
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_RemoveCommentedDeclarations VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_RemoveCommentedDeclarations VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveLikeEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim xSel As String
    Dim m As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    'Dim z As String
    VBInst.ActiveCodePane.GetSelection m, n, X, Y
    xSel = VBInst.ActiveCodePane.CodeModule.lines(m, 1)
    procStart = Procedure_Start
    procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Dim StrLine As String
    Dim strFalse As String
    Dim strRemove As String
    xSel = Trim$(xSel) & ",True"
ReEnter:
    StrLine = InputBox("Please type in the string to remove below, a comma and then a match keyword:" & vbCr & vbCr & "These could be - " & vbCr & vbCr & "DoEvents,True" & vbCr & ",True" & vbCr & ",False" & vbCr & vbCr & "Where True means the line should be an exact match and False signifies that the line should contain the text.", "Remove Lines Like", xSel)
    StrLine = Trim$(StrLine)
    If Len(StrLine) = 0 Then
        Err.Clear
        Exit Sub
    End If
    If InStr(1, StrLine, ",") = 0 Then
        aAll = MsgBox("The control comma has not been found!", vbRetryCancel + vbExclamation + vbApplicationModal, "Control Character Error")
        Select Case aAll
        Case vbRetry
            GoTo ReEnter
        Case Else
            Err.Clear
            Exit Sub
        End Select
    End If
    strFalse = Split(StrLine, ",")(1)
    strRemove = Split(StrLine, ",")(0)
    strFalse = LCase$(Trim$(strFalse))
    strRemove = Trim$(strRemove)
    Select Case strFalse
    Case "true", "false"
    Case Else
        aAll = MsgBox("The control word(s) False/True has not been found after the comma!", vbRetryCancel + vbExclamation + vbApplicationModal, "Control Word False/True Error")
        Select Case aAll
        Case vbRetry
            GoTo ReEnter
        Case Else
            Err.Clear
            Exit Sub
        End Select
    End Select
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        If strFalse = "true" Then
            Code_RemoveLines VbCp, True, 1, -1, strRemove
        Else
            Code_RemoveLines VbCp, False, 1, -1, strRemove
        End If
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            If strFalse = "true" Then
                Code_RemoveLines VbCp, True, 1, -1, strRemove
            Else
                Code_RemoveLines VbCp, False, 1, -1, strRemove
            End If
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccRemoveLineContinuationEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_RemoveLineContinuation VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_RemoveLineContinuation VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_RemoveLineContinuation VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccSortEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_SortProcedures VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_SortProcedures VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccSpeedNextLoopEndEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_SpeedNextLoopEnd VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_SpeedNextLoopEnd VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_SpeedNextLoopEnd VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccSpeedStringFunctionsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_ConvertStringFunctions VbCp, , procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_ConvertStringFunctions VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_ConvertStringFunctions VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccSwapEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    SwapSelectedText
    Err.Clear
End Sub
Private Sub ccTextBoxHiliteEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_Intellisense VbCp, "h"
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_Intellisense VbCp, "h"
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccToolTipsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    With frmToolTips
        .ClearAll
        .Show
        .UpdateAll
    End With
    Err.Clear
End Sub
Private Sub ccTrimLinesEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start
    procEnd = Procedure_End
    sWhatToProcess
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    ''BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 1
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_TrimLines VbCp, procStart, procEnd
        Code_Format VbCp, procStart, procEnd
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        Code_TrimLines VbCp
        Code_Format VbCp
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_TrimLines VbCp
            Code_Format VbCp
NextModule:
            DoEvents
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub Code_BreakMultiDeclarations(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim newL As String
    Dim spLine() As String
    Dim spCnt As Integer
    Dim spTot As Integer
    Dim AsPos As Long
    Dim varType As String
    Dim newDecl As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = tLines To StartLine Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Break Multi Declarations") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        newL = vbNullString
        If Left$(currLine, 4) = "Dim " Then
            If InStr(1, currLine, ", ", vbTextCompare) > 0 Then
                currLine = Mid$(currLine, 5)
                AsPos = InStr(1, currLine, " As ", vbTextCompare)
                If AsPos > 0 Then
                    varType = Trim$(Mid$(currLine, AsPos + 4))
                Else
                    varType = vbNullString
                End If
                spLine = Split(currLine, ", ")
                spTot = UBound(spLine)
                For spCnt = 0 To spTot
                    newDecl = "Dim " & Trim$(spLine(spCnt))
                    newDecl = FixDeclaration(newDecl)
                    If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                        newDecl = newDecl & " As " & varType
                    End If
                    newL = newL & newDecl & vbNewLine
                    Err.Clear
                Next
                VbCp.CodeModule.ReplaceLine tCount, newL
            Else
                newDecl = FixDeclaration(currLine)
                If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                    newDecl = newDecl & " As Variant"
                End If
                VbCp.CodeModule.ReplaceLine tCount, newDecl
            End If
        ElseIf Left$(currLine, 6) = "Const " Then
            If InStr(1, currLine, ", ", vbTextCompare) > 0 Then
                currLine = Mid$(currLine, 7)
                AsPos = InStr(1, currLine, " As ", vbTextCompare)
                If AsPos > 0 Then
                    varType = Trim$(Mid$(currLine, AsPos + 6))
                Else
                    varType = vbNullString
                End If
                spLine = Split(currLine, ", ")
                spTot = UBound(spLine)
                For spCnt = 0 To spTot
                    newDecl = "Const " & Trim$(spLine(spCnt))
                    newDecl = FixDeclaration(newDecl)
                    If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                        newDecl = newDecl & " As " & varType
                    End If
                    newL = newL & newDecl & vbNewLine
                    Err.Clear
                Next
                VbCp.CodeModule.ReplaceLine tCount, newL
            Else
                newDecl = FixDeclaration(currLine)
                If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                    newDecl = newDecl & " As Variant"
                End If
                VbCp.CodeModule.ReplaceLine tCount, newDecl
            End If
        ElseIf Left$(currLine, 7) = "Static " Then
            If InStr(1, currLine, ", ", vbTextCompare) > 0 Then
                currLine = Mid$(currLine, 8)
                AsPos = InStr(1, currLine, " As ", vbTextCompare)
                If AsPos > 0 Then
                    varType = Trim$(Mid$(currLine, AsPos + 7))
                Else
                    varType = vbNullString
                End If
                spLine = Split(currLine, ", ")
                spTot = UBound(spLine)
                For spCnt = 0 To spTot
                    newDecl = "Static " & Trim$(spLine(spCnt))
                    newDecl = FixDeclaration(newDecl)
                    If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                        newDecl = newDecl & " As " & varType
                    End If
                    newL = newL & newDecl & vbNewLine
                    Err.Clear
                Next
                VbCp.CodeModule.ReplaceLine tCount, newL
            Else
                newDecl = FixDeclaration(currLine)
                If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                    newDecl = newDecl & " As Variant"
                End If
                VbCp.CodeModule.ReplaceLine tCount, newDecl
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_ConvertStringFunctions(VbCp As VBIDE.VBComponent, Optional bImprove As Boolean = True, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Speeding String Functions") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        If bImprove = True Then
            currLine = ImprovedStringFunction(currLine)
        Else
            currLine = UnImprovedStringFunction(currLine)
        End If
        VbCp.CodeModule.ReplaceLine tCount, currLine
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Function Code_DimStatements(ByVal StrValue As String) As String
    On Error Resume Next
    Dim spLines() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim dimStart As Long
    Dim newDim As New Collection
    dimStart = 0
    spLines = Split(StrValue, vbNewLine)
    spTot = UBound(spLines)
    For spCnt = 0 To spTot
        spStr = Trim$(spLines(spCnt))
        If Left$(spStr, 4) = "Dim " Or Left$(spStr, 6) = "Const " Or Left$(spStr, 7) = "Static " Then
            spStr = FixDeclaration(spStr)
            newDim.Add spStr, spStr
            spLines(spCnt) = vbNullString
            If dimStart = 0 Then
                dimStart = spCnt
            End If
        End If
        Err.Clear
    Next
    If newDim.Count > 0 Then
        spStr = MvFromCollection(newDim, vbNewLine)
    End If
    If dimStart > 0 Then
        spLines(dimStart) = spStr
    End If
    Code_DimStatements = MvFromArray(spLines, vbNewLine)
    Err.Clear
End Function
Private Sub Code_FixComplexLoops(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    Dim procStart As Long
    Dim procEnd As Long
    Dim varNames As New Collection
    Dim StrLine As String
    Dim procValues As String
    Dim tLines As Long
    Dim cLines As Long
    Dim currLine As String
    Dim newLines As String
    Dim SourceCode As String
    Dim spLines() As String
    Dim Db As DAO.Database
    Dim tb As DAO.Recordset
    Dim olds As String
    Dim news As String
    Dbase = ExactPath(App.Path) & "\CoolCode.mdb"
    Dao_CreateDatabase Dbase, True
    Do Until Dao_TableExists(Dbase, "Complex") = True
        Dao_CreateTable Dbase, "Complex", "File,Source,Improved", "te,me,me", "255,,", "1"
    Loop
    Set Db = DAO.OpenDatabase(Dbase)
    Set tb = Db.OpenRecordset("Complex")
    tLines = VbCp.CodeModule.CountOfLines
    SourceCode = VbCp.CodeModule.lines(1, tLines)
    spLines = Split(SourceCode, vbNewLine)
    tLines = UBound(spLines)
    Call InitProgress(tLines)
    For cLines = 0 To tLines
        If UpdateProgress(cLines, tLines, VbCp.Name & ".Recoding Complex Loops") = False Then
            Exit For
        End If
        currLine = Trim$(spLines(cLines))
        If IsMethod(currLine) = True Then
            procValues = currLine & vbNewLine
            procStart = cLines
        ElseIf IsEndOfMethod(currLine) = True Then
            procValues = procValues & currLine
            procEnd = cLines
            If IsSourceComplex(procValues) = True Then
                newLines = RecodeComplex(procValues, varNames)
                newLines = Code_DimStatements(newLines)
                olds = Code_RemoveBlanks(procValues)
                news = Code_RemoveBlanks(newLines)
                If LCase$(olds) <> LCase$(news) Then
                    tb.AddNew
                    tb!File = VbCp.CodeModule.Parent.FileNames(1)
                    tb!Source = procValues
                    tb!improved = newLines
                    tb.Update
                End If
            End If
            Set varNames = New Collection
        ElseIf IsVariable(currLine) = True Then
            procValues = procValues & currLine & vbNewLine
            currLine = FixDeclaration(currLine)
            StrLine = VariableName(currLine) & "," & VariableType(currLine)
            varNames.Add StrLine
        Else
            procValues = procValues & currLine & vbNewLine
        End If
        Err.Clear
    Next
    CloseProgress
    If tb.RecordCount > 0 Then
        Code_TrimLines VbCp
        currLine = VbCp.CodeModule.lines(1, VbCp.CodeModule.CountOfLines)
        tLines = tb.RecordCount
        tb.MoveFirst
        Call InitProgress(tLines)
        For cLines = 1 To tLines
            If UpdateProgress(cLines, tLines, VbCp.Name & ".Recoding Complex Loops") = False Then
                Exit For
            End If
            olds = tb!Source & vbNullString
            news = tb!improved & vbNullString
            currLine = Replace$(currLine, olds, news, , , vbTextCompare)
            tb.MoveNext
            Err.Clear
        Next
        CloseProgress
        VbCp.CodeModule.DeleteLines 1, VbCp.CodeModule.CountOfLines
        VbCp.CodeModule.AddFromString currLine
    End If
    tb.Close
    Db.Close
    Set tb = Nothing
    Set Db = Nothing
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_Format(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim CurrTab As Integer
    Dim CurrCommand As String
    Dim currLine As String
    Dim KydStart As Integer
    Dim TabAfter As Boolean
    Dim TabSpace As String
    Dim Count2 As Integer
    Dim LastKeyWord As String
    Dim tCount As Long
    Dim tLines As Long
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Formatting Code") = False Then
            Exit For
        End If
        currLine = LCase$(Trim$(VbCp.CodeModule.lines(tCount, 1)))
        KydStart = InStr(1, currLine, " ")
        If KydStart = 0 Then
            KydStart = Len(currLine)
        End If
        CurrCommand = Trim$(Mid$(currLine, 1, KydStart))
        Select Case CurrCommand
        Case "declare", "function", "sub", "type", "enum", "property", "withevents", "event"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            VbCp.CodeModule.ReplaceLine tCount, "Public " & VbCp.CodeModule.lines(tCount, 1)
            currLine = LCase$(Trim$(VbCp.CodeModule.lines(tCount, 1)))
            KydStart = InStr(1, currLine, " ")
            If KydStart = 0 Then
                KydStart = Len(currLine)
            End If
            CurrCommand = Trim$(Mid$(currLine, 1, KydStart))
        End Select
        Select Case CurrCommand
        Case "public", "private", "friend"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If Mid$(Trim$(Mid$(currLine, Len(CurrCommand) + 1)), 1, Len("sub")) = "sub" Or Mid$(Trim$(Mid$(currLine, Len(CurrCommand) + 1)), 1, Len("function")) = "function" Or Mid$(Trim$(Mid$(currLine, Len(CurrCommand) + 1)), 1, Len("type")) = "type" Or Mid$(Trim$(Mid$(currLine, Len(CurrCommand) + 1)), 1, Len("property")) = "property" Or Mid$(Trim$(Mid$(currLine, Len(CurrCommand) + 1)), 1, Len("enum")) = "enum" Then
                CurrTab = CurrTab + 1
                TabAfter = True
            End If
        Case "if"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If Len(currLine) > 6 Then
                If Mid$(currLine, Len(currLine) - 3, 4) <> "then" Then
                    If Right$(currLine, 1) = "_" Then
                        CurrTab = CurrTab + 1
                        TabAfter = True
                    End If
                Else
                    CurrTab = CurrTab + 1
                    TabAfter = True
                End If
            End If
        Case "while", "do", "for", "sub", "function", "type", "enum", "property", "with"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            CurrTab = CurrTab + 1
            TabAfter = True
        Case "select"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If Mid$(currLine, 1, Len("select case")) = "select case" Then
                TabAfter = False
            End If
        Case "end", "wend", "loop", "next", "exit"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If Mid$(currLine, 1, Len("end if")) = "end if" Or Mid$(currLine, 1, Len("next")) = "next" Or Mid$(currLine, 1, Len("wend")) = "wend" Or Mid$(currLine, 1, Len("loop")) = "loop" Or Mid$(currLine, 1, Len("end select")) = "end select" Or Mid$(currLine, 1, Len("end with")) = "end with" Then
                CurrTab = CurrTab - 1
                TabAfter = False
            End If
            If Mid$(currLine, 1, Len("end function")) = "end function" Or Mid$(currLine, 1, Len("end sub")) = "end sub" Or Mid$(currLine, 1, Len("end type")) = "end type" Or Mid$(currLine, 1, Len("end property")) = "end property" Or Mid$(currLine, 1, Len("end enum")) = "end enum" Then
                CurrTab = 0
                TabAfter = False
            End If
        Case "else", "case", "elseif"
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If LastKeyWord <> "select" Then
                CurrTab = CurrTab - 1
                TabSpace = vbNullString
                If Len(TabSpace) / 4 <> CurrTab Then
                    For Count2 = 1 To CurrTab
                        TabSpace = TabSpace & vbTab
                        Err.Clear
                    Next
                End If
            End If
            CurrTab = CurrTab + 1
            TabAfter = True
        End Select
        If TabAfter = True Then
            VbCp.CodeModule.ReplaceLine tCount, TabSpace & Trim$(VbCp.CodeModule.lines(tCount, 1))
            TabSpace = vbNullString
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                    Err.Clear
                Next
            End If
        Else
            TabSpace = vbNullString
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                    Err.Clear
                Next
            End If
            VbCp.CodeModule.ReplaceLine tCount, TabSpace & Trim$(VbCp.CodeModule.lines(tCount, 1))
        End If
        If Len(CurrCommand) > 1 Then
            If Mid$(CurrCommand, Len(CurrCommand) - 1, 1) = ":" Then
                VbCp.CodeModule.ReplaceLine tCount, Trim$(VbCp.CodeModule.lines(tCount, 1))
            End If
        End If
        LastKeyWord = CurrCommand
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_InitializeOptionalVariables(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tCount As Long
    Dim tLines As Long
    Dim currLine As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Initializing Optional Variables") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        If IsMethod(currLine) = True Then
            currLine = OptionalizedProcedure(currLine)
            VbCp.CodeModule.ReplaceLine tCount, currLine
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_FixGetRow(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tCount As Long
    Dim tLines As Long
    Dim currLine As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Fixing Get Row Code") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        If InStr(1, currLine, "LstViewGetRow", vbTextCompare) > 0 Then
            currLine = LstViewGetRowCode(currLine)
            VbCp.CodeModule.ReplaceLine tCount, currLine
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_InsertDoEvents(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim NextLine As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = tLines To StartLine Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Inserting DoEvents On Loops") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        NextLine = Trim$(VbCp.CodeModule.lines(tCount + 1, 1))
        If Left$(currLine, Len("While ")) = "While " Then
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                VbCp.CodeModule.ReplaceLine tCount, currLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(currLine, Len("For ")) = "For " Then
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                VbCp.CodeModule.ReplaceLine tCount, currLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(currLine, Len("Do ")) = "Do " Then
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                VbCp.CodeModule.ReplaceLine tCount, currLine & vbNewLine & "DoEvents"
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_InsertOnError(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    If clearError = 1 Then
        Code_RemoveLines VbCp, True, 1, -1, "Err.Clear", "On Error Resume Next", "'Err.Clear", "'On Error Resume Next"
        Code_RemoveBlankLines VbCp, 1, -1
    End If
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim NextLine As String
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For tCount = tLines To 1 Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Inserting Err.Clear") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        If IsMethodEnd(currLine) = True Then
            VbCp.CodeModule.CodePane.TopLine = tCount
            VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
            VbCp.CodeModule.InsertLines tCount, vbTab & "Err.Clear"
        End If
        Err.Clear
    Next
    CloseProgress
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For tCount = tLines To 1 Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Inserting On Error Resume Next") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        NextLine = Trim$(VbCp.CodeModule.lines(tCount + 1, 1))
        If Left$(NextLine, Len("On Error ")) <> "On Error " Then
            If IsMethod(currLine) = True Then
                VbCp.CodeModule.CodePane.TopLine = tCount
                VbCp.CodeModule.CodePane.SetSelection tCount, 1, tCount, -1
                If Right$(currLine, 1) <> "_" Then
                    VbCp.CodeModule.InsertLines tCount + 1, vbTab & "On Error Resume Next"
                End If
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_Intellisense(VbCp As VBIDE.VBComponent, Optional ByVal IH As String = "i")
    On Error Resume Next
    Dim myFileName As String
    Dim myFileData() As String
    Dim myFileTot As Long
    Dim myFileCnt As Long
    Dim myFileStr As String
    Dim myTextBox As String
    Dim myEndFnd As Boolean
    Dim myIndex As String
    Dim myCol As New Collection
    Dim fCnt As Integer
    Dim fTot As Integer
    Dim myCode As String
    Dim myFileD As String
    Dim fLine As String
    myFileName = VbCp.CodeModule.Parent.FileNames(1)
    myFileD = FileData(myFileName)
    Call StringParse(myFileData, myFileD, vbNewLine)
    myFileTot = UBound(myFileData)
    Call InitProgress(myFileTot)
    For myFileCnt = 1 To myFileTot
        If UpdateProgress(myFileCnt, myFileTot, VbCp.Name & ".Inserting Intellisense") = False Then
            Exit For
        End If
        myFileStr = Trim$(myFileData(myFileCnt))
        If Left$(myFileStr, 6) = "Begin " Then
            myTextBox = vbNullString
            myIndex = vbNullString
            Select Case LCase$(IH)
            Case "v"
                If Split(myFileStr, " ")(1) = "MSMask.MaskEdBox" Then
                    myTextBox = Split(myFileStr, " ")(2)
                    myEndFnd = False
                End If
            Case "i", "h", "p"
                If Split(myFileStr, " ")(1) = "VB.TextBox" Then
                    myTextBox = Split(myFileStr, " ")(2)
                    myEndFnd = False
                End If
            End Select
        ElseIf Left$(myFileStr, 6) = "Index " Then
            myIndex = Split(myFileStr, "=")(1)
            myIndex = Trim$(myIndex)
        ElseIf myFileStr = "End" Then
            Select Case LCase$(IH)
            Case "i", "h", "p", "v"
                If Len(myTextBox) > 0 Then
                    If Len(myIndex) > 0 Then
                        myTextBox = myTextBox & ",True"
                    Else
                        myTextBox = myTextBox & ",False"
                    End If
                    myCol.Add myTextBox, myTextBox
                End If
            End Select
            myTextBox = vbNullString
            myIndex = vbNullString
        End If
        Err.Clear
    Next
    fTot = myCol.Count
    For fCnt = 1 To fTot
        Select Case LCase$(IH)
        Case "i"
            myCode = IntellisenseCode(myCol.Item(fCnt))
        Case "h"
            myCode = TextBoxHiliteCode(myCol.Item(fCnt))
        Case "p"
            myCode = ProperCaseCode(myCol.Item(fCnt))
        Case "v"
            myCode = TextBoxHiliteCode(myCol.Item(fCnt))
            myCode = Trim$(myCode)
            fLine = MvField(myCode, 1, vbNewLine)
            If InStr(1, myFileD, fLine, vbTextCompare) = 0 Then
                VbCp.CodeModule.InsertLines VbCp.CodeModule.CountOfLines + 1, myCode
            End If
            myCode = ValidateDateMaskCode(myCol.Item(fCnt))
            myCode = Trim$(myCode)
            fLine = MvField(myCode, 1, vbNewLine)
            If InStr(1, myFileD, fLine, vbTextCompare) = 0 Then
                VbCp.CodeModule.InsertLines VbCp.CodeModule.CountOfLines + 1, myCode
            End If
        End Select
        Select Case LCase$(IH)
        Case "v"
        Case Else
            myCode = Trim$(myCode)
            fLine = MvField(myCode, 1, vbNewLine)
            If InStr(1, myFileD, fLine, vbTextCompare) = 0 Then
                VbCp.CodeModule.InsertLines VbCp.CodeModule.CountOfLines + 1, myCode
            End If
        End Select
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_MoveComment(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim ComPos As Long
    Dim comStr As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = tLines To StartLine Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Moving Comments") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        ' if this line is a comment, then ignore it
        If Left$(currLine, 1) = "'" Then
        Else
            ' find out if this line contains a comment and is not a comment
            ComPos = InStrRev(currLine, "'")
            If ComPos > 1 Then
                ' if there is a space before the comment
                If Mid$(currLine, ComPos - 1) = " " Then
                    ' the contents of the comment
                    comStr = Mid$(currLine, ComPos)               ' the contents before the comment
                    currLine = Left$(currLine, ComPos - 1)
                    VbCp.CodeModule.ReplaceLine tCount, currLine & vbNewLine & comStr
                End If
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_MoveDeclarations(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    Dim procStart As Long
    Dim procEnd As Long
    Dim procValues As String
    Dim tLines As Long
    Dim cLines As Long
    Dim currLine As String
    Dim newLines As String
    Dim SourceCode As String
    Dim spLines() As String
    Dim Db As DAO.Database
    Dim tb As DAO.Recordset
    Dim olds As String
    Dim news As String
    Dbase = ExactPath(App.Path) & "\CoolCode.mdb"
    Dao_CreateDatabase Dbase, True
    Do Until Dao_TableExists(Dbase, "Complex") = True
        Dao_CreateTable Dbase, "Complex", "File,Source,Improved", "te,me,me", "255,,", "1"
    Loop
    Set Db = DAO.OpenDatabase(Dbase)
    Set tb = Db.OpenRecordset("Complex")
    tLines = VbCp.CodeModule.CountOfLines
    SourceCode = VbCp.CodeModule.lines(1, tLines)
    Call StringParse(spLines, SourceCode, vbNewLine)
    tLines = UBound(spLines)
    Call InitProgress(tLines)
    For cLines = 1 To tLines
        If UpdateProgress(cLines, tLines, VbCp.Name & ".Moving Declarations") = False Then
            Exit For
        End If
        currLine = Trim$(spLines(cLines))
        If IsMethod(currLine) = True Then
            procValues = currLine & vbNewLine
            procStart = cLines
        ElseIf IsEndOfMethod(currLine) = True Then
            procValues = procValues & currLine
            procEnd = cLines
            newLines = Code_DimStatements(procValues)
            olds = Code_RemoveBlanks(procValues)
            news = Code_RemoveBlanks(newLines)
            If LCase$(olds) <> LCase$(news) Then
                tb.AddNew
                tb!File = VbCp.CodeModule.Parent.FileNames(1)
                tb!Source = procValues
                tb!improved = newLines
                tb.Update
            End If
        Else
            procValues = procValues & currLine & vbNewLine
        End If
        Err.Clear
    Next
    CloseProgress
    If tb.RecordCount > 0 Then
        Code_TrimLines VbCp
        currLine = VbCp.CodeModule.lines(1, VbCp.CodeModule.CountOfLines)
        tLines = tb.RecordCount
        tb.MoveFirst
        Call InitProgress(tLines)
        For cLines = 1 To tLines
            If UpdateProgress(cLines, tLines, VbCp.Name & ".Moving Declarations") = False Then
                Exit For
            End If
            olds = tb!Source & vbNullString
            news = tb!improved & vbNullString
            currLine = Replace$(currLine, olds, news, , , vbTextCompare)
            tb.MoveNext
            Err.Clear
        Next
        CloseProgress
        VbCp.CodeModule.DeleteLines 1, VbCp.CodeModule.CountOfLines
        VbCp.CodeModule.AddFromString currLine
    End If
    tb.Close
    Db.Close
    Set tb = Nothing
    Set Db = Nothing
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_NormalizeIf(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim IfPos As Long
    Dim ThenPos As Long
    Dim ElsePos As Long
    Dim ComPos As Long
    Dim spIf() As String
    Dim useLine As String
    Dim comStr As String
    Dim midStr As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Breaking If Statements") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        IfPos = InStr(1, currLine, "If ", vbTextCompare)
        ThenPos = InStr(1, currLine, " Then", vbTextCompare)
        ElsePos = InStr(1, currLine, " Else ", vbTextCompare)
        ComPos = InStrRev(currLine, "'", , vbTextCompare)
        comStr = vbNullString
        If IfPos = 1 Then
            If Right$(currLine, 1) = "_" Then
                GoTo NextLine
            End If
            If ComPos > 0 Then
                ' we have a comment
                If ComPos > ThenPos Then
                    If Mid$(currLine, ComPos - 1, 1) = Quote Then
                        If Mid$(currLine, ComPos + 1, 1) = Quote Then
                            ComPos = 0
                        Else
                            comStr = Mid$(currLine, ComPos)
                            currLine = Left$(currLine, ComPos - 1)
                        End If
                    End If
                End If
            End If
            If ElsePos > 0 Then
                currLine = currLine & vbNewLine & "End If"
                currLine = Replace$(currLine, "Else", vbNewLine & "Else" & vbNewLine)
            End If
            spIf = Split(currLine, vbNewLine)
            useLine = Trim$(spIf(0))
            ThenPos = InStr(useLine, " Then")
            If Right$(useLine, 4) <> "Then" Then
                midStr = Mid$(useLine, ThenPos + 5)
                useLine = Left$(useLine, ThenPos + 5) & vbNewLine & midStr
                If ElsePos = 0 Then
                    useLine = useLine & vbNewLine & "End If"
                End If
                spIf(0) = useLine
                currLine = Join(spIf, vbNewLine)
            End If
            If ComPos > 0 Then
                currLine = Replace$(currLine, " Then", " Then" & vbNewLine & comStr)
            End If
            VbCp.CodeModule.ReplaceLine tCount, currLine
        End If
NextLine:
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_PassBy(VbCp As VBIDE.VBComponent, Optional ByVal PassBy As String = "ByVal", Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim currLine As String
    Dim tLines As Long
    Dim tCount As Long
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Passing Strings and Variants " & PassBy) = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        If IsMethod(currLine) = True Then
            currLine = StringNewProcedure(currLine, PassBy)
            VbCp.CodeModule.ReplaceLine tCount, currLine
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Function Code_RemoveBlanks(ByVal StrValue As String) As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    Dim xLine As String
    xRslt = vbNullString
    xData = Split(StrValue, vbNewLine)
    xTot = UBound(xData)
    For xCnt = 0 To xTot
        xLine = Trim$(xData(xCnt))
        If Len(xLine) <> 0 Then
            xRslt = xRslt & xLine & vbNewLine
        End If
        Err.Clear
    Next
    Code_RemoveBlanks = StringRemoveDelim(xRslt, vbNewLine)
    Erase xData
    Err.Clear
End Function
Private Sub Code_RemoveEmpty(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    If clearError = 1 Then
        Code_RemoveLines VbCp, True, 1, -1, "'"
        Code_RemoveBlankLines VbCp, 1, -1
    End If
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim NextLine As String
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For tCount = tLines To 1 Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Removing Empty Procedures") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        NextLine = Trim$(VbCp.CodeModule.lines(tCount + 1, 1))
        If IsMethod(currLine) = True Then
            If IsMethodEnd(NextLine) = True Then
                VbCp.CodeModule.ReplaceLine tCount, " "
                VbCp.CodeModule.ReplaceLine tCount + 1, " "
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Code_RemoveBlankLines VbCp, 1, -1
    Err.Clear
End Sub
Private Sub Code_RemoveLineContinuation(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    Dim NextLine As String
    Dim nLen As Long
    tLines = VbCp.CodeModule.CountOfLines - 1
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Remove Line Continuation") = False Then
            Exit For
        End If
foundone:
        currLine = VbCp.CodeModule.lines(tCount, 1)
        NextLine = VbCp.CodeModule.lines(tCount + 1, 1)
        nLen = Len(currLine) + Len(NextLine)
        If Right$(currLine, 1) = "_" Then
            If Len(currLine) + Len(NextLine) < 255 Then
                VbCp.CodeModule.ReplaceLine tCount, Trim$(Left$(currLine, Len(currLine) - 1)) & " " & NextLine
                VbCp.CodeModule.DeleteLines tCount + 1
                GoTo foundone
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_RemoveLines(VbCp As VBIDE.VBComponent, ExactMatch As Boolean, StartLine As Long, EndLine As Long, ParamArray TextToDelete())
    On Error Resume Next
    Dim totLines As Long
    'Dim cntLines As Long
    Dim curLine As String
    Dim StrSearch As String
    'Dim varText As Variant
    'Dim compCode As String
    Dim lngLine As Long
    Dim bFound As Boolean
    Dim thisLine As Long
    Dim totSearchU As Long
    Dim cntSearchL As Long
    totSearchU = UBound(TextToDelete)
    cntSearchL = LBound(TextToDelete)
    Call InitProgress(totSearchU)
    For thisLine = cntSearchL To totSearchU
        DoEvents
        StrSearch = CStr(TextToDelete(thisLine))
        If UpdateProgress(thisLine + 1, totSearchU + 1, VbCp.Name & "." & StrSearch) = False Then
            Exit For
        End If
        lngLine = StartLine
        totLines = VbCp.CodeModule.CountOfLines
        'lngLine will return the line containing the search string when found
        bFound = VbCp.CodeModule.Find(StrSearch, lngLine, 1, totLines, 1, ExactMatch, , Not ExactMatch)
        Do Until bFound = False
            DoEvents
            VbCp.CodeModule.CodePane.TopLine = lngLine
            VbCp.CodeModule.CodePane.SetSelection lngLine, 1, lngLine, -1
            If ExactMatch = True Then
                curLine = Trim$(VbCp.CodeModule.lines(lngLine, 1))
                If LCase$(curLine) = LCase$(StrSearch) Then
                    VbCp.CodeModule.DeleteLines lngLine
                Else
                    ' if we are doing an exact match and the line is found but does not meet criteria
                    ' search from next line
                    lngLine = lngLine + 1
                End If
            Else
                curLine = Trim$(VbCp.CodeModule.lines(lngLine, 1))
                If InStr(1, curLine, StrSearch, vbTextCompare) > 0 Then
                    VbCp.CodeModule.DeleteLines lngLine
                End If
            End If
            totLines = VbCp.CodeModule.CountOfLines
            bFound = VbCp.CodeModule.Find(StrSearch, lngLine, 1, totLines, 1, ExactMatch, , Not ExactMatch)
        Loop
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_RemoveCommentedDeclarations(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    Dim totLines As Long
    Dim cntLines As Long
    Dim curLine As String
    totLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(totLines)
    For cntLines = totLines To 1 Step -1
        If UpdateProgress(cntLines, totLines, VbCp.Name & ".Removing Lines With Commented Declarations...") = False Then
            Exit For
        End If
        curLine = VbCp.CodeModule.lines(cntLines, 1)
        curLine = Trim$(curLine)
        If Left$(curLine, 1) = "'" Then
            If InStr(1, curLine, "Dim ") > 0 Then
                If InStr(1, curLine, " As ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                End If
            ElseIf InStr(1, curLine, "Private ") > 0 Then
                If InStr(1, curLine, " As ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                ElseIf InStr(1, curLine, " Const ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                End If
            ElseIf InStr(1, curLine, "Public ") > 0 Then
                If InStr(1, curLine, " As ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                ElseIf InStr(1, curLine, " Const ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                End If
            ElseIf InStr(1, curLine, "Const ") > 0 Then
                If InStr(1, curLine, " As ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                End If
            ElseIf InStr(1, curLine, "Static ") > 0 Then
                If InStr(1, curLine, " As ") > 0 Then
                    VbCp.CodeModule.DeleteLines cntLines
                End If
            End If
        End If
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_SortProcedures(VbCp As VBIDE.VBComponent)
    On Error GoTo Err_Handler:
    Dim lng_cnt As Long
    Dim lng_splitCount As Long
    Dim lng_procStart As Long
    Dim lng_procEnd As Long
    Dim lng_procCount As Long
    Dim lng_propType As Long
    Dim str_procName As String
    Dim str_propType As String
    Dim str_CodeLine As String
    Dim str_currline As String
    Dim str_split() As String
    Dim lng_cnt_Tot As Long
    Dim lng_cnt_Cnt As Long
    Dim lng_splitCount_Tot As Long
    With VbCp.CodeModule
        'examine each line of code starting from the end
        'of declarations to the end of the code module
        lng_cnt_Tot = .CountOfLines
        lng_cnt_Cnt = .CountOfDeclarationLines
        For lng_cnt = lng_cnt_Cnt To lng_cnt_Tot
            'cache current line
            str_currline = LCase$(Trim$(.lines(lng_cnt, 1)))
            ' if the left part of the line is one of the following
            'then we know the current line is the first line
            'of a method of some sort
            If Left$(str_currline, 12) = "private sub " Or Left$(str_currline, 11) = "public sub " Or Left$(str_currline, 11) = "friend sub " Or Left$(str_currline, 17) = "private function " Or Left$(str_currline, 16) = "public function " Or Left$(str_currline, 16) = "friend function " Or Left$(str_currline, 4) = "sub " Or Left$(str_currline, 9) = "function " Or Left$(str_currline, 11) = "friend sub " Then
                'split the line up by " "
                str_split = Split(str_currline)
                'loop through each word of the line split up
                lng_splitCount_Tot = UBound(str_split)
                For lng_splitCount = 0 To lng_splitCount_Tot
                    'if that word is "sub" or "function" then
                    'we know the next word is the methods name
                    If str_split(lng_splitCount) = "sub" Or str_split(lng_splitCount) = "function" Then
                        'cache the next array index of split
                        'which we have just said is the method name
                        str_CodeLine = str_split(lng_splitCount + 1)
                        'every sub or function name follows with
                        'argument parenthases "()" so trim to the
                        'first "("..which leaves us with just the
                        'current methods name
                        str_procName = Mid$(str_CodeLine, 1, (InStr(1, str_CodeLine, "(") - 1))
                        'cache the starting line of this method
                        '(which begins with the first line after the last end sub
                        ' or end function) and the methods last line
                        ' (this methods "end sub" or "end function")
                        lng_procStart = .ProcStartLine(str_procName, vbext_pk_Proc)
                        lng_procEnd = .ProcCountLines(str_procName, vbext_pk_Proc)
                        'store this methods code in the array
                        'with the name of the method as the key
                        ReDim Preserve ArrProcCode(1, lng_procCount)
                        'store the method name
                        ArrProcCode(0, lng_procCount) = str_procName
                        'store the method code
                        ArrProcCode(1, lng_procCount) = .lines(lng_procStart, lng_procEnd)
                        'increment function counter
                        lng_procCount = (lng_procCount + 1)
                    End If
                    Err.Clear
                Next
                '-------------------------------------------
                'properties need to be treated differently because
                'in order to get the start and end line of the proc
                'if its a property we need to know if its let/get/set
                '-------------------------------------------
            ElseIf Left$(str_currline, 17) = "private property " Or Left$(str_currline, 16) = "public property " Or Left$(str_currline, 16) = "friend property " Or Left$(str_currline, 9) = "property " Then
                'split the line up by " "
                str_split = Split(str_currline)
                'loop through each word of the line split up
                lng_splitCount_Tot = UBound(str_split)
                For lng_splitCount = 0 To lng_splitCount_Tot
                    'if that word is "property"  then
                    'we know the next word is the property name
                    If str_split(lng_splitCount) = "property" Then
                        'the next array index of the split is
                        'going to be either "get", "let" or "set"
                        str_propType = str_split(lng_splitCount + 1)
                        'we need to convert the property type to
                        'a numeric value that is compatable with
                        '.ProcStartLine  and  ProcCountLines arguments
                        lng_propType = funcPropNum(str_propType)
                        'cache the next array index of split
                        'which we have just said is the method name
                        str_CodeLine = str_split(lng_splitCount + 2)
                        'every property name follows with
                        'argument parenthases "()" so trim to the
                        'first "("..which leaves us with just the
                        'current methods name
                        str_procName = Mid$(str_CodeLine, 1, (InStr(1, str_CodeLine, "(") - 1))
                        'cache the starting line of this method
                        '(which begins with the first line after the last end sub
                        ' or end function or end property) and the methods last line
                        ' (this properties "end property")
                        lng_procStart = .ProcStartLine(str_procName, lng_propType)
                        lng_procEnd = .ProcCountLines(str_procName, lng_propType)
                        'store this methods code in the array
                        'with the name of the method as the key
                        ReDim Preserve ArrProcCode(1, lng_procCount)
                        'store the property name + type
                        ArrProcCode(0, lng_procCount) = str_procName & str_propType
                        'store the properties code
                        ArrProcCode(1, lng_procCount) = .lines(lng_procStart, lng_procEnd)
                        'increment function counter
                        lng_procCount = (lng_procCount + 1)
                    End If
                    Err.Clear
                Next
            End If
            Err.Clear
        Next
        'erase the code in this module so its blank
        .DeleteLines (.CountOfDeclarationLines + 1), (.CountOfLines - (.CountOfDeclarationLines + 1))
        'pause time for createion of module
        'insert the sorted code into the module
        .AddFromString AlphaSortString(arrToSentence)
        'pause time for createion of module
        'set focus back to the original codemodule
        'erase array
        Erase ArrProcCode
    End With
    Err.Clear
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".cBarButtonEvents1_Click"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description, , App.Title
    Resume Next
    Err.Clear
End Sub
Private Sub Code_SpeedNextLoopEnd(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim currLine As String
    tLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then tLines = EndLine
    Call InitProgress(tLines)
    For tCount = StartLine To tLines
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Speeding Next Loop Counter") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        currLine = ImprovedNextLoop(currLine)
        VbCp.CodeModule.ReplaceLine tCount, currLine
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Code_TrimLines(VbCp As VBIDE.VBComponent, Optional StartLine As Long = 1, Optional EndLine As Long = -1)
    On Error Resume Next
    Dim totLines As Long
    Dim cntLines As Long
    Dim curLine As String
    totLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then totLines = EndLine
    Call InitProgress(totLines)
    For cntLines = StartLine To totLines
        If UpdateProgress(cntLines, totLines, VbCp.Name & ".Trimming Lines") = False Then
            Exit For
        End If
        curLine = VbCp.CodeModule.lines(cntLines, 1)
        curLine = Trim$(curLine)
        VbCp.CodeModule.ReplaceLine cntLines, curLine
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Private Sub Dao_CreateDatabase(ByVal DbName As String, Optional ByVal Overwrite As Boolean = False, Optional ByVal Version As DAO.DatabaseTypeEnum = dbVersion40)
    On Error Resume Next
    Dim fExist As Boolean
    Dim wrkDefault As DAO.Workspace
    Dim dbsNew As DAO.Database
    fExist = boolFileExists(DbName)
    If fExist = True Then
        If Overwrite = True Then
            Kill DbName
        Else
            Err.Clear
            Exit Sub
        End If
    End If
    Set wrkDefault = DAO.DBEngine.Workspaces(0)
    Set dbsNew = wrkDefault.CreateDatabase(DbName, DAO.dbLangGeneral, Version)
    Set dbsNew = Nothing
    Set wrkDefault = Nothing
    Err.Clear
End Sub
Private Sub Dao_CreateTable(ByVal DbName As String, ByVal dbTable As String, ByVal fldName As String, Optional ByVal fldType As String = vbNullString, _
    Optional ByVal fldSize As String = vbNullString, Optional ByVal Fldidx As String = vbNullString, Optional ByVal FldAutoIncrement As String = vbNullString)
    On Error Resume Next
    Dim spFlds() As String
    Dim spType() As String
    Dim spSize() As String
    Dim spIndx() As String
    Dim spAuto() As String
    Dim totFld As Integer
    Dim totIdx As Integer
    Dim NewFld As DAO.Field
    Dim NewIdx As DAO.Index
    Dim NewTb As DAO.TableDef
    Dim NewDb As DAO.Database
    Dim newCnt As Integer
    Dim newPos As Integer
    Dim NewType As Integer
    Call StringParse(spFlds, fldName, ",")
    Call StringParse(spType, fldType, ",")
    Call StringParse(spSize, fldSize, ",")
    Call StringParse(spIndx, Fldidx, ",")
    Call StringParse(spAuto, FldAutoIncrement, ",")
    totFld = UBound(spFlds)
    totIdx = UBound(spIndx)
    ReDim Preserve spType(totFld)
    ReDim Preserve spSize(totFld)
    Set NewDb = DAO.OpenDatabase(DbName)
    Set NewTb = NewDb.CreateTableDef(dbTable)
    For newCnt = 1 To totFld
        spType(newCnt) = Trim$(spType(newCnt))
        spFlds(newCnt) = Trim$(spFlds(newCnt))
        spSize(newCnt) = Trim$(spSize(newCnt))
        If Len(spType(newCnt)) = 0 Then
            spType(newCnt) = "Text"
        End If
        If Len(spSize(newCnt)) = 0 Then
            spSize(newCnt) = "255"
        End If
        NewType = Dao_FieldType(spType(newCnt))
        spFlds(newCnt) = spFlds(newCnt)
        Set NewFld = NewTb.CreateField(spFlds(newCnt), NewType)
        Select Case NewType
        Case dbText
            NewFld.AllowZeroLength = True
            NewFld.Size = spSize(newCnt)
        Case dbMemo
            NewFld.AllowZeroLength = True
        Case dbLong, dbInteger, dbDouble
            NewFld.DefaultValue = vbNullString
        End Select
        If MvSearch(FldAutoIncrement, spFlds(newCnt), ",") > 0 Then
            NewFld.Attributes = DAO.dbAutoIncrField
        End If
        NewTb.Fields.Append NewFld
        Err.Clear
    Next
    For newCnt = 1 To totIdx
        spIndx(newCnt) = Trim$(spIndx(newCnt))
        newPos = spIndx(newCnt)
        Set NewIdx = NewTb.CreateIndex(spFlds(newPos))
        Set NewFld = NewIdx.CreateField(spFlds(newPos))
        NewIdx.Fields.Append NewFld
        NewTb.Indexes.Append NewIdx
        Err.Clear
    Next
NextSection:
    NewDb.TableDefs.Append NewTb
    Select Case Err
    Case 3010, 3012       ' already exists/locked
        NewDb.TableDefs.Delete dbTable
        GoTo NextSection
    Case 3006, 3009, 3008
        GoTo NextSection1
    End Select
NextSection1:
    NewDb.Close
    Erase spFlds
    Erase spType
    Erase spSize
    Erase spIndx
    Set NewFld = Nothing
    Set NewIdx = Nothing
    Set NewTb = Nothing
    Set NewDb = Nothing
    Err.Clear
End Sub
Public Function Dao_DatabaseCompress(ByVal Datab As String) As Boolean
    On Error GoTo Compact_Repair_Error
    Dim RepairDb As String
    Dim TemporDb As String
    Dim TestDb As DAO.Database
    Dim Path As String
    Path = StringGetFileToken(Datab, "p")
    RepairDb = Datab
    TemporDb = ExactPath(Path) & "\tmp.mdb"
    If boolFileExists(TemporDb) = True Then
        Kill TemporDb
    End If
    Set TestDb = DAO.OpenDatabase(RepairDb, True, False)  ' open exclusive, read write
    TestDb.Close
    Set TestDb = Nothing
    DAO.DBEngine.RepairDatabase RepairDb
    DAO.DBEngine.CompactDatabase RepairDb, TemporDb
    FileCopy TemporDb, RepairDb
    Kill TemporDb
    Dao_DatabaseCompress = True
    Err.Clear
    Exit Function
Compact_Repair_Error:
    Select Case Err
    Case 401
        Resume Next
    Case Else
        Dao_DatabaseCompress = False
        Set TestDb = Nothing
        Err.Clear
        Exit Function
    End Select
    Err.Clear
End Function
Private Function Dao_FieldType(ByVal StrType As String) As Integer
    On Error Resume Next
    Dim StrTp As String
    StrTp = LCase$(Trim$(StrType))
    Select Case StrTp
    Case "big", "bigint": Dao_FieldType = dbBigInt
    Case "bi", "bin", "binary": Dao_FieldType = dbLongBinary
    Case "cha", "char": Dao_FieldType = dbChar
    Case "dec", "decimal": Dao_FieldType = dbDecimal
    Case "flo", "float": Dao_FieldType = dbFloat
    Case "gui", "guid": Dao_FieldType = dbGUID
    Case "tim", "time": Dao_FieldType = dbTime
    Case "tis", "timestamp": Dao_FieldType = dbTimeStamp
    Case "num", "numeric": Dao_FieldType = dbNumeric
    Case "var", "varbinary": Dao_FieldType = dbVarBinary
    Case "bo", "boo", "boolean": Dao_FieldType = dbBoolean
    Case "by", "byt", "byte": Dao_FieldType = dbByte
    Case "in", "int", "integer": Dao_FieldType = dbInteger
    Case "lo", "lon", "long": Dao_FieldType = dbLong
    Case "cu", "cur", "currency": Dao_FieldType = dbCurrency
    Case "si", "sin", "single": Dao_FieldType = dbSingle
    Case "do", "dou", "double": Dao_FieldType = dbDouble
    Case "da", "dat", "date": Dao_FieldType = dbDate
    Case "te", "tex", "text": Dao_FieldType = dbText
    Case "lob", "longbinary", "long binary": Dao_FieldType = dbLongBinary
    Case "ole", "object": Dao_FieldType = dbLongBinary
    Case "me", "mem", "memo": Dao_FieldType = dbMemo
    Case "st", "str", "string": Dao_FieldType = dbText
    End Select
    Err.Clear
End Function
Private Function Dao_TableExists(ByVal Dbase As String, ByVal TbName As String) As Boolean
    On Error Resume Next
    Dim DatCt As Long
    Dim StrDt As String
    Dim zCnt As Long
    Dim Db As DAO.Database
    TbName = LCase$(TbName)
    Dao_TableExists = False
    Set Db = DAO.OpenDatabase(Dbase)
    With Db
        zCnt = .TableDefs.Count - 1
        For DatCt = 0 To zCnt
            StrDt = LCase$(.TableDefs(DatCt).Name)
            Select Case StrDt
            Case TbName
                Dao_TableExists = True
                Exit For
            End Select
            Err.Clear
        Next
    End With
    Db.Close
    Set Db = Nothing
    Err.Clear
End Function
Public Sub Destroy()
    On Error Resume Next
    DestroyButton ccChangeFont, ccChangeFontEvents
    DestroyButton ccMySQL, ccMySQLEvents
    DestroyButton ccMySQL1, ccMySQL1Events
    DestroyButton ccSwap, ccSwapEvents
    DestroyButton ccCommentBlock, ccCommentBlockEvents
    DestroyButton ccUnCommentBlock, ccUnCommentBlockEvents
    DestroyButton ccConvertToUpperCase, ccConvertToUpperCaseEvents
    DestroyButton ccConvertToLowerCase, ccConvertToLowerCaseEvents
    DestroyButton ccConvertToProperCase, ccConvertToProperCaseEvents
    DestroyButton ccClearText, ccClearTextEvents
    DestroyButton ccInno, ccInnoEvents
    DestroyButton ccComment, ccCommentEvents
    DestroyButton ccExplore, ccExploreEvents
    DestroyButton ccRemoveError, ccRemoveErrorEvents
    DestroyButton ccRemoveBlanks, ccRemoveBlanksEvents
    DestroyButton ccTrimLines, ccTrimLinesEvents
    DestroyButton ccRemoveLineContinuation, ccRemoveLineContinuationEvents
    DestroyButton ccBreakDeclarations, ccBreakDeclarationsEvents
    DestroyButton ccSpeedStringFunctions, ccSpeedStringFunctionsEvents
    DestroyButton ccMoveComments, ccMoveCommentsEvents
    DestroyButton ccSpeedNextLoopEnd, ccSpeedNextLoopEndEvents
    DestroyButton ccInsertDoEvents, ccInsertDoEventsEvents
    DestroyButton ccCommentQuote, ccCommentQuoteEvents
    DestroyButton ccPassByVal, ccPassByValEvents
    DestroyButton ccPassByRef, ccPassByRefEvents
    DestroyButton ccInitializeOptionalVariables, ccInitializeOptionalVariablesEvents
    DestroyButton ccRemoveEmptyProcedures, ccRemoveEmptyProceduresEvents
    DestroyButton ccNormalizeIf, ccNormalizeIfEvents
    DestroyButton ccBreakComplexForLoops, ccBreakComplexForLoopsEvents
    DestroyButton ccMoveDeclarations, ccMoveDeclarationsEvents
    DestroyButton ccLineStartWithComments, ccLineStartWithCommentsEvents
    DestroyButton ccInsertOnErrorResumeNext, ccInsertOnErrorResumeNextEvents
    DestroyButton ccFormat, ccFormatEvents
    DestroyButton ccRemoveLike, ccRemoveLikeEvents
    DestroyButton ccSort, ccSortEvents
    DestroyButton ccIntellisense, ccIntellisenseEvents
    DestroyButton ccValidDateMask, ccValidDateMaskEvents
    DestroyButton ccTextBoxHilite, ccTextBoxHiliteEvents
    DestroyButton ccProperCase, ccProperCaseEvents
    DestroyButton ccBackUp, ccBackUpEvents
    DestroyButton ccAbout, ccAboutEvents
    DestroyButton ccClose, ccCloseEvents
    DestroyButton ccToolTips, ccToolTipsEvents
    DestroyButton ccAddCode, ccAddCodeEvents
    DestroyButton ccVariables, ccVariablesEvents
    DestroyButton ccProcedures, ccProceduresEvents
    DestroyButton ccRemove, ccRemoveEvents
    Set VBInst = Nothing
    Err.Clear
End Sub
Private Function FileData(ByVal Filename As String) As String
    On Error Resume Next
    Dim sLen As Long
    Dim myBuf As String
    Dim FileNum As Long
    Dim Size As Double
    FileNum = FreeFile
    Size = FileLen(Filename)
    myBuf = String$(Size, "*")
    Open Filename For Input Access Read As #FileNum
    sLen = LOF(FileNum)
    FileData = Input(sLen, #FileNum)
    Close #FileNum
    Err.Clear
End Function
Private Function FixDeclaration(ByVal StrLine As String) As String
    On Error Resume Next
    StrLine = Replace$(StrLine, "%", " As Integer")
    StrLine = Replace$(StrLine, "&", " As Long")
    StrLine = Replace$(StrLine, "!", " As Single")
    StrLine = Replace$(StrLine, "#", " As Double")
    StrLine = Replace$(StrLine, "@", " As Currency")
    StrLine = Replace$(StrLine, "$", " As String")
    StrLine = TrimDeclaration(StrLine)
    FixDeclaration = StrLine
    Err.Clear
End Function
Private Function funcPropNum(sVal As String) As Long
    On Error Resume Next
    '-------------------------------------------
    'this function will convert a string value
    'of a properties type (let,get,set) and convert
    'it to a long value compatable with .ProcStartLine
    ' and .ProcCountLines of VBinst.ActiveCodePane.CodeModule
    '-------------------------------------------
    Select Case sVal
    Case Is = "get"
        funcPropNum = 3
    Case Is = "let"
        funcPropNum = 1
    Case Is = "set"
        funcPropNum = 2
    End Select
    Err.Clear
End Function
Public Function ImprovedNextLoop(ByVal StrLine As String) As String
    On Error Resume Next
    Dim sParts() As String
    Dim pTot As Long
    Dim pCnt As Long
    Dim pStr As String
    If InStr(1, StrLine, "Next") = 0 Then
        ImprovedNextLoop = StrLine
    Else
        sParts = Split(StrLine, ":")
        pTot = UBound(sParts)
        For pCnt = 0 To pTot
            pStr = Trim$(sParts(pCnt))
            pStr = Trim$(pStr)
            If Left$(pStr, 5) = "Next " Then
                sParts(pCnt) = "Next"
            End If
            Err.Clear
        Next
        ImprovedNextLoop = Join(sParts, ":")
    End If
    Err.Clear
End Function
Public Function ImprovedStringFunction(ByVal StrLine As String) As String
    On Error Resume Next
    Dim lngF As Long
    Dim lngN As Long
    Dim StrF As String
    Dim StrN As String
    Dim SeekLen As Long
    Dim SeekFnd As String
    lngF = UBound(StrFunc)
    For lngN = 1 To lngF
        ' if function name starts with a space
        StrF = StrFunc(lngN)
        StrN = StrF
        StrF = " " & StrF & "("
        StrN = " " & StrN & "$("
        StrLine = Replace$(StrLine, StrF, StrN)
        ' if function name is called
        StrF = StrFunc(lngN)
        StrN = StrF
        StrF = "(" & StrF & "("
        StrN = "(" & StrN & "$("
        StrLine = Replace$(StrLine, StrF, StrN)
        ' if function name is the beginning of the sentence
        StrF = StrFunc(lngN)
        StrN = StrF
        StrF = StrF & "("
        StrN = StrN & "$("
        SeekLen = Len(StrF)
        SeekFnd = Left$(StrLine, SeekLen)
        If SeekFnd = StrF Then
            StrLine = Replace$(StrLine, StrF, StrN)
        End If
        Err.Clear
    Next
    ImprovedStringFunction = StrLine
    Err.Clear
End Function
Public Sub IniStrFunctions()
    On Error Resume Next
    ReDim StrFunc(24) As String
    StrFunc(1) = "Space"
    StrFunc(2) = "UCase"
    StrFunc(3) = "Left"
    StrFunc(4) = "Format"
    StrFunc(5) = "LCase"
    StrFunc(6) = "Trim"
    StrFunc(7) = "Hex"
    StrFunc(8) = "Mid"
    StrFunc(9) = "Chr"
    StrFunc(10) = "ChrB"
    StrFunc(11) = "LeftB"
    StrFunc(12) = "RightB"
    StrFunc(13) = "MidB"
    StrFunc(14) = "LTrim"
    StrFunc(15) = "RTrim"
    StrFunc(16) = "Right"
    StrFunc(17) = "Dir"
    StrFunc(18) = "String"
    StrFunc(19) = "Replace"
    StrFunc(20) = "Command"
    StrFunc(21) = "CurDir"
    StrFunc(22) = "Environ"
    StrFunc(23) = "Oct"
    StrFunc(24) = "Str"
    Err.Clear
End Sub
Public Sub Initialize(YourCommandBar As Office.CommandBar)
    On Error GoTo Err_Handler:
    'Set ccAboutEvents = AddButton(YourCommandBar, ccAbout, "About...", "About Kimmo - Cool Source")
    Set ccBackUpEvents = AddButton(YourCommandBar, ccBackUp, vbNullString, "Backup project", , , , 942)
    Set ccInnoEvents = AddButton(YourCommandBar, ccInno, vbNullString, "Inno Setup Compiler", , , , 304)
    'Set ccExploreEvents = AddButton(YourCommandBar, ccExplore, vbNullString, "Project Explorer", , , , 2556)
    Set ccBreakDeclarationsEvents = AddButton(YourCommandBar, ccBreakDeclarations, vbNullString, "Break multiple declarations", , , , 2476)
    Set ccVariablesEvents = AddButton(YourCommandBar, ccVariablesEvents, vbNullString, "Check variable declarations and comment unused", , , , 503)
    Set ccRemoveEvents = AddButton(YourCommandBar, ccRemove, vbNullString, "Remove commented declarations", , , , 1592)
    Set ccProceduresEvents = AddButton(YourCommandBar, ccProcedures, vbNullString, "Check procedure calls and comment unused", , , , 346)
    Set ccRemoveEmptyProceduresEvents = AddButton(YourCommandBar, ccRemoveEmptyProcedures, vbNullString, "Remove empty procedures", , , , 215)
    Set ccInitializeOptionalVariablesEvents = AddButton(YourCommandBar, ccInitializeOptionalVariables, vbNullString, "Initialize optional variables", , , , 288)
    Set ccMoveDeclarationsEvents = AddButton(YourCommandBar, ccMoveDeclarations, vbNullString, "Move declarations to beginning of procedures", , , , 319)
    Set ccPassByValEvents = AddButton(YourCommandBar, ccPassByVal, vbNullString, "Pass string variables by value", , , , 1058)
    Set ccPassByRefEvents = AddButton(YourCommandBar, ccPassByRef, vbNullString, "Pass string variables by reference", , , , 1061)
    Set ccIntellisenseEvents = AddButton(YourCommandBar, ccIntellisense, vbNullString, "Insert intellisense code", , , , "1662")
    Set ccValidDateMaskEvents = AddButton(YourCommandBar, ccValidDateMask, vbNullString, "Insert date mask validation code", , , , "600")
    Set ccTextBoxHiliteEvents = AddButton(YourCommandBar, ccTextBoxHilite, vbNullString, "Insert highlight code", , , , 462)
    Set ccProperCaseEvents = AddButton(YourCommandBar, ccProperCase, vbNullString, "Insert propercase code", , , , 309)
    Set ccToolTipsEvents = AddButton(YourCommandBar, ccToolTips, vbNullString, "Update tool tips", , , , 1546)
    Set ccAddCodeEvents = AddButton(YourCommandBar, ccAddCode, vbNullString, "Insert code after / before particular line", , , , 240)
    Set ccRemoveBlanksEvents = AddButton(YourCommandBar, ccRemoveBlanks, vbNullString, "Remove blank lines", , , , 234)
    Set ccRemoveLikeEvents = AddButton(YourCommandBar, ccRemoveLike, vbNullString, "Remove lines containing...", , , , 1953)
    Set ccRemoveLineContinuationEvents = AddButton(YourCommandBar, ccRemoveLineContinuation, vbNullString, "Remove line continuation", , , , 464)
    Set ccCommentQuoteEvents = AddButton(YourCommandBar, ccCommentQuote, vbNullString, "Find lines starting with comment and quote", , , , 240)
    Set ccSortEvents = AddButton(YourCommandBar, ccSort, vbNullString, "Sort procedures", , , , 3157)
    Set ccTrimLinesEvents = AddButton(YourCommandBar, ccTrimLines, vbNullString, "Trim lines", , , , 1039)
    Set ccSpeedStringFunctionsEvents = AddButton(YourCommandBar, ccSpeedStringFunctions, vbNullString, "Speed string functions", , , , 879)
    Set ccMoveCommentsEvents = AddButton(YourCommandBar, ccMoveComments, vbNullString, "Move comments to line below in if then else statements", , , , 192)
    Set ccCommentEvents = AddButton(YourCommandBar, ccComment, vbNullString, "Insert comments", , , , 127)
    Set ccLineStartWithCommentsEvents = AddButton(YourCommandBar, ccLineStartWithComments, vbNullString, "Extract lines starting with comments", , , , 183)
    Set ccSpeedNextLoopEndEvents = AddButton(YourCommandBar, ccSpeedNextLoopEnd, vbNullString, "Speed next loop ending", , , , 504)
    Set ccInsertDoEventsEvents = AddButton(YourCommandBar, ccInsertDoEvents, vbNullString, "Insert DoEvents on loops", , , , 620)
    Set ccNormalizeIfEvents = AddButton(YourCommandBar, ccNormalizeIf, vbNullString, "Break if statements", , , , 2213)
    Set ccBreakComplexForLoopsEvents = AddButton(YourCommandBar, ccBreakComplexForLoops, vbNullString, "Speed complex for loops", , , , 466)
    Set ccInsertOnErrorResumeNextEvents = AddButton(YourCommandBar, ccInsertOnErrorResumeNext, vbNullString, "Insert on error resume next", , , , 674)
    Set ccRemoveErrorEvents = AddButton(YourCommandBar, ccRemoveError, vbNullString, "Remove on error resume next", , , , 241)
    Set ccFormatEvents = AddButton(YourCommandBar, ccFormat, vbNullString, "Format source code including indents", , , , 737)
    Set ccClearTextEvents = AddButton(YourCommandBar, ccClearText, vbNullString, "Clear text property of text boxes", , , , 1671)
    Set ccConvertToUpperCaseEvents = AddButton(YourCommandBar, ccConvertToUpperCase, vbNullString, "Convert text to upper case", , , , 62)
    Set ccConvertToLowerCaseEvents = AddButton(YourCommandBar, ccConvertToLowerCase, vbNullString, "Convert text to lower case", , , , 63)
    Set ccConvertToProperCaseEvents = AddButton(YourCommandBar, ccConvertToProperCase, "P", "Convert text to proper case")
    Set ccSwapEvents = AddButton(YourCommandBar, ccSwap, vbNullString, "Swap delimited text separated by =", , , , 1977)
    Set ccCommentBlockEvents = AddButton(YourCommandBar, ccCommentBlock, vbNullString, "Comment block", , , , 1046)
    Set ccUnCommentBlockEvents = AddButton(YourCommandBar, ccUnCommentBlock, vbNullString, "Remove Comments", , , , 2189)
    Set ccFontNameEvents = AddCommandBarControl(YourCommandBar, ccFontName, vbNullString, "Select the font name to change the font of controls to", , , "cboFontName", , True)
    Set ccFontSizeEvents = AddCommandBarControl(YourCommandBar, ccFontSize, vbNullString, "Select the font size to change the font of controls to", , , "cboFontSize")
    Set ccChangeFontEvents = AddButton(YourCommandBar, ccChangeFont, vbNullString, "Change font property of controls", , , , 625)
    ccFontName.BeginGroup = False
    Set ccCompileCompressEvents = AddButton(YourCommandBar, ccCompileCompress, vbNullString, "Compile and compress project", , , , 1023)
    Set ccCloseEvents = AddButton(YourCommandBar, ccClose, vbNullString, "Close all windows", , , , 303)
    'Set ccMySQLEvents = AddButton(YourCommandBar, ccMySQL, "MySQL>", "MySQL what must be updated")
    'Set ccMySQL1Events = AddButton(YourCommandBar, ccMySQL, "MySQL<", "MySQL what is being read")
    RefreshFonts
    IniStrFunctions
    Err.Clear
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".Initialize"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description, , App.Title
    Resume Next
    Err.Clear
End Sub

Public Sub Initialize1(YourCommandBar As Office.CommandBar)
    On Error GoTo Err_Handler:
    Set ccUpdateRsEvents = AddButton(YourCommandBar, ccUpdateRs, vbNullString, "Fix UpdateRs Code", , , , 1025)
    Set ccWPFEvents = AddButton(YourCommandBar, ccWPF, vbNullString, "WPF Ribbon", , , , 1050)
    Err.Clear
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".Initialize1"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description, , App.Title
    Resume Next
    Err.Clear
End Sub

Private Sub RefreshFonts()
    On Error Resume Next
    Dim lngLength As Long
    Dim rsCnt As Long
    Dim rsTot As Long
    ccFontName.Clear
    ccFontSize.Clear
    ccFontSize.AddItem 8
    ccFontSize.AddItem 9
    ccFontSize.AddItem 10
    ccFontSize.AddItem 11
    ccFontSize.AddItem 12
    ccFontSize.AddItem 14
    ccFontSize.AddItem 16
    ccFontSize.AddItem 18
    ccFontSize.AddItem 20
    ccFontSize.AddItem 22
    ccFontSize.AddItem 24
    ccFontSize.AddItem 26
    ccFontSize.AddItem 28
    ccFontSize.AddItem 36
    ccFontSize.AddItem 48
    ccFontSize.AddItem 72
    
    frmAddCode.cboFonts.Clear
    lngLength = Len(Screen.Fonts(0))
    rsTot = Screen.FontCount - 1
    For rsCnt = 0 To rsTot
        frmAddCode.cboFonts.AddItem Screen.Fonts(rsCnt)
        If Len(Screen.Fonts(rsCnt)) > lngLength Then lngLength = Len(Screen.Fonts(rsCnt))
    Next
    rsTot = frmAddCode.cboFonts.ListCount - 1
    For rsCnt = 0 To rsTot
        ccFontName.AddItem frmAddCode.cboFonts.list(rsCnt)
    Next
    ccFontName.ListIndex = 1
    ccFontSize.ListIndex = 1
    ccFontName.Width = 6 * lngLength
    ccFontSize.Width = 50
    Err.Clear
End Sub

Private Function IntellisenseCode(ByVal ControlName As String) As String
    On Error Resume Next
    Dim txtName As String
    Dim txtIdx As String
    Dim txtCode As String
    txtName = Split(ControlName, ",")(0)
    txtIdx = Split(ControlName, ",")(1)
    txtCode = vbNullString
    If txtIdx = "False" Then
        txtCode = "Private Sub " & txtName & "_KeyUp(KeyCode As Integer, Shift As Integer)" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & "If " & txtName & ".MultiLine = False Then Intel.Sense App.Path, " & txtName & ", KeyCode, Shift" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & "End Sub"
    Else
        txtCode = "Private Sub " & txtName & "_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & "If " & txtName & "(Index).MultiLine = False Then Intel.Sense App.Path, " & txtName & "(Index), KeyCode, Shift" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & "End Sub"
    End If
    IntellisenseCode = txtCode
    Err.Clear
End Function
Private Function ValidateDateMaskCode(ByVal ControlName As String) As String
    On Error Resume Next
    Dim txtName As String
    Dim txtIdx As String
    Dim txtCode As String
    txtName = Split(ControlName, ",")(0)
    txtIdx = Split(ControlName, ",")(1)
    txtCode = vbNullString
    If txtIdx = "False" Then
        txtCode = "Private Sub " & txtName & "_Validate(Cancel As Boolean)" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & txtName & ".Text = MyADO.FixMaskDate(" & txtName & ")" & vbNewLine
        txtCode = txtCode & vbTab & "If MyADO.DateError = 1 Then" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & txtName & ".SetFocus" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & txtName & "_GotFocus" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & "Cancel = True" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & vbTab & "Exit Sub" & vbNewLine
        txtCode = txtCode & "End If" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & "End Sub"
    Else
        txtCode = "Private Sub " & txtName & "_Validate(Index As Integer, cancel As Boolean)" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & txtName & "(Index).Text = MyADO.FixMaskDate(" & txtName & "(Index))" & vbNewLine
        txtCode = txtCode & vbTab & "If MyADO.DateError = 1 Then" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & txtName & "(Index).SetFocus" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & txtName & "(Index)_GotFocus" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & "Cancel = True" & vbNewLine
        txtCode = txtCode & vbTab & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & vbTab & vbTab & "Exit Sub" & vbNewLine
        txtCode = txtCode & "End If" & vbNewLine
        txtCode = txtCode & vbTab & vbNullString & vbNewLine
        txtCode = txtCode & "End Sub"
    End If
    ValidateDateMaskCode = txtCode
    Err.Clear
End Function
Private Function IsEndOfMethod(ByVal currLine As String) As Boolean
    On Error Resume Next
    Select Case currLine
    Case "End Sub", "End Function"
        IsEndOfMethod = True
    Case Else
        IsEndOfMethod = False
    End Select
    Err.Clear
End Function
Private Function IsMethod(ByVal currLine As String) As Boolean
    On Error Resume Next
    If Left$(currLine, Len("Private Sub ")) = "Private Sub " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Private Function ")) = "Private Function " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Function ")) = "Function " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Sub ")) = "Sub " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Public Sub ")) = "Public Sub " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Public Function ")) = "Public Function " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Public Property ")) = "Public Property " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Property ")) = "Property " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Private Property ")) = "Private Property " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Friend Property ")) = "Friend Property " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Friend Function ")) = "Friend Function " Then
        IsMethod = True
    ElseIf Left$(currLine, Len("Friend Sub ")) = "Friend Sub " Then
        IsMethod = True
    Else
        IsMethod = False
    End If
    Err.Clear
End Function
Private Function IsMethodEnd(ByVal currLine As String) As Boolean
    On Error Resume Next
    Select Case currLine
    Case "End Sub", "Exit Function", "Exit Sub", "End Function", "End Property", "Exit Property", "Return", "Next"
        IsMethodEnd = True
    Case Else
        IsMethodEnd = False
    End Select
    Err.Clear
End Function
Private Function IsSourceComplex(ByVal SourceCode As String) As Boolean
    On Error Resume Next
    Dim spLines() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim ForPos As Long
    Dim EquPos As Long
    Dim TooPos As Long
    Dim StePos As Long
    Dim varNam As String
    Dim CounterOne As String
    Dim CounterTwo As String
    Dim CounterThree As String
    Dim complexCnt As Integer
    complexCnt = 0
    spLines = Split(SourceCode, vbNewLine)
    spTot = UBound(spLines)
    For spCnt = 0 To spTot
        spStr = Trim$(spLines(spCnt))
        If Left$(spStr, 4) = "For " Then
            If Mid$(spStr, 5, 4) = "Each" Then
                GoTo NextLine
            End If
            ForPos = 1
            EquPos = InStr(1, spStr, " = ")
            TooPos = InStr(1, spStr, " To ")
            StePos = InStr(1, spStr, " Step ")
            varNam = Mid$(spStr, 5, EquPos - 5)
            CounterOne = Mid$(spStr, (EquPos + 3), (TooPos - (EquPos + 3)))
            If StePos > 0 Then
                CounterTwo = Mid$(spStr, (TooPos + 4), (StePos - (TooPos + 4)))
                CounterThree = Mid$(spStr, StePos + 6)
            Else
                CounterTwo = Mid$(spStr, (TooPos + 4))
            End If
            If InStr(1, CounterOne, ".") > 0 Then
                complexCnt = complexCnt + 1
            ElseIf InStr(1, CounterOne, "(") > 0 Then
                complexCnt = complexCnt + 1
            ElseIf InStr(1, CounterOne, " ") > 0 Then
                complexCnt = complexCnt + 1
            End If
            If InStr(1, CounterTwo, ".") > 0 Then
                complexCnt = complexCnt + 1
            ElseIf InStr(1, CounterTwo, "(") > 0 Then
                complexCnt = complexCnt + 1
            ElseIf InStr(1, CounterTwo, " ") > 0 Then
                complexCnt = complexCnt + 1
            End If
        End If
NextLine:
        Err.Clear
    Next
    If complexCnt = 0 Then
        IsSourceComplex = False
    Else
        IsSourceComplex = True
    End If
    Err.Clear
End Function
Public Function OptionalizedProcedure(ByVal strDeclaration As String) As String
    On Error Resume Next
    Dim sArguements As String
    Dim sNewArguements As String
    Dim sProcedure As String
    Dim nProcedure As String
    Select Case Right$(strDeclaration, 1)
    Case "_"
        OptionalizedProcedure = FixDeclaration(strDeclaration)
    Case Else
        sArguements = StringArguements(strDeclaration)
        sNewArguements = StringOptionalInitialize(sArguements)
        sProcedure = StringProcedure(strDeclaration)
        nProcedure = Replace$(sProcedure, "()", "(" & sNewArguements & ")")
        OptionalizedProcedure = FixDeclaration(nProcedure)
    End Select
    Err.Clear
End Function
Private Function RecodeComplex(ByVal SourceCode As String, colVariables As Collection) As String
    On Error Resume Next
    Dim spLines() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim ForPos As Long
    Dim EquPos As Long
    Dim TooPos As Long
    Dim StePos As Long
    Dim varNam As String
    Dim CounterOne As String
    Dim CounterTwo As String
    Dim CounterThree As String
    Dim newTyp As String
    Dim linBefore As String
    Dim linAfter As String
    Dim decBefore As String
    Dim decAfter As String
    spLines = Split(SourceCode, vbNewLine)
    spTot = UBound(spLines)
    For spCnt = 0 To spTot
        spStr = Trim$(spLines(spCnt))
        If Left$(spStr, 4) = "For " Then
            If Mid$(spStr, 5, 4) = "Each" Then
                GoTo NextLine
            End If
            ForPos = 1
            EquPos = InStr(1, spStr, " = ")
            TooPos = InStr(1, spStr, " To ")
            StePos = InStr(1, spStr, " Step ")
            varNam = Mid$(spStr, 5, EquPos - 5)
            CounterOne = Mid$(spStr, (EquPos + 3), (TooPos - (EquPos + 3)))
            If StePos > 0 Then
                CounterTwo = Mid$(spStr, (TooPos + 4), (StePos - (TooPos + 4)))
                CounterThree = Mid$(spStr, StePos + 6)
            Else
                CounterTwo = Mid$(spStr, (TooPos + 4))
            End If
            If InStr(1, CounterOne, ".") > 0 Then
                GoSub CorrectCode1
            ElseIf InStr(1, CounterOne, "(") > 0 Then
                GoSub CorrectCode1
            ElseIf InStr(1, CounterOne, " ") > 0 Then
                GoSub CorrectCode1
            End If
            If InStr(1, CounterTwo, ".") > 0 Then
                GoSub CorrectCode
            ElseIf InStr(1, CounterTwo, "(") > 0 Then
                GoSub CorrectCode
            ElseIf InStr(1, CounterTwo, " ") > 0 Then
                GoSub CorrectCode
            End If
            spLines(spCnt) = spStr
        End If
NextLine:
        Err.Clear
    Next
    RecodeComplex = MvFromArray(spLines, vbNewLine)
    Err.Clear
    Exit Function
CorrectCode:
    newTyp = SearchCollection(colVariables, varNam)
    decAfter = "Dim " & varNam & "_Tot As " & newTyp
    linAfter = varNam & "_Tot = " & CounterTwo
    spStr = Replace$(spStr, CounterTwo, varNam & "_Tot")
    spStr = decAfter & vbNewLine & linAfter & vbNewLine & spStr
    Err.Clear
    Return
CorrectCode1:
    newTyp = SearchCollection(colVariables, varNam)
    decBefore = "Dim " & varNam & "_Cnt As " & newTyp
    linBefore = varNam & "_Cnt = " & CounterOne
    spStr = Replace$(spStr, CounterOne, varNam & "_Cnt")
    spStr = decBefore & vbNewLine & linBefore & vbNewLine & spStr
    Err.Clear
    Return
    Err.Clear
End Function
Private Function SearchCollection(colName As Collection, ByVal StrSearch As String) As String
    On Error Resume Next
    Dim colTot As Long
    Dim colCnt As Long
    Dim colStr As String
    Dim ItemKey As String
    StrSearch = LCase$(StrSearch)
    colTot = colName.Count
    For colCnt = 1 To colTot
        colStr = colName.Item(colCnt)
        ItemKey = LCase$(Split(colStr, ",")(0))
        If ItemKey = StrSearch Then
            SearchCollection = Split(colStr, ",")(1)
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function StringArguements(ByVal strDeclaration As String) As String
    On Error Resume Next
    Dim fBracket As Long
    Dim sBracket As Long
    fBracket = InStr(1, strDeclaration, "(")
    sBracket = InStrRev(strDeclaration, ")")
    Select Case sBracket
    Case Is = fBracket + 1
        Err.Clear
        Exit Function
    Case Else
        StringArguements = Mid$(strDeclaration, fBracket + 1, (sBracket - fBracket - 1))
    End Select
    Err.Clear
End Function
Public Function StringArguementsPass(ByVal strArguements As String, Optional ByVal NewValue As String = "ByVal") As String
    On Error Resume Next
    Dim StrB() As String
    Dim aCnt As Long
    Dim aTot As Long
    Dim StrA As String
    Dim ByPos As Long
    Dim OpPos As Long
    Dim passS As String
    Dim errS As String
    errS = Chr$(34) & ", ByVal " & Chr$(34)
    StrB = Split(strArguements, ",")
    aTot = UBound(StrB)
    For aCnt = 0 To aTot
        StrA = StrB(aCnt)
        StrA = Trim$(StrA)
        StrA = Replace$(StrA, "$", " As String")
        Select Case NewValue
        Case "ByVal"
            StrA = Replace$(StrA, "ByRef ", "ByVal ")
            ByPos = InStr(StrA, "ByVal ")
        Case "ByRef"
            StrA = Replace$(StrA, "ByVal ", "ByRef ")
            ByPos = InStr(StrA, "ByRef ")
        End Select
        Select Case ByPos
        Case 0
            OpPos = InStr(StrA, "Optional")
            Select Case OpPos
            Case 0
                StrA = NewValue & " " & StrA
            Case Else
                StrA = Replace$(StrA, "Optional", "Optional " & NewValue)
            End Select
        End Select
        If InStr(1, StrA, "As String") > 0 Then
        ElseIf InStr(1, StrA, "As Variant") > 0 Then
        Else
            StrA = Replace$(StrA, "ByVal ", vbNullString)
        End If
        If InStr(1, StrA, "ByRef", vbTextCompare) > 0 Then
            StrA = Replace$(StrA, "ByRef", vbNullString)
        End If
        StrB(aCnt) = Trim$(StrA)
        Err.Clear
    Next
    passS = Trim$(Join(StrB, ", "))
    ' a trivial error that is detrimental, an optional variable defined as a comma
    If InStr(1, passS, errS) > 0 Then
        passS = Replace$(passS, errS, Chr$(34) & "," & Chr$(34))
    End If
    StringArguementsPass = passS
    Erase StrB
    Err.Clear
End Function
Public Function StringConcat(ByVal dest As String, ByVal Source As String) As String
    On Error Resume Next
    Dim sL As Long
    Dim dL As Long
    Dim NL As Long
    Dim sN As String
    Const cI As Long = 50000
    sN = dest
    sL = Len(Source)
    dL = Len(dest)
    NL = dL + sL
    Select Case NL
    Case Is >= dL
        Select Case sL
        Case Is > cI
            sN = sN & Space$(sL)
        Case Else
            sN = sN & Space$(sL + 1)
        End Select
    End Select
    Mid$(sN, dL + 1, sL) = Source
    StringConcat = Left$(sN, NL)
    Err.Clear
End Function
Public Function StringNewProcedure(ByVal strDeclaration As String, Optional ByVal PassBy As String = "ByVal") As String
    On Error Resume Next
    Dim sArguements As String
    Dim sNewArguements As String
    Dim sProcedure As String
    Dim nProcedure As String
    Select Case Right$(strDeclaration, 1)
    Case "_"
        StringNewProcedure = strDeclaration
    Case Else
        sArguements = StringArguements(strDeclaration)
        sNewArguements = StringArguementsPass(sArguements, PassBy)
        sProcedure = StringProcedure(strDeclaration)
        nProcedure = Replace$(sProcedure, "()", "(" & sNewArguements & ")")
        StringNewProcedure = nProcedure
    End Select
    Err.Clear
End Function
Public Function StringOptionalInitialize(ByVal strArguements As String) As String
    On Error Resume Next
    Dim StrB() As String
    Dim aCnt As Long
    Dim aTot As Long
    Dim StrA As String
    Dim OPos As Long
    Dim StrN() As String
    Dim EquT As Long
    StrB = Split(strArguements, ",")
    aTot = UBound(StrB)
    For aCnt = 0 To aTot
        StrA = StrB(aCnt)
        StrA = Trim$(StrA)
        OPos = InStr(StrA, "Optional ")
        Select Case OPos
        Case Is > 0
            StrN = Split(StrA, "=")
            EquT = UBound(StrN)
            Select Case EquT
            Case Is < 1
                If InStr(StrA, "As String") > 0 Then
                    StrA = StringConcat(StrA, " = vbnNullString")
                End If
                If InStr(StrA, "As Integer") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Long") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Single") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Currency") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Double") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Boolean") > 0 Then
                    StrA = StringConcat(StrA, " = False")
                End If
                If InStr(StrA, "As Byte") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "As Variant") > 0 Then
                    StrA = StringConcat(StrA, " = Nothing")
                End If
                If InStr(StrA, "%") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "!") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "#") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "@") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
                If InStr(StrA, "$") > 0 Then
                    StrA = StringConcat(StrA, " = vbNullString")
                End If
                If InStr(StrA, "&") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
            End Select
            StrB(aCnt) = StrA
        End Select
        Err.Clear
    Next
    StringOptionalInitialize = StringRemoveDelim(Join(StrB, ","), ",")
    Erase StrB
    StrA = vbNullString
    Erase StrN
    Err.Clear
End Function
Public Function StringProcedure(ByVal strDeclaration As String, Optional IncludeBrackets As Boolean = True) As String
    On Error Resume Next
    Dim fBracket As Long
    Dim sBracket As Long
    Dim sResult As String
    fBracket = InStr(1, strDeclaration, "(")
    sBracket = InStrRev(strDeclaration, ")")
    If fBracket > 0 Then
        sResult = Left$(strDeclaration, fBracket - 1)
    Else
        sResult = strDeclaration
    End If
    If IncludeBrackets = True Then
        sResult = sResult & "()"
    End If
    If sBracket > 0 Then
        sResult = sResult & Mid$(strDeclaration, sBracket + 1)
    End If
    StringProcedure = sResult
    Err.Clear
End Function
Public Function TextBoxHiliteCode(ByVal ControlName As String) As String
    On Error Resume Next
    Dim txtName As String
    Dim txtIdx As String
    Dim txtCode As String
    txtName = Split(ControlName, ",")(0)
    txtIdx = Split(ControlName, ",")(1)
    txtCode = ""
    If txtIdx = "False" Then
        txtCode = "Private Sub " & txtName & "_GotFocus()" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & vbTab & "MyADO.TextBoxHilite " & txtName & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & "End Sub"
    Else
        txtCode = "Private Sub " & txtName & "_GotFocus(Index As Integer)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & vbTab & "MyADO.TextBoxHilite " & txtName & "(Index)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & "End Sub"
    End If
    TextBoxHiliteCode = txtCode
    Err.Clear
End Function
Public Function ProperCaseCode(ByVal ControlName As String) As String
    On Error Resume Next
    Dim txtName As String
    Dim txtIdx As String
    Dim txtCode As String
    txtName = Split(ControlName, ",")(0)
    txtIdx = Split(ControlName, ",")(1)
    txtCode = ""
    If txtIdx = "False" Then
        txtCode = "Private Sub " & txtName & "_Validate(Cancel As Boolean)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & vbTab & "If Len(" & txtName & ".Text) = 0 Then Exit Sub" & vbNewLine
        txtCode = txtCode & vbTab & "If " & txtName & ".MultiLine = False Then " & txtName & ".Text = MyADO.ProperCase(" & txtName & ".Text)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & "End Sub"
    Else
        txtCode = "Private Sub " & txtName & "_Validate(Index As Integer, cancel As Boolean)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & vbTab & "If Len(" & txtName & "(Index).Text) = 0 Then Exit Sub" & vbNewLine
        txtCode = txtCode & vbTab & "If " & txtName & "(Index).MultiLine = False Then " & txtName & "(Index).Text = MyADO.ProperCase(" & txtName & "(Index).Text)" & vbNewLine
        txtCode = txtCode & vbTab & "" & vbNewLine
        txtCode = txtCode & "End Sub"
    End If
    ProperCaseCode = txtCode
    Err.Clear
End Function
Private Function TrimDeclaration(ByVal StrLine As String) As String
    On Error Resume Next
    Dim AsPos As Long
    Dim strL As String
    Dim Strr As String
    AsPos = InStr(1, StrLine, " As ", vbTextCompare)
    If AsPos > 0 Then
        strL = Trim$(Left$(StrLine, AsPos - 1))
        Strr = Trim$(Mid$(StrLine, AsPos))
        TrimDeclaration = strL & " " & Strr
    Else
        TrimDeclaration = StrLine
    End If
    Err.Clear
End Function
Public Function UnImprovedStringFunction(ByVal StrLine As String) As String
    On Error Resume Next
    Dim lngF As Long
    Dim lngN As Long
    Dim StrF As String
    Dim StrN As String
    lngF = UBound(StrFunc)
    For lngN = 1 To lngF
        StrF = StrFunc(lngN)
        StrN = StrF & "$("
        StrLine = Replace$(StrLine, StrN, StrF & "(")
        Err.Clear
    Next
    UnImprovedStringFunction = StrLine
    Err.Clear
End Function
Private Function VariableType(ByVal currLine As String) As String
    On Error Resume Next
    VariableType = Split(currLine, " ")(3)
    Err.Clear
End Function
Private Sub sWhatToProcess(Optional IncludeProc As Boolean = True, Optional ClearFirst As Boolean = False)
    On Error Resume Next
    frmProcess.optProcess(1).Value = False
    frmProcess.optProcess(2).Value = False
    frmProcess.optProcess(3).Value = False
    frmProcess.chkClear.Value = 0
    frmProcess.chkClear.Enabled = ClearFirst
    frmProcess.optProcess(1).Enabled = IncludeProc
    frmProcess.Show vbModal
    Err.Clear
End Sub
Private Sub DestroyButton(oButton As Office.CommandBarButton, oEvent As VBIDE.CommandBarEvents)
    On Error Resume Next
    oButton.Delete
    Set oButton = Nothing
    Set oEvent = Nothing
    Err.Clear
End Sub
Private Function Code_CheckProcedures(VbCp As VBIDE.VBComponent) As Long
    On Error Resume Next
    Dim procStart As Long
    Dim procEnd As Long
    Dim procValues As String
    Dim tLines As Long
    Dim cLines As Long
    Dim currLine As String
    Dim spLines() As String
    Dim procName As String
    Dim vPos As Long
    Dim vAll As String
    Dim cDead As Long
    cDead = 0
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For cLines = 1 To tLines
        If UpdateProgress(cLines, tLines, VbCp.Name & ".Checking procedure declarations") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(cLines, 1))
        If IsMethod(currLine) = True Then
            procName = StringProcedureName(currLine)
            procValues = currLine & vbNewLine
            procStart = cLines
            ReDim spLines(1 To 5) As String
            spLines(1) = VbCp.Name
            spLines(2) = procName
            spLines(3) = cLines
            spLines(4) = ""
            spLines(5) = ""
            vAll = VbCp.CodeModule.lines(1, VbCp.CodeModule.CountOfLines)
            spLines(5) = ScanVariablesUSE(vAll, procName) - 1
            If spLines(5) = "0" Then cDead = cDead + 1
            vPos = LstViewUpdate(spLines, frmVariables.lstVariables, "")
            frmVariables.lstVariables.ListItems(vPos).EnsureVisible
        ElseIf IsEndOfMethod(currLine) = True Then
            procValues = procValues & currLine
            procEnd = cLines
            spLines = LstViewGetRow(frmVariables.lstVariables, vPos)
            spLines(4) = cLines
            vPos = LstViewUpdate(spLines, frmVariables.lstVariables, vPos)
            frmVariables.lstVariables.ListItems(vPos).EnsureVisible
        Else
            procValues = procValues & currLine & vbNewLine
        End If
        Err.Clear
    Next
    CloseProgress
    Code_CheckProcedures = cDead
    Err.Clear
End Function
Private Sub ccUnCommentBlockEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    UnCommentSelectedText
    Err.Clear
End Sub

Private Sub ccUpdateRsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    UpdateRsSelectedText
    Err.Clear
End Sub

Private Sub ccValidDateMaskEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    procStart = Procedure_Start:  procEnd = Procedure_End
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    'BackupFirst
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        Code_Intellisense VbCp, "v"
    Case 3
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            Code_Intellisense VbCp, "v"
NextModule:
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub ccVariablesEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim cDead As Long
    allCode = ProjectCode
    sWhatToProcess False
    aAll = WhatToProcess
    If aAll = 4 Then Exit Sub
    'BackupFirst
    frmVariables.lstVariables.ListItems.Clear
    cDead = 0
    Screen.MousePointer = vbHourglass
    Select Case aAll
    Case 4
        Err.Clear
        Exit Sub
    Case 2
        Set VbCp = VBInst.SelectedVBComponent
        If IsModuleAppropriate(VbCp) = False Then GoTo EndProc
        VbCp.CodeModule.CodePane.Show
        cDead = Code_CheckVariablesNew(VbCp)
    Case 3
        cDead = 0
        totPanes = VBInst.ActiveVBProject.VBComponents.Count
        For cntPanes = 1 To totPanes
            Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
            If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
            VbCp.CodeModule.CodePane.Show
            cDead = cDead + Code_CheckVariablesNew(VbCp)
NextModule:
            Err.Clear
        Next
    End Select
EndProc:
    Screen.MousePointer = vbDefault
    If frmVariables.lstVariables.ListItems.Count > 0 Then
        frmVariables.Show
        frmVariables.Refresh
        LstViewAutoResize frmVariables.lstVariables
    End If
    'frmVariables.cmdRemove_Click
    MsgBox "Approximately " & frmVariables.lstVariables.ListItems.Count & " dead variables/constants have been found in your project / selected component!", vbOKOnly + vbExclamation + vbApplicationModal, "Dead Variables"
    Err.Clear
End Sub

Public Function LineContinuationsBeforePosition(VbCp As VBIDE.VBComponent, CodeLocation As Long) As Long
    On Error Resume Next
    Dim rsStr As String
    Dim rsCnt As Long
    Dim lines As Long
    
    lines = 0
    For rsCnt = 1 To CodeLocation - 1
        rsStr = Trim$(VbCp.CodeModule.lines(rsCnt, 1))
        If Right$(rsStr, 1) = "_" Then lines = lines + 1
    Next
    LineContinuationsBeforePosition = lines
    Err.Clear
End Function

Public Sub Code_Explore(VbCp As VBIDE.VBComponent)
    On Error Resume Next
    Dim memTot As Long
    Dim memCnt As Long
    Dim spLine(1 To 10) As String
'Dim cCode As String
'Dim spPos As Long
    Dim mNodes(4) As MSComctlLib.Node
    Dim strKey1 As String
    Dim strKey2 As String
    
    memTot = VbCp.CodeModule.Members.Count
    'Call InitProgress(memTot)
    For memCnt = 1 To memTot
        'DoEvents
        'If UpdateProgress(memCnt, memTot, VbCp.Name & ".Loading members...") = False Then Exit Sub
        spLine(1) = ComponentType(VbCp)
        spLine(2) = VbCp.Name
        spLine(4) = MemberType(VbCp.CodeModule.Members(memCnt))
        spLine(3) = MemberScope(VbCp.CodeModule.Members(memCnt))
        spLine(5) = VbCp.CodeModule.Members(memCnt).Name
        ' the code location returns the location irrespective of line continuation
        ' we need to find out how many line continuations before the current location
        spLine(6) = VbCp.CodeModule.Members(memCnt).CodeLocation
        spLine(6) = Val(spLine(6)) - LineContinuationsBeforePosition(VbCp, Val(spLine(6)))
        spLine(7) = "0"
        Select Case spLine(4)
        Case "Constant", "Variable"
            spLine(8) = VbCp.CodeModule.lines(Val(spLine(6)), 1)
            spLine(9) = spLine(6)
        Case Else
            spLine(8) = Procedure_Code(VbCp, spLine(5), Val(spLine(6)))
            spLine(9) = Val(spLine(6)) + Val(MvCount(spLine(8), vbNewLine))
        End Select
        spLine(10) = VbCp.CodeModule.Members(memCnt).Category
        
        ' add component type
        Set mNodes(1) = frmProject.treeProject.Nodes.Add(, , ComponentType(VbCp), ComponentType(VbCp), "close", "open")
        Set mNodes(2) = frmProject.treeProject.Nodes.Add(spLine(1), tvwChild, spLine(2), spLine(2), "close", "open")
        mNodes(2).Tag = "Description: " & VbCp.Description & vbNewLine & "Members: " & memTot & vbNewLine & "Lines: " & VbCp.CodeModule.CountOfLines & vbNewLine & _
        "Declaration Lines: " & VbCp.CodeModule.CountOfDeclarationLines & vbNewLine & "Start: " & spLine(6)
        
        ' name and type
        strKey1 = spLine(2) & "-" & spLine(4)
        strKey1 = Replace$(strKey1, " ", "")
        strKey1 = Replace$(strKey1, "/", "")
        Set mNodes(3) = frmProject.treeProject.Nodes.Add(spLine(2), tvwChild, strKey1, spLine(4), "close", "open")
        
        ' add member
        strKey2 = strKey1 & "-" & spLine(5)
        Set mNodes(4) = frmProject.treeProject.Nodes.Add(strKey1, tvwChild, strKey2, spLine(5), "close", "open")
        mNodes(4).Tag = "Member Name: " & spLine(5) & vbNewLine & "Member Type: " & spLine(4) & vbNewLine & "Member Scope: " & spLine(3) & vbNewLine & "Member Category: " & spLine(10) & vbNewLine & _
        "Start: " & spLine(6) & vbNewLine & "End: " & spLine(9) & vbNewLine & "Usage: " & spLine(7) & vbNewLine & "Code: " & vbNewLine & spLine(8)
        'Select Case spLine(3)
        'Case "Private"
        '    cCode = VbCp.CodeModule.Lines(1, VbCp.CodeModule.CountOfLines)
        '    spLine(7) = ScanVariablesUSE(cCode, spLine(5)) - 1
        'End Select
        'spPos = LstViewUpdate(spLine, frmProject.lstProject, "")
        'frmProject.lstProject.ListItems(spPos).EnsureVisible
        Err.Clear
    Next
    'CloseProgress
    Err.Clear
End Sub
Public Sub Code_CheckUsage()
    On Error Resume Next
    Dim lTot As Long
    Dim lCnt As Long
    Dim lLines() As String
    Dim cName As String
    Dim vName As String
    Dim vScop As String
    Dim vType As String
    'lTot = frmProject.lstProject.ListItems.Count
    If InitProgress(lTot) = False Then Exit Sub
    For lCnt = 1 To lTot
        DoEvents
        'lLines = LstViewGetRow(frmProject.lstProject, lCnt)
        cName = lLines(2)   ' component
        vScop = lLines(3)   ' member scope
        vType = lLines(4)   ' member type
        vName = lLines(5)   ' member name
        If UpdateProgress(lCnt, lTot, "Checking member usage for: " & cName & "." & vName) = False Then Exit For
        ' set the component
        Set VbCp = VBInst.ActiveVBProject.VBComponents(cName)
        If TypeName(VbCp) = "Nothing" Then GoTo NextLine
        Select Case vScop
        Case "Friend", "Public"
            lLines(7) = ScanVariablesUSE(allCode, vName)
            'lPos = LstViewUpdate(lLines, frmProject.lstProject, CStr(lCnt))
        End Select
        'frmProject.lstProject.ListItems(lCnt).EnsureVisible
NextLine:
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Public Function Procedure_End() As Long
    On Error Resume Next
    Dim oCodePane As VBIDE.CodePane
    Dim oCodeMod As VBIDE.CodeModule
    Dim iCurrentLine As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim StartLine As Long
    Dim sProcName As String
    Dim eProcKind As VBIDE.vbext_ProcKind
    Set oCodePane = VBInst.ActiveCodePane
    Set oCodeMod = oCodePane.CodeModule
    'returns the current line that cursor is on
    oCodePane.GetSelection iCurrentLine, b, c, d
    'returns the procedure name and the ProcKind (procedure, property get, etc)
    sProcName = oCodeMod.ProcOfLine(iCurrentLine, eProcKind)
    If Len(sProcName) > 0 Then
        ' what is the start and end line numbers of the procedure
        StartLine = oCodeMod.ProcStartLine(sProcName, eProcKind)
        Procedure_End = oCodeMod.ProcCountLines(sProcName, eProcKind) + StartLine
    Else
        Procedure_End = -1
    End If
    Set oCodePane = Nothing
    Set oCodeMod = Nothing
    Err.Clear
End Function
Public Function Procedure_Code(VbCp As VBIDE.VBComponent, sProcName As String, sProcStart As Long) As String
    On Error Resume Next
    Dim StartLine As Long
    Dim EndLine As Long
    Dim eProcKind As VBIDE.vbext_ProcKind
    Dim xLine As String
    Dim fLine As String
    'just check if this is not api
    fLine = VbCp.CodeModule.lines(sProcStart, 1)
    If InStr(1, fLine, "Declare ", vbTextCompare) > 0 Then
        EndLine = VbCp.CodeModule.ProcCountLines(sProcName, eProcKind) + StartLine
        Procedure_Code = VbCp.CodeModule.lines(StartLine, EndLine - StartLine)
    Else
        'returns the procedure name and the ProcKind (procedure, property get, etc)
        xLine = VbCp.CodeModule.ProcOfLine(sProcStart, eProcKind)
        If Len(xLine) > 0 Then
            ' what is the start and end line numbers of the procedure
            StartLine = VbCp.CodeModule.ProcStartLine(sProcName, eProcKind)
            EndLine = VbCp.CodeModule.ProcCountLines(sProcName, eProcKind) + StartLine
            Procedure_Code = VbCp.CodeModule.lines(StartLine, EndLine - StartLine)
        Else
            Procedure_Code = ""
        End If
    End If
    Err.Clear
End Function
Public Function Procedure_Start() As Long
    On Error Resume Next
    Dim oCodePane As VBIDE.CodePane
    Dim oCodeMod As VBIDE.CodeModule
    Dim iCurrentLine As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim sProcName As String
    Dim eProcKind As VBIDE.vbext_ProcKind
    Set oCodePane = VBInst.ActiveCodePane
    Set oCodeMod = oCodePane.CodeModule
    'returns the current line that cursor is on
    oCodePane.GetSelection iCurrentLine, b, c, d
    'returns the procedure name and the ProcKind (procedure, property get, etc)
    sProcName = oCodeMod.ProcOfLine(iCurrentLine, eProcKind)
    If Len(sProcName) > 0 Then
        'what is the start and end line numbers of the procedure
        Procedure_Start = oCodeMod.ProcStartLine(sProcName, eProcKind)
    Else
        Procedure_Start = -1
    End If
    Set oCodePane = Nothing
    Set oCodeMod = Nothing
    Err.Clear
End Function
Private Function Code_CheckVariablesNew(VbCp As VBIDE.VBComponent) As Long
    On Error Resume Next
    Dim tLines As Long
    Dim cLines As Long
    Dim spLines(1 To 7) As String
    Dim cDead As Long
    Dim sComponent As String
    Dim sMemberName As String
    Dim sMemberType As String
    Dim sMemberScope As String
    Dim sMemberLocation As String
    Dim sMemberUsage As String
    Dim sMemberCode As String
    Dim sMemberPos As Long
    Dim sMemberEnd As String
    'Dim compCode As String
    Dim procCnt As Long
    Dim sMemberLine As String
    cDead = 0
    tLines = VbCp.CodeModule.Members.Count
    Call InitProgress(tLines)
    For cLines = 1 To tLines
        If UpdateProgress(cLines, tLines, VbCp.Name & ".Checking variable declarations") = False Then
            Exit For
        End If
        sComponent = VbCp.Name
        sMemberName = VbCp.CodeModule.Members(cLines).Name
        sMemberType = MemberType(VbCp.CodeModule.Members(cLines))
        sMemberScope = MemberScope(VbCp.CodeModule.Members(cLines))
        sMemberLocation = VbCp.CodeModule.Members(cLines).CodeLocation
        sMemberUsage = "0"
        sMemberCode = Procedure_Code(VbCp, sMemberName, Val(sMemberLocation))
        sMemberEnd = Val(sMemberLocation) + Val(MvCount(sMemberCode, vbNewLine))
        spLines(1) = sComponent
        spLines(2) = sMemberName
        spLines(3) = sMemberType
        spLines(4) = sMemberScope
        spLines(5) = sMemberLocation
        spLines(6) = sMemberEnd
        spLines(7) = sMemberUsage
        If Val(spLines(6)) < Val(spLines(5)) Then spLines(6) = spLines(5)
        'Select Case spLines(4)
        'Case "Private"
        '    compCode = VbCp.CodeModule.Lines(1, VbCp.CodeModule.CountOfLines)
        '    spLines(7) = ScanVariablesUSE(compCode, spLines(2)) - 1
        'Case Else
        '    spLines(7) = ScanVariablesUSE(allCode, spLines(2)) - 1
        'End Select
        'If spLines(7) = "0" Then cDead = cDead + 1
        'sMemberPos = LstViewUpdate(spLines, frmVariables.lstVariables, "")
        'frmVariables.lstVariables.ListItems(sMemberPos).Tag = sMemberCode
        'frmVariables.lstVariables.ListItems(sMemberPos).EnsureVisible
        ' check variables declared in procedures
        If Procedure_Variables(sMemberCode).Count > 0 Then
            For procCnt = Val(sMemberLocation) To Val(sMemberEnd)
                sMemberLine = Trim$(VbCp.CodeModule.lines(procCnt, 1))
                If IsVariable(sMemberLine) = True Then
                    spLines(1) = sComponent & "." & sMemberName
                    spLines(2) = VariableName(sMemberLine)
                    spLines(3) = IIf((InStr(1, sMemberLine, "Dim ")), "Variable", "Constant")
                    spLines(4) = "Private"
                    spLines(5) = procCnt
                    spLines(6) = procCnt
                    spLines(7) = ScanVariablesUSE(sMemberCode, spLines(2)) - 1
                    If spLines(7) = "0" Then cDead = cDead + 1
                    If Val(spLines(6)) < Val(spLines(5)) Then spLines(6) = spLines(5)
                    If spLines(7) = "0" Then
                        sMemberPos = LstViewUpdate(spLines, frmVariables.lstVariables, "")
                        frmVariables.lstVariables.ListItems(sMemberPos).Tag = sMemberLine
                        frmVariables.lstVariables.ListItems(sMemberPos).EnsureVisible
                    End If
                End If
                Err.Clear
            Next
        End If
        Err.Clear
    Next
    Code_CheckVariablesNew = cDead
    CloseProgress
    Err.Clear
End Function
Private Sub Code_CheckStartWithComments(VbCp As VBIDE.VBComponent, Optional ByVal StartText As String = "'")
    On Error Resume Next
    Dim tLines As Long
    Dim cLines As Long
    Dim spLines(1 To 3) As String
    Dim vPos As Long
    Dim sComponent As String
    Dim sMemberLine As String
    Dim lLen As Long
    lLen = Len(StartText)
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For cLines = 1 To tLines
        If UpdateProgress(cLines, tLines, VbCp.Name & ".Checking lines starting with comment...") = False Then
            Exit For
        End If
        sComponent = VbCp.Name
        sMemberLine = Trim$(VbCp.CodeModule.lines(cLines, 1))
        If Left$(sMemberLine, lLen) = StartText Then
            spLines(1) = sComponent
            spLines(2) = sMemberLine
            spLines(3) = cLines
            vPos = LstViewUpdate(spLines, frmLinesComments.lstVariables, "")
            frmLinesComments.lstVariables.ListItems(vPos).EnsureVisible
        End If
        Err.Clear
    Next
    LstViewAutoResize frmLinesComments.lstVariables
    CloseProgress
    Err.Clear
End Sub
Public Function GetSelectedText() As String
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Function
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Function
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = Mid$(cmo.lines(StartLine, 1), StartCol, EndCol - StartCol)
    Else
        ' the selection spans multiple lines of code
        ' first, get the selection of the first line
        CodeText = Mid$(cmo.lines(StartLine, 1), StartCol) & vbCrLf
        ' then get the lines in the middle, that are fully highlighted
        If StartLine + 1 < EndLine Then
            CodeText = CodeText & cmo.lines(StartLine + 1, EndLine - StartLine - 1)
        End If
        ' finally, get the highlighted portion of the last line
        CodeText = CodeText & Left$(cmo.lines(EndLine, 1), EndCol - 1)
    End If
    GetSelectedText = CodeText
    Err.Clear
End Function
Public Sub ConvertSelectedText(Optional Conversion As VbStrConv = vbUpperCase)
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = cmo.lines(StartLine, 1)
        Mid$(CodeText, StartCol, EndCol - StartCol) = StrConv(Mid$(CodeText, StartCol, EndCol - StartCol), Conversion)
        cmo.ReplaceLine StartLine, CodeText
    Else
        ' the selection spans multiple lines of code
        ' first, convert the highlighted text on the first line
        CodeText = cmo.lines(StartLine, 1)
        Mid$(CodeText, StartCol, Len(CodeText) + 1 - StartCol) = StrConv(Mid$(CodeText, StartCol, Len(CodeText) + 1 - StartCol), Conversion)
        cmo.ReplaceLine StartLine, CodeText
        ' then convert the lines in the middle, that are
        ' fully highlighted
        For i = StartLine + 1 To EndLine - 1
            CodeText = cmo.lines(i, 1)
            CodeText = StrConv(CodeText, Conversion)
            cmo.ReplaceLine i, CodeText
            Err.Clear
        Next
        ' finally, convert the highlighted portion of the last line
        CodeText = cmo.lines(EndLine, 1)
        Mid$(CodeText, 1, EndCol - 1) = StrConv(Mid$(CodeText, 1, EndCol - 1), Conversion)
        cmo.ReplaceLine EndLine, CodeText
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    cpa.SetSelection StartLine, StartCol, EndLine, EndCol
    Err.Clear
End Sub
Private Function SwappedLine(ByVal CodeText As String)
    On Error Resume Next
    Dim xNew() As String
    Dim Strr As String
    Dim SPos As Long
    Dim EPos As Long
    Dim dSiz As Long
    Dim SMid As String
    SwappedLine = CodeText
    If InStr(1, CodeText, "=") = 0 Then Exit Function
    StringParse xNew, CodeText, "="
    ReDim Preserve xNew(2)
    xNew(1) = StringClean(xNew(1))
    xNew(2) = StringClean(xNew(2))
    CodeText = Trim$(xNew(2) & " = " & xNew(1))
    Strr = CodeText
    SPos = InStr(Strr, "'")
    EPos = InStr(Strr, "=")
    If SPos > 0 Then
        Select Case SPos
        Case Is > EPos
            dSiz = SPos - EPos
            SMid = Mid$(Strr, SPos, dSiz)
            Mid$(Strr, SPos, dSiz) = ""
            CodeText = Strr
        Case Else
            Strr = Left$(CodeText, SPos - 1) & Mid$(CodeText, EPos) & String$(4, " ") & Mid$(CodeText, SPos, (EPos - SPos))
            CodeText = Strr
        End Select
    End If
    SwappedLine = CodeText
    Err.Clear
End Function
Public Sub SwapSelectedText()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = cmo.lines(StartLine, 1)
        CodeText = SwappedLine(CodeText)
        cmo.ReplaceLine StartLine, CodeText
    Else
        ' the selection spans multiple lines of code
        ' first, convert the highlighted text on the first line
        CodeText = cmo.lines(StartLine, 1)
        CodeText = SwappedLine(CodeText)
        cmo.ReplaceLine StartLine, CodeText
        ' then convert the lines in the middle, that are
        ' fully highlighted
        For i = StartLine + 1 To EndLine - 1
            CodeText = cmo.lines(i, 1)
            CodeText = SwappedLine(CodeText)
            cmo.ReplaceLine i, CodeText
            Err.Clear
        Next
        ' finally, convert the highlighted portion of the last line
        CodeText = cmo.lines(EndLine, 1)
        CodeText = SwappedLine(CodeText)
        cmo.ReplaceLine EndLine, CodeText
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub
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
Public Function ReadFormControls() As Long
    On Error Resume Next
    Dim ctl As VBIDE.VBControl
    Set formControls = New Collection
    Dim sTmp As String
    'load the component
    Set VbCp = VBInst.SelectedVBComponent
    'check to see if we have a valid component
    If TypeName(VbCp) = "Nothing" Then
        MsgBox "Please select a component first!", vbOKOnly + vbExclamation + vbApplicationModal, "Component Error"
        Err.Clear
        Exit Function
    End If
    'make sure the active component is a form, user control or property page
    If (VbCp.Type <> vbext_ct_VBForm) And (VbCp.Type <> vbext_ct_UserControl) And (VbCp.Type <> vbext_ct_DocObject) And (VbCp.Type <> vbext_ct_PropPage) Then
        MsgBox "The selected component is not valid, you need to be working on a form, a user control or a property page!", vbOKOnly + vbExclamation + vbApplicationModal, "Component Error"
        Err.Clear
        Exit Function
    End If
    Set mcolCtls = VbCp.Designer.VBControls
    For Each ctl In VbCp.Designer.VBControls
        If IsPropertyAvailable(ctl.Properties!Font) = False Then GoTo SkipIt
        sTmp = ControlName(ctl)
        formControls.Add sTmp
SkipIt:
        Err.Clear
    Next
    ReadFormControls = formControls.Count
    Err.Clear
End Function
Private Function ControlName(ctl As VBIDE.VBControl) As String
    On Error Resume Next
    Dim sTmp As String
    Dim i As Integer
    If IsPropertyAvailable(ctl.Properties!Name) = True Then sTmp = ctl.Properties!Name
    'will be null if there isn't one
    i = ctl.Properties!Index
    If i >= 0 Then
        sTmp = sTmp & "(" & i & ")"
    End If
    ControlName = sTmp
    Err.Clear
End Function
Public Sub ChangeFonts(sFontName As String, iFontSize As Variant)
    On Error Resume Next
    Dim ctl As VBIDE.VBControl
    Set VbCp = VBInst.SelectedVBComponent
    For Each ctl In VbCp.Designer.VBControls
        MsgBox ctl.Properties("FontName") & vbCr & ctl.Properties("FontSize")
        'If IsPropertyAvailable(ctl.Properties!FontName) = True Then
        '    ctl.Properties!FontName = sFontName
        'End If
        'If IsPropertyAvailable(ctl.Properties!FontSize) = True Then
        '    ctl.Properties!FontName = Val(iFontSize)
        'End If
        Err.Clear
    Next
SkipIt:
    Err.Clear
End Sub
Public Sub MySQLSelectedText()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = cmo.lines(StartLine, 1)
        CodeText = MySQLTo(CodeText)
        cmo.ReplaceLine StartLine, CodeText
    Else
        ' the selection spans multiple lines of code
        ' first, convert the highlighted text on the first line
        CodeText = cmo.lines(StartLine, 1)
        CodeText = MySQLTo(CodeText)
        cmo.ReplaceLine StartLine, CodeText
        ' then convert the lines in the middle, that are
        ' fully highlighted
        For i = StartLine + 1 To EndLine - 1
            CodeText = cmo.lines(i, 1)
            CodeText = MySQLTo(CodeText)
            cmo.ReplaceLine i, CodeText
            Err.Clear
        Next
        ' finally, convert the highlighted portion of the last line
        CodeText = cmo.lines(EndLine, 1)
        CodeText = MySQLTo(CodeText)
        cmo.ReplaceLine EndLine, CodeText
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub
Public Sub MySQLSelectedText1()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = cmo.lines(StartLine, 1)
        CodeText = MySQLTo1(CodeText)
        cmo.ReplaceLine StartLine, CodeText
    Else
        ' the selection spans multiple lines of code
        ' first, convert the highlighted text on the first line
        CodeText = cmo.lines(StartLine, 1)
        CodeText = MySQLTo1(CodeText)
        cmo.ReplaceLine StartLine, CodeText
        ' then convert the lines in the middle, that are
        ' fully highlighted
        For i = StartLine + 1 To EndLine - 1
            CodeText = cmo.lines(i, 1)
            CodeText = MySQLTo1(CodeText)
            cmo.ReplaceLine i, CodeText
            Err.Clear
        Next
        ' finally, convert the highlighted portion of the last line
        CodeText = cmo.lines(EndLine, 1)
        CodeText = MySQLTo1(CodeText)
        cmo.ReplaceLine EndLine, CodeText
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub
Private Function MySQLTo(ByVal CodeText As String) As String
    On Error Resume Next
    Dim xNew() As String
    Dim Strr As String
    Dim SPos As Long
    Dim EPos As Long
    Dim dSiz As Long
    Dim SMid As String
    Dim sTable As String
    Dim sFld As String
    MySQLTo = CodeText
    ' the line must have an exclamation
    If InStr(1, CodeText, "!") = 0 Then Exit Function
    StringParse xNew, CodeText, "="
    ReDim Preserve xNew(2)
    xNew(1) = StringClean(xNew(1))
    xNew(2) = StringClean(xNew(2))
    sTable = MvField(xNew(1), 1, "!")
    sFld = MvField(xNew(1), 2, "!")
    CodeText = sTable & ".Fields(" & Quote & sFld & Quote & ") = " & xNew(2)
    Strr = CodeText
    SPos = InStr(Strr, "'")
    EPos = InStr(Strr, "=")
    If SPos > 0 Then
        Select Case SPos
        Case Is > EPos
            dSiz = SPos - EPos
            SMid = Mid$(Strr, SPos, dSiz)
            Mid$(Strr, SPos, dSiz) = ""
            CodeText = Strr
        Case Else
            Strr = Left$(CodeText, SPos - 1) & Mid$(CodeText, EPos) & String$(4, " ") & Mid$(CodeText, SPos, (EPos - SPos))
            CodeText = Strr
        End Select
    End If
    MySQLTo = CodeText
    Err.Clear
End Function
Private Function MySQLTo1(ByVal CodeText As String) As String
    On Error Resume Next
    Dim xNew() As String
    Dim Strr As String
    Dim SPos As Long
    Dim EPos As Long
    Dim dSiz As Long
    Dim SMid As String
    Dim sTable As String
    Dim sFld As String
    MySQLTo1 = CodeText
    ' the line must have an exclamation
    If InStr(1, CodeText, "!") = 0 Then Exit Function
    StringParse xNew, CodeText, "="
    ReDim Preserve xNew(2)
    xNew(1) = StringClean(xNew(1))
    xNew(2) = StringClean(xNew(2))
    sTable = MvField(xNew(2), 1, "!")
    sFld = MvField(xNew(2), 2, "!")
    CodeText = sTable & ".Fields(" & Quote & sFld & Quote & ") = " & xNew(2)
    Strr = CodeText
    SPos = InStr(Strr, "'")
    EPos = InStr(Strr, "=")
    If SPos > 0 Then
        Select Case SPos
        Case Is > EPos
            dSiz = SPos - EPos
            SMid = Mid$(Strr, SPos, dSiz)
            Mid$(Strr, SPos, dSiz) = ""
            CodeText = Strr
        Case Else
            Strr = Left$(CodeText, SPos - 1) & Mid$(CodeText, EPos) & String$(4, " ") & Mid$(CodeText, SPos, (EPos - SPos))
            CodeText = Strr
        End Select
    End If
    MySQLTo1 = CodeText
    Err.Clear
End Function
Private Sub Code_RemoveBlankLines(VbCp As VBIDE.VBComponent, StartLine As Long, EndLine As Long)
    On Error Resume Next
    Dim totLines As Long
    Dim cntLines As Long
    Dim curLine As String
    'Dim StrSearch As String
    'Dim varText As Variant
    'Dim compCode As String
    totLines = VbCp.CodeModule.CountOfLines
    If EndLine > 0 Then totLines = EndLine
    Call InitProgress(totLines)
    For cntLines = totLines To StartLine Step -1
        curLine = VbCp.CodeModule.lines(cntLines, 1)
        curLine = Trim$(curLine)
        If UpdateProgress(cntLines, totLines, VbCp.Name & ".Removing blank Lines") = False Then
            Exit For
        End If
        If Len(curLine) = 0 Then VbCp.CodeModule.DeleteLines cntLines
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Public Sub CommentSelectedText()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = "'" & Trim$(cmo.lines(StartLine, 1))
        cmo.ReplaceLine StartLine, CodeText
    Else
        ' the selection spans multiple lines of code
        For i = StartLine To EndLine
            CodeText = "'" & Trim$(cmo.lines(i, 1))
            cmo.ReplaceLine i, CodeText
            Err.Clear
        Next
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub
Public Sub UnCommentSelectedText()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        CodeText = Trim$(cmo.lines(StartLine, 1))
        If Left$(CodeText, 1) = "'" Then
            CodeText = Mid$(CodeText, 2)
            cmo.ReplaceLine StartLine, CodeText
        End If
    Else
        ' the selection spans multiple lines of code
        For i = StartLine To EndLine
            CodeText = Trim$(cmo.lines(i, 1))
            If Left$(CodeText, 1) = "'" Then
                CodeText = Mid$(CodeText, 2)
                cmo.ReplaceLine i, CodeText
            End If
            Err.Clear
        Next
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub

'Private Sub LstBoxSort(ccFontName As Office.CommandBarControl)
'    On Error Resume Next
'    Dim rsCnt As Long
'    Dim rsTot As Long
'    Dim rsStr As String
'    Dim spFonts() As String
'
'    rsStr = ""
'    rsTot = ccFontName.ListItems.Count - 1
'    For rsCnt = 0 To rsTot
'        rsStr = rsStr & ccFontName.list(rsCnt) & ";"
'    Next
'    rsStr = StringRemoveDelim(rsStr, ";")
'    rsStr = MvSort_String(rsStr, ";")
'
'    spFonts = Split(rsStr, ";")
'    rsTot = UBound(spFonts)
'    ccFontName.Clear
'
'    For rsCnt = 0 To rsTot
'        ccFontName.AddItem spFonts(rsCnt)
'    Next
'    Err.Clear
'End Sub
'
Private Function MvSort_String(vString As String, Optional Delim As String = vbNullString) As String
    On Error Resume Next
    Dim lLoop1 As Long
    Dim lHold As Long
    Dim lHValue As String
    Dim lTemp As String
    Dim Varray() As String
    Dim lLoop1_Tot As Long
    Dim lLoop1_Cnt As Long
    Varray = Split(vString, Delim)
    lHValue = LBound(Varray)
    Do
        lHValue = 3 * lHValue + 1
    Loop Until lHValue > UBound(Varray)
    Do
        lHValue = lHValue / 3
        lLoop1_Tot = UBound(Varray)
        lLoop1_Cnt = lHValue + LBound(Varray)
        For lLoop1 = lLoop1_Cnt To lLoop1_Tot
            lTemp = Varray(lLoop1)
            lHold = Val(lLoop1)
            Do While Varray(lHold - lHValue) > lTemp
                Varray(lHold) = Varray(lHold - lHValue)
                lHold = lHold - lHValue
                If lHold < lHValue Then
                    Exit Do
                End If
            Loop
            Varray(lHold) = lTemp
            Err.Clear
        Next
    Loop Until lHValue = LBound(Varray)
    MvSort_String = MvFromArray(Varray, Delim, 0)
    Err.Clear
End Function

Private Function MvFromArray(Varray() As String, Optional ByVal Delim As String = "", Optional StartingAt As Long = 1, Optional TrimItem As Boolean = True, Optional Reverse As Boolean = False) As String
    On Error Resume Next
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    If IsArray(Varray) = False Then Exit Function
    totArray = UBound(Varray)
    For i = StartingAt To totArray
        strL = Varray(i)
        If TrimItem = True Then
            strL = Trim$(strL)
        End If
        If i = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delim
        End If
        Err.Clear
    Next
    If Reverse = True Then
        For i = totArray To StartingAt Step -1
            strL = Varray(i)
            If TrimItem = True Then
                strL = Trim$(strL)
            End If
            If i = totArray Then
                BldStr = BldStr & strL
            Else
                BldStr = BldStr & strL & Delim
            End If
            Err.Clear
        Next
    End If
    MvFromArray = BldStr
    Err.Clear
End Function


Public Sub UpdateRsSelectedText()
    On Error Resume Next
    Dim StartLine As Long
    Dim StartCol As Long
    Dim EndLine As Long
    Dim EndCol As Long
    'Dim CodeText As String
    Dim cpa As VBIDE.CodePane
    Dim cmo As VBIDE.CodeModule
    Dim i As Long
    Dim p1 As String
    Dim p2 As String
    Dim sL As String
    ' get a reference to the active code window and the
    ' underlying module
    ' exit if no one is available
    Set cpa = VBInst.ActiveCodePane
    Set cmo = cpa.CodeModule
    If Err Then Exit Sub
    ' get the current selection coordinates
    cpa.GetSelection StartLine, StartCol, EndLine, EndCol
    ' exit if no text is highlighted
    If StartLine = EndLine And StartCol = EndCol Then Exit Sub
    ' get the code text
    If StartLine = EndLine Then
        ' only one line is partially or fully highlighted
        sL = Trim$(cmo.lines(StartLine, 1))
        ' this should be xxx.Update
        If InStr(1, sL, ".") > 0 Then
            p1 = Split(sL, ".")(0)
            p2 = Split(sL, ".")(1)
            If p2 = "Update" Then
                cmo.ReplaceLine StartLine, "MyAdo.UpdateRs " & p1
            End If
        End If
    Else
        ' the selection spans multiple lines of code
        For i = StartLine To EndLine
            sL = Trim$(cmo.lines(i, 1))
            ' this should be xxx.Update
            If InStr(1, sL, ".") > 0 Then
                p1 = Split(sL, ".")(0)
                p2 = Split(sL, ".")(1)
                If p2 = "Update" Then
                    cmo.ReplaceLine i, "MyAdo.UpdateRs " & p1
                End If
            End If
        Next
    End If
    ' after replacing code we must restore the old selection
    ' this seems to be a side-effect of the ReplaceLine method
    Err.Clear
End Sub

Private Sub ccWPFEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim xSel As String
    Dim m As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    'Dim z As String
    VBInst.ActiveCodePane.GetSelection m, n, X, Y
    xSel = VBInst.ActiveCodePane.CodeModule.lines(m, X)
       
    frmConvert.Show
    frmConvert.ZOrder 0
    frmConvert.txtSource.Text = xSel
    Err.Clear
End Sub
