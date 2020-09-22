VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoolCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect
Private vbCp As VBIDE.CodePane
Private StrFunc(19) As String
Private Const Quote As String = """"
Private Archive As Collection
Private ArchiveFilename As String
Private Dbase As String
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Private Declare Function AddFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean

Private ccRemoveBlanks As Office.CommandBarButton
Private ccTrimLines As Office.CommandBarButton
   
Private WithEvents ccRemoveBlanksEvents   As CommandBarEvents
Attribute ccRemoveBlanksEvents.VB_VarHelpID = -1
Private WithEvents ccTrimLinesEvents   As CommandBarEvents
Attribute ccTrimLinesEvents.VB_VarHelpID = -1

Private Function FixDeclaration(ByVal StrLine As String) As String
    On Error Resume Next
    StrLine = Replace$(StrLine, "%", " As Integer")
    StrLine = Replace$(StrLine, "&", " As Long")
    StrLine = Replace$(StrLine, "!", " As Single")
    StrLine = Replace$(StrLine, "#", " As Double")
    StrLine = Replace$(StrLine, "@", " As Currency")
    StrLine = Replace$(StrLine, "$", " As String")
    FixDeclaration = StrLine
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
    Next
    Err.Clear
End Function
Private Function win_Function_Exist(sModule As String, sFunction As String) As Boolean
    On Error Resume Next
    Dim hHandle As Long
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        If GetProcAddress(hHandle, sFunction) = 0 Then
            win_Function_Exist = False
        Else
            win_Function_Exist = True
        End If
        FreeLibrary hHandle
    Else
        If GetProcAddress(hHandle, sFunction) <> 0 Then
            win_Function_Exist = True
        End If
    End If
    Err.Clear
End Function
Private Sub Form_Load()
    On Error Resume Next
    IniStrFunctions
    Dbase = ExactPath(App.Path) & "\CoolCode.mdb"
    Dao_CreateDatabase Dbase, True
    Err.Clear
End Sub
Private Sub OKButton_Click()
    On Error Resume Next
    Dim totPanes As Long
    Dim cntPanes As Long
    Dim bPath As String
    If Connect.VBInstance.ActiveVBProject Is Nothing Then
        MsgBox "The addin could not see any VB Project - you must open a project first.", vbExclamation, "Open Project"
        Err.Clear
        Exit Sub
    End If
    If optModule.Value = False Then
        If optProject.Value = False Then
            MsgBox "Please select a module or project to process!", vbOKOnly + vbExclamation, "Processing Scope"
            Err.Clear
            Exit Sub
        End If
    End If
    cntPanes = MsgBox("Do you want to backup the project to a compressed file first?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Backup")
    If cntPanes = vbYes Then
        Screen.MousePointer = vbHourglass
        Caption = "Cool Code: Compressing Project Folder"
        bPath = StringGetFileToken(Connect.VBInstance.ActiveVBProject.Filename, "p")
        bPath = StringGetFileToken(bPath, "p")
        ArchiveFilename = Connect.VBInstance.ActiveVBProject.Filename
        ArchiveFilename = ExactPath(bPath) & "\" & StringGetFileToken(ArchiveFilename, "fo") & " " & Format$(Now(), "ddmmyyyy hhmmss") & ".zip"
        Backup_ToCompressedFile ArchiveFilename, ExactPath(StringGetFileToken(Connect.VBInstance.ActiveVBProject.Filename, "p"))
        Screen.MousePointer = vbHourglass
        Caption = "Cool Code"
    End If
    If optModule.Value = True Then
        Set vbCp = Connect.VBInstance.ActiveCodePane
        GoSub ProcessCode
    End If
    If optProject.Value = True Then
        totPanes = Connect.VBInstance.CodePanes.Count
        progBar1.Max = totPanes
        progBar1.Min = 0
        progBar1.Value = 0
        For cntPanes = 1 To totPanes
            progBar1.Value = cntPanes
            Set vbCp = Connect.VBInstance.CodePanes.Item(cntPanes)
            GoSub ProcessCode
        Next
        progBar1.Value = 0
    End If
    Caption = "Cool Code"
    Screen.MousePointer = vbDefault
    Err.Clear
    Exit Sub
ProcessCode:
    If chkRemBlanks.Value = 1 Then
        Code_RemoveLines vbCp, ""
    End If
    If chkTrim.Value = 1 Then
        Code_TrimLines vbCp
    End If
    If chkRemLineCont.Value = 1 Then
        Code_RemoveLineContinuation vbCp
    End If
    If chkBreakCode.Value = 1 Then
        Code_BreakMultiDeclarations vbCp
    End If
    If chkStringFunctions.Value = 1 Then
        Code_ConvertStringFunctions vbCp
    End If
    If chkUnconvertStringFunc.Value = 1 Then
        Code_ConvertStringFunctions vbCp, False
    End If
    If chkMoveComments.Value = 1 Then
        Code_MoveComment vbCp
    End If
    If chkNextEnd.Value = 1 Then
        Code_SpeedNextLoopEnd vbCp
    End If
    If chkDoEvents.Value = 1 Then
        Code_InsertDoEvents vbCp
    End If
    If chkRemLines.Value = 1 Then
        If chkExact.Value = 1 Then
            Code_RemoveLines vbCp, cboRemove, True
        Else
            Code_RemoveLines vbCp, cboRemove.Text, False
        End If
    End If
    If chkPassByVal.Value = 1 Then
        Code_PassBy vbCp
    End If
    If chkPassByRef.Value = 1 Then
        Code_PassBy vbCp, "ByRef"
    End If
    If chkInitialize.Value = 1 Then
        Code_InitializeOptionalVariables vbCp
    End If
    If chkRemEmpty.Value = 1 Then
        Code_RemoveEmpty vbCp
    End If
    If chkNormalizeIf.Value = 1 Then
        Code_NormalizeIf vbCp
    End If
    If chkFixComplex.Value = 1 Then
        Code_BreakMultiDeclarations vbCp
        Code_FixComplexLoops vbCp
    End If
    If chkMoveDeclarations.Value = 1 Then
        Code_BreakMultiDeclarations vbCp
        Code_MoveDeclarations vbCp
    End If
    If chkOnError.Value = 1 Then
        Code_InsertOnError vbCp
    End If
    If chkFormat.Value = 1 Then
        Code_Format vbCp
    End If
    Err.Clear
    Return
    Err.Clear
End Sub
Private Sub Dao_CreateTable(ByVal DbName As String, ByVal dbTable As String, ByVal fldName As String, Optional ByVal fldType As String = "", Optional ByVal fldSize As String = "", Optional ByVal Fldidx As String = "", Optional ByVal FldAutoIncrement As String = "")
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
            NewFld.DefaultValue = ""
        End Select
        If MvSearch(FldAutoIncrement, spFlds(newCnt), ",") > 0 Then
            NewFld.Attributes = DAO.dbAutoIncrField
        End If
        NewTb.Fields.Append NewFld
    Next
    For newCnt = 1 To totIdx
        spIndx(newCnt) = Trim$(spIndx(newCnt))
        newPos = spIndx(newCnt)
        Set NewIdx = NewTb.CreateIndex(spFlds(newPos))
        Set NewFld = NewIdx.CreateField(spFlds(newPos))
        NewIdx.Fields.Append NewFld
        NewTb.Indexes.Append NewIdx
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
Private Sub Code_RemoveLines(vbCp As VBIDE.CodePane, ByVal StrSearch As String, Optional ExactMatch As Boolean = True)
    On Error Resume Next
    Dim totLines As Long
    Dim cntLines As Long
    Dim curLine As String
    totLines = vbCp.CodeModule.CountOfLines
    progBar.Max = totLines
    progBar.Min = 0
    progBar.Value = totLines
    For cntLines = totLines To 1 Step -1
        DoEvents
        progBar.Value = cntLines
        curLine = vbCp.CodeModule.Lines(cntLines, 1)
        curLine = Trim$(curLine)
        If ExactMatch = True Then
            If LCase$(StrSearch) = LCase$(curLine) Then
                vbCp.CodeModule.DeleteLines cntLines
            End If
        Else
            If InStr(1, LCase$(curLine), LCase$(StrSearch), vbTextCompare) > 0 Then
                vbCp.CodeModule.DeleteLines cntLines
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_TrimLines(vbCp As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Trimming Lines"
    Dim totLines As Long
    Dim cntLines As Long
    Dim curLine As String
    totLines = vbCp.CodeModule.CountOfLines
    progBar.Max = totLines
    progBar.Min = 0
    progBar.Value = totLines
    For cntLines = 1 To totLines
        DoEvents
        progBar.Value = cntLines
        curLine = vbCp.CodeModule.Lines(cntLines, 1)
        curLine = Trim$(curLine)
        vbCp.CodeModule.ReplaceLine cntLines, curLine
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_Format(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Formating Code"
    Dim CurrTab As Integer
    Dim CurrCommand As String
    Dim CurrLine As String
    Dim KydStart As Integer
    Dim TabAfter As Boolean
    Dim TabSpace As String
    Dim Count2 As Integer
    Dim LastKeyWord As String
    Dim tCount As Long
    Dim tLines As Long
    Code_RemoveLines CompoA, ""
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = LCase$(Trim$(CompoA.CodeModule.Lines(tCount, 1)))
        KydStart = InStr(1, CurrLine, " ")
        If KydStart = 0 Then
            KydStart = Len(CurrLine)
        End If
        CurrCommand = Trim$(Mid$(CurrLine, 1, KydStart))
        Select Case CurrCommand
        Case "public", "private"
            If Mid$(Trim$(Mid$(CurrLine, Len(CurrCommand) + 1)), 1, Len("sub")) = "sub" Or Mid$(Trim$(Mid$(CurrLine, Len(CurrCommand) + 1)), 1, Len("function")) = "function" Or Mid$(Trim$(Mid$(CurrLine, Len(CurrCommand) + 1)), 1, Len("type")) = "type" Or _
                Mid$(Trim$(Mid$(CurrLine, Len(CurrCommand) + 1)), 1, Len("property")) = "property" Or Mid$(Trim$(Mid$(CurrLine, Len(CurrCommand) + 1)), 1, Len("enum")) = "enum" Then
                CurrTab = CurrTab + 1
                TabAfter = True
            End If
        Case "if"
            If Len(CurrLine) > 6 Then
                If Mid$(CurrLine, Len(CurrLine) - 3, 4) <> "then" Then
                    If Right$(CurrLine, 1) = "_" Then
                        CurrTab = CurrTab + 1
                        TabAfter = True
                    End If
                Else
                    CurrTab = CurrTab + 1
                    TabAfter = True
                End If
            End If
        Case "while", "do", "select", "for", "sub", "function", "type", "enum", "property", "open", "with"
            CurrTab = CurrTab + 1
            TabAfter = True
        Case "end", "wend", "loop", "next", "exit"
            If Mid$(CurrLine, 1, Len("end if")) = "end if" Or Mid$(CurrLine, 1, Len("next")) = "next" Or Mid$(CurrLine, 1, Len("wend")) = "wend" Or Mid$(CurrLine, 1, Len("loop")) = "loop" Or Mid$(CurrLine, 1, Len("end select")) = "end select" Or _
                Mid$(CurrLine, 1, Len("end with")) = "end with" Then
                CurrTab = CurrTab - 1
                TabAfter = False
            End If
            If Mid$(CurrLine, 1, Len("end function")) = "end function" Or Mid$(CurrLine, 1, Len("end sub")) = "end sub" Or Mid$(CurrLine, 1, Len("end type")) = "end type" Or Mid$(CurrLine, 1, Len("end property")) = "end property" Or _
                Mid$(CurrLine, 1, Len("end enum")) = "end enum" Then
                CurrTab = 0
                TabAfter = False
            End If
        Case "close"
            If Mid$(CurrLine, 1, Len("close")) = "close" Then
                CurrTab = CurrTab - 1
                TabAfter = False
            End If
        Case "else", "case", "elseif"
            'If LastKeyWord <> "select" Then
            CurrTab = CurrTab - 1
            TabSpace = ""
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                Next
            End If
            'End If
            CurrTab = CurrTab + 1
            TabAfter = True
        End Select
        If TabAfter = True Then
            CompoA.CodeModule.ReplaceLine tCount, TabSpace & Trim$(CompoA.CodeModule.Lines(tCount, 1))
            TabSpace = ""
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                Next
            End If
        Else
            TabSpace = ""
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                Next
            End If
            CompoA.CodeModule.ReplaceLine tCount, TabSpace & Trim$(CompoA.CodeModule.Lines(tCount, 1))
        End If
        If Len(CurrCommand) > 1 Then
            If Mid$(CurrCommand, Len(CurrCommand) - 1, 1) = ":" Then
                CompoA.CodeModule.ReplaceLine tCount, Trim$(CompoA.CodeModule.Lines(tCount, 1))
            End If
        End If
        LastKeyWord = CurrCommand
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_RemoveLineContinuation(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Removing Line Continuation"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim NextLine As String
    Dim nLen As Long
    tLines = CompoA.CodeModule.CountOfLines - 1
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
foundone:
        CurrLine = CompoA.CodeModule.Lines(tCount, 1)
        NextLine = CompoA.CodeModule.Lines(tCount + 1, 1)
        nLen = Len(CurrLine) + Len(NextLine)
        If Right$(CurrLine, 1) = "_" Then
            If nLen <= 255 Then
                CompoA.CodeModule.ReplaceLine tCount, Trim$(Left$(CurrLine, Len(CurrLine) - 1)) & " " & NextLine
                CompoA.CodeModule.DeleteLines tCount + 1
                GoTo foundone
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_InsertDoEvents(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Inserting DoEvents On Loops"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim NextLine As String
    tLines = CompoA.CodeModule.CountOfLines - 1
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If Left$(CurrLine, Len("While ")) = "While " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(CurrLine, Len("For ")) = "For " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(CurrLine, Len("Do ")) = "Do " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        End If
    Next
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = tLines
    progBar.Min = 0
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If Left$(CurrLine, Len("While ")) = "While " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(CurrLine, Len("For ")) = "For " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        ElseIf Left$(CurrLine, Len("Do ")) = "Do " Then
            If InStr(1, NextLine, "DoEvents", vbTextCompare) = 0 Then
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & "DoEvents"
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_BreakMultiDeclarations(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Break Multi Declarations"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim newL As String
    Dim spLine() As String
    Dim spCnt As Integer
    Dim spTot As Integer
    Dim AsPos As Long
    Dim varType As String
    Dim newDecl As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        newL = ""
        If Left$(CurrLine, 4) = "Dim " Then
            If InStr(1, CurrLine, ", ", vbTextCompare) > 0 Then
                CurrLine = Mid$(CurrLine, 5)
                AsPos = InStr(1, CurrLine, " As ", vbTextCompare)
                If AsPos > 0 Then
                    varType = Trim$(Mid$(CurrLine, AsPos + 4))
                Else
                    varType = ""
                End If
                spLine = Split(CurrLine, ", ")
                spTot = UBound(spLine)
                For spCnt = 0 To spTot
                    newDecl = "Dim " & Trim$(spLine(spCnt))
                    newDecl = FixDeclaration(newDecl)
                    If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                        newDecl = newDecl & " As " & varType
                    End If
                    newL = newL & newDecl & vbNewLine
                Next
                CompoA.CodeModule.ReplaceLine tCount, newL
            Else
                newDecl = FixDeclaration(CurrLine)
                If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then newDecl = newDecl & " As Variant"
                CompoA.CodeModule.ReplaceLine tCount, newDecl
            End If
        End If
    Next
    progBar.Value = 0
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = tLines
    progBar.Min = 0
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        newL = ""
        If Left$(CurrLine, 4) = "Dim " Then
            If InStr(1, CurrLine, ", ", vbTextCompare) > 0 Then
                CurrLine = Mid$(CurrLine, 5)
                AsPos = InStr(1, CurrLine, " As ", vbTextCompare)
                If AsPos > 0 Then
                    varType = Trim$(Mid$(CurrLine, AsPos + 4))
                Else
                    varType = ""
                End If
                spLine = Split(CurrLine, ", ")
                spTot = UBound(spLine)
                For spCnt = 0 To spTot
                    newDecl = "Dim " & Trim$(spLine(spCnt))
                    newDecl = FixDeclaration(newDecl)
                    If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then
                        newDecl = newDecl & " As " & varType
                    End If
                    newL = newL & newDecl & vbNewLine
                Next
                CompoA.CodeModule.ReplaceLine tCount, newL
            Else
                newDecl = FixDeclaration(CurrLine)
                If InStr(1, newDecl, " As ", vbTextCompare) = 0 Then newDecl = newDecl & " As Variant"
                CompoA.CodeModule.ReplaceLine tCount, newDecl
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_ConvertStringFunctions(CompoA As VBIDE.CodePane, Optional bImprove As Boolean = True)
    On Error Resume Next
    Caption = "Cool Code: Speeding String Functions"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        If bImprove = True Then
            CurrLine = ImprovedStringFunction(CurrLine)
        Else
            CurrLine = UnImprovedStringFunction(CurrLine)
        End If
        CompoA.CodeModule.ReplaceLine tCount, CurrLine
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Code_MoveComment(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Moving Comments"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim ComPos As Long
    Dim comStr As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        ComPos = InStrRev(CurrLine, "'")
        If ComPos > 1 Then
            If Mid$(CurrLine, ComPos - 1, 1) = Quote Then
                If Mid$(CurrLine, ComPos + 1, 1) = Quote Then
                    ComPos = 0
                Else
                    comStr = Mid$(CurrLine, ComPos)
                    CurrLine = Left$(CurrLine, ComPos - 1)
                    CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & comStr
                End If
            End If
        End If
    Next
    progBar.Value = 0
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = tLines
    progBar.Min = 0
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        ComPos = InStr(1, CurrLine, "'")
        If ComPos > 1 Then
            If Mid$(CurrLine, ComPos - 1, 1) = " " Then
                comStr = Mid$(CurrLine, ComPos)
                CurrLine = Left$(CurrLine, ComPos - 1)
                CompoA.CodeModule.ReplaceLine tCount, CurrLine & vbNewLine & comStr
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Function ImprovedStringFunction(ByVal StrLine As String) As String
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
    Next
    ImprovedStringFunction = StrLine
    Err.Clear
End Function
Sub IniStrFunctions()
    On Error Resume Next
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
    Err.Clear
End Sub
Function UnImprovedStringFunction(ByVal StrLine As String) As String
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
    Next
    UnImprovedStringFunction = StrLine
    Err.Clear
End Function
Private Sub Code_SpeedNextLoopEnd(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Speeding Next Loop Counter"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Min = 0
    progBar.Value = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        CurrLine = ImprovedNextLoop(CurrLine)
        CompoA.CodeModule.ReplaceLine tCount, CurrLine
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Function ImprovedNextLoop(ByVal StrLine As String) As String
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
        Next
        ImprovedNextLoop = Join(sParts, ":")
    End If
    Err.Clear
End Function
Private Sub Code_PassBy(CompoA As VBIDE.CodePane, Optional ByVal PassBy As String = "ByVal")
    On Error Resume Next
    Caption = "Cool Code: Passing String Variables"
    Dim CurrLine As String
    Dim tLines As Long
    Dim tCount As Long
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        If IsMethod(CurrLine) = True Then
            CurrLine = StringNewProcedure(CurrLine, PassBy)
            CompoA.CodeModule.ReplaceLine tCount, CurrLine
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Function IsMethod(ByVal CurrLine As String) As Boolean
    On Error Resume Next
    If Left$(CurrLine, Len("Private Sub ")) = "Private Sub " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Private Function ")) = "Private Function " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Function ")) = "Function " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Sub ")) = "Sub " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Public Sub ")) = "Public Sub " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Public Function ")) = "Public Function " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Public Property ")) = "Public Property " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Property ")) = "Property " Then
        IsMethod = True
    ElseIf Left$(CurrLine, Len("Private Property ")) = "Private Property " Then
        IsMethod = True
    Else
        IsMethod = False
    End If
    Err.Clear
End Function
Private Function IsVariable(ByVal CurrLine As String) As Boolean
    On Error Resume Next
    If Left$(CurrLine, Len("Dim ")) = "Dim " Then
        IsVariable = True
    Else
        IsVariable = False
    End If
    Err.Clear
End Function
Function StringNewProcedure(ByVal strDeclaration As String, Optional ByVal PassBy As String = "ByVal") As String
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
Function StringArguements(ByVal strDeclaration As String) As String
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
Function StringArguementsPass(ByVal strArguements As String, Optional ByVal NewValue As String = "ByVal") As String
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
        Else
            StrA = Replace$(StrA, "ByVal ", "")
        End If
        If InStr(1, StrA, "ByRef", vbTextCompare) > 0 Then
            StrA = Replace$(StrA, "ByRef", "")
        End If
        StrB(aCnt) = Trim$(StrA)
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
Function StringProcedure(ByVal strDeclaration As String, Optional IncludeBrackets As Boolean = True) As String
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
Private Sub Code_InitializeOptionalVariables(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Initializing Optional Variables"
    Dim tCount As Long
    Dim tLines As Long
    Dim CurrLine As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        If IsMethod(CurrLine) = True Then
            CurrLine = OptionalizedProcedure(CurrLine)
            CompoA.CodeModule.ReplaceLine tCount, CurrLine
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Function OptionalizedProcedure(ByVal strDeclaration As String) As String
    On Error Resume Next
    Dim sArguements As String
    Dim sNewArguements As String
    Dim sProcedure As String
    Dim nProcedure As String
    Select Case Right$(strDeclaration, 1)
    Case "_"
        OptionalizedProcedure = strDeclaration
    Case Else
        sArguements = StringArguements(strDeclaration)
        sNewArguements = StringOptionalInitialize(sArguements)
        sProcedure = StringProcedure(strDeclaration)
        nProcedure = Replace$(sProcedure, "()", "(" & sNewArguements & ")")
        OptionalizedProcedure = nProcedure
    End Select
    Err.Clear
End Function
Function StringOptionalInitialize(ByVal strArguements As String) As String
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
                    StrA = StringConcat(StrA, " = """"")
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
                    StrA = StringConcat(StrA, " = """"")
                End If
                If InStr(StrA, "&") > 0 Then
                    StrA = StringConcat(StrA, " = 0")
                End If
            End Select
            StrB(aCnt) = StrA
        End Select
    Next
    StringOptionalInitialize = StringRemoveDelim(Join(StrB, ","), ",")
    Erase StrB
    StrA = vbNullString
    Erase StrN
    Err.Clear
End Function
Private Function StringRemoveDelim(ByVal StrData As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(StrData)
    intDelimSize = Len(Delim)
    strLast = Right$(StrData, intDelimSize)
    Select Case LCase$(strLast)
    Case LCase$(Delim)
        StringRemoveDelim = Left$(StrData, (intDataSize - intDelimSize))
    Case Else
        StringRemoveDelim = StrData
    End Select
    strLast = vbNullString
    Err.Clear
End Function
Function StringConcat(ByVal dest As String, ByVal Source As String) As String
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
Private Sub Code_InsertOnError(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Inserting On Error Resume Next"
    Code_RemoveLines CompoA, "On Error Resume Next"
    Code_RemoveLines CompoA, "Err.Clear"
    Code_RemoveLines CompoA, ""
    'On Error Resume Next"
    Code_RemoveLines CompoA, ""
    'Err.Clear"
    Code_RemoveLines CompoA, ""
    Code_RemoveLines CompoA, ""
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim NextLine As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Min = 0
    progBar.Value = tLines
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        If IsMethodEnd(CurrLine) = True Then
            CompoA.CodeModule.InsertLines tCount, vbTab & "Err.Clear"
        End If
    Next
    progBar.Value = 0
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If Left$(NextLine, Len("On Error ")) <> "On Error " Then
            If IsMethod(CurrLine) = True Then
                If Right$(CurrLine, 1) <> "_" Then
                    CompoA.CodeModule.InsertLines tCount + 1, vbTab & "On Error Resume Next"
                End If
            End If
        End If
    Next
    progBar.Value = 0
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = tLines
    progBar.Min = 0
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If Left$(NextLine, Len("On Error ")) <> "On Error " Then
            If IsMethod(CurrLine) = True Then
                If Right$(CurrLine, 1) <> "_" Then
                    CompoA.CodeModule.InsertLines tCount + 1, vbTab & "On Error Resume Next"
                End If
            End If
        End If
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Function IsMethodEnd(ByVal CurrLine As String) As Boolean
    On Error Resume Next
    Select Case CurrLine
    Case "End Sub", "Exit Function", "Exit Sub", "End Function", "End Property", "Exit Property", "Return"
        IsMethodEnd = True
    Case Else
        IsMethodEnd = False
    End Select
    Err.Clear
End Function
Private Function IsEndOfMethod(ByVal CurrLine As String) As Boolean
    On Error Resume Next
    Select Case CurrLine
    Case "End Sub", "End Function"
        IsEndOfMethod = True
    Case Else
        IsEndOfMethod = False
    End Select
    Err.Clear
End Function
Private Function VariableType(ByVal CurrLine As String) As String
    On Error Resume Next
    VariableType = Split(CurrLine, " ")(3)
    Err.Clear
End Function
Private Function VariableName(ByVal CurrLine As String) As String
    On Error Resume Next
    Dim bPos As Long
    Dim bStr As String
    bStr = Split(CurrLine, " ")(1)
    bPos = InStr(1, bStr, "(")
    If bPos > 0 Then
        bStr = Left$(bStr, bPos - 1)
    End If
    VariableName = Trim$(bStr)
    Err.Clear
End Function
Private Sub Code_RemoveEmpty(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Removing Empty Procedures"
    Code_RemoveLines CompoA, "On Error Resume Next"
    Code_RemoveLines CompoA, "Err.Clear"
    Code_RemoveLines CompoA, ""
    'On Error Resume Next"
    Code_RemoveLines CompoA, ""
    'Err.Clear"
    Code_RemoveLines CompoA, ""
    Code_RemoveLines CompoA, ""
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim NextLine As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Min = 0
    progBar.Value = tLines
    For tCount = tLines To 1 Step -1
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If IsMethod(CurrLine) = True Then
            If IsMethodEnd(NextLine) = True Then
                CompoA.CodeModule.ReplaceLine tCount, " "
                CompoA.CodeModule.ReplaceLine tCount + 1, " "
            End If
        End If
    Next
    progBar.Value = 0
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Min = 0
    progBar.Value = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        NextLine = Trim$(CompoA.CodeModule.Lines(tCount + 1, 1))
        If IsMethod(CurrLine) = True Then
            If IsMethodEnd(NextLine) = True Then
                CompoA.CodeModule.ReplaceLine tCount, " "
                CompoA.CodeModule.ReplaceLine tCount + 1, " "
            End If
        End If
    Next
    progBar.Value = 0
    Code_RemoveLines CompoA, ""
    Err.Clear
End Sub
Private Sub Code_NormalizeIf(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Normalizing If Statement"
    Dim tLines As Long
    Dim tCount As Long
    Dim CurrLine As String
    Dim IfPos As Long
    Dim ThenPos As Long
    Dim ElsePos As Long
    Dim ComPos As Long
    Dim spIf() As String
    Dim useLine As String
    Dim comStr As String
    Dim midStr As String
    tLines = CompoA.CodeModule.CountOfLines
    progBar.Max = tLines
    progBar.Value = 0
    progBar.Min = 0
    For tCount = 1 To tLines
        DoEvents
        progBar.Value = tCount
        CurrLine = Trim$(CompoA.CodeModule.Lines(tCount, 1))
        IfPos = InStr(1, CurrLine, "If ", vbTextCompare)
        ThenPos = InStr(1, CurrLine, " Then", vbTextCompare)
        ElsePos = InStr(1, CurrLine, " Else ", vbTextCompare)
        ComPos = InStrRev(CurrLine, "'", , vbTextCompare)
        comStr = ""
        If IfPos = 1 Then
            If Right$(CurrLine, 1) = "_" Then
                GoTo NextLine
            End If
            If ComPos > 0 Then
                ' we have a comment
                If ComPos > ThenPos Then
                    If Mid$(CurrLine, ComPos - 1, 1) = Quote Then
                        If Mid$(CurrLine, ComPos + 1, 1) = Quote Then
                            ComPos = 0
                        Else
                            comStr = Mid$(CurrLine, ComPos)
                            CurrLine = Left$(CurrLine, ComPos - 1)
                        End If
                    End If
                End If
            End If
            If ElsePos > 0 Then
                CurrLine = CurrLine & vbNewLine & "End If"
                CurrLine = Replace$(CurrLine, "Else", vbNewLine & "Else" & vbNewLine)
            End If
            spIf = Split(CurrLine, vbNewLine)
            useLine = Trim$(spIf(0))
            ThenPos = InStr(useLine, " Then")
            If Right$(useLine, 4) <> "Then" Then
                midStr = Mid$(useLine, ThenPos + 5)
                useLine = Left$(useLine, ThenPos + 5) & vbNewLine & midStr
                If ElsePos = 0 Then
                    useLine = useLine & vbNewLine & "End If"
                End If
                spIf(0) = useLine
                CurrLine = Join(spIf, vbNewLine)
            End If
            If ComPos > 0 Then
                CurrLine = Replace$(CurrLine, " Then", " Then" & vbNewLine & comStr)
            End If
            CompoA.CodeModule.ReplaceLine tCount, CurrLine
        End If
NextLine:
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub Backup_ToCompressedFile(ByVal ArchiveName As String, ParamArray FoldersToBackup())
    On Error Resume Next
    Dim eachFolder As Variant
    Dim xFolder As String
    Dim vFiles As New Collection
    Dim tAdded As Long
    Screen.MousePointer = vbHourglass
    For Each eachFolder In FoldersToBackup
        DoEvents
        xFolder = CStr(eachFolder)
        Call TotalDirFiles(xFolder, vFiles)
    Next
    tAdded = Zip_Add2Archive(ArchiveName, vFiles, 1, True, False, 9)
    If tAdded <> 0 Then
        Call MsgBox("Not all files could be added to the archive!", vbOKOnly + vbExclamation, tAdded & " Files Not Added")
    End If
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Public Function TotalDirFiles(ByVal DirPath As String, FilesCollection As Collection, Optional ByVal FilePattern As String = "*.*") As Long
    On Error Resume Next
    Dim sFile As String
    Dim StrF As String
    Dim StrP As String
    Dim lngL As Long
    StrP = DirPath & "\"
    StrF = StrP & FilePattern
    sFile = Dir$(StrF)
    lngL = Len(sFile)
    Do While lngL
        sFile = StrP & sFile
        FilesCollection.Add sFile, sFile
        sFile = Dir$
        lngL = Len(sFile)
    Loop
    TotalDirFiles = FilesCollection.Count
    sFile = vbNullString
    StrF = vbNullString
    StrP = vbNullString
    Err.Clear
End Function
Private Function Zip_Add2Archive(ZipFileName As String, Files As Collection, Action As Integer, StorePathInfo As Boolean, UseDOS83 As Boolean, CompressionLevel As Integer) As Long
    On Error Resume Next
    Dim i As Long
    Dim Result As Long
    Dim nAdded As Long
    Dim FilesToAdd As Collection
    Dim i_Tot As Long
    nAdded = 0
    If Not win_Function_Exist("zipit.dll", "AddFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Not win_Function_Exist("zipit.dll", "ExtractFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Not win_Function_Exist("zipit.dll", "DeleteFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Archive.Count = 0 Then
        If Dir$(ZipFileName, vbHidden Or vbSystem Or vbReadOnly) <> "" Then
            Kill ZipFileName
        End If
    End If
    Set FilesToAdd = FindFiles(Files)
    i_Tot = FilesToAdd.Count
    progBar.Max = i_Tot
    progBar.Value = 0
    progBar.Min = 0
    For i = 1 To i_Tot
        DoEvents
        progBar.Value = i
        If AddFile(ZipFileName, FilesToAdd(i), StorePathInfo, UseDOS83, Action, CompressionLevel) Then
            Result = Result + 1
        Else
            nAdded = nAdded + 1
        End If
    Next
    progBar.Value = 0
    Zip_Add2Archive = nAdded
    Err.Clear
End Function
Private Function FindFiles(Files As Collection)
    On Error Resume Next
    Dim Result As New Collection
    Dim Path As String
    Dim r As String
    Dim i As Long
    Dim i_Tot As Long
    i_Tot = Files.Count
    For i = 1 To i_Tot
        Path = File_ParsePath(Files(i))
        r = Dir$(Files(i), vbHidden Or vbSystem Or vbReadOnly)
        Do Until r = ""
            Result.Add Path & r
            r = Dir$()
        Loop
    Next
    Set FindFiles = Result
    Err.Clear
End Function
Private Function File_ParsePath(Path As String) As String
    On Error Resume Next
    Dim A As Long
    Dim A_Cnt As Long
    A_Cnt = Len(Path)
    For A = A_Cnt To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            If Mid$(Path, A, 1) = "\" Then
                File_ParsePath = LCase$(Left$(Path, A - 1) & "\")
            Else
                File_ParsePath = LCase$(Left$(Path, A - 1) & "/")
            End If
            Err.Clear
            Exit Function
        End If
    Next
    Err.Clear
End Function
Private Function StringGetFileToken(ByVal StrFileName As String, Optional ByVal Sretrieve As String = "", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim SNew As String
    StringGetFileToken = StrFileName
    If Len(Sretrieve) = 0 Then
        Sretrieve = "F"
    End If
    If Len(Delim) = 0 Then
        Delim = "\"
    End If
    Select Case UCase$(Sretrieve)
    Case "D"
        StringGetFileToken = Left$(StrFileName, 3)
    Case "F"
        intNum = InStrRev(StrFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(StrFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(StrFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(StrFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(StrFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Mid$(StrFileName, intNum + 1)
        End If
    Case "FO"
        SNew = StrFileName
        intNum = InStrRev(SNew, Delim)
        If intNum <> 0 Then
            SNew = Mid$(SNew, intNum + 1)
        End If
        intNum = InStrRev(SNew, ".")
        If intNum <> 0 Then
            SNew = Left$(SNew, intNum - 1)
        End If
        StringGetFileToken = SNew
    Case "PF"
        intNum = InStrRev(StrFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Left$(StrFileName, intNum - 1)
        End If
    End Select
    SNew = vbNullString
    Err.Clear
End Function
Private Function ExactPath(ByVal StrValue As String) As String
    On Error Resume Next
    If Right$(StrValue, 1) = "\" Then
        StrValue = Left$(StrValue, Len(StrValue) - 1)
    End If
    ExactPath = StrValue
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
            If Mid$(spStr, 5, 4) = "Each" Then GoTo NextLine
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
    Next
    RecodeComplex = MvFromArray(spLines)
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
            If Mid$(spStr, 5, 4) = "Each" Then GoTo NextLine
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
    Next
    If complexCnt = 0 Then
        IsSourceComplex = False
    Else
        IsSourceComplex = True
    End If
    Err.Clear
End Function
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
Private Sub Code_FixComplexLoops(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Recoding Complex Loops"
    Dim procStart As Long
    Dim procEnd As Long
    Dim varNames As New Collection
    Dim StrLine As String
    Dim procValues As String
    Dim tLines As Long
    Dim cLines As Long
    Dim CurrLine As String
    Dim newLines As String
    Dim SourceCode As String
    Dim spLines() As String
    Dim Db As DAO.Database
    Dim tb As DAO.Recordset
    Dim olds As String
    Dim news As String
    Do Until Dao_TableExists(Dbase, "Complex") = True
        Dao_CreateTable Dbase, "Complex", "File,Source,Improved", "te,me,me", "255,,", "1"
    Loop
    Set Db = DAO.OpenDatabase(Dbase)
    Set tb = Db.OpenRecordset("Complex")
    tLines = CompoA.CodeModule.CountOfLines
    SourceCode = CompoA.CodeModule.Lines(1, tLines)
    spLines = Split(SourceCode, vbNewLine)
    tLines = UBound(spLines)
    progBar.Max = tLines + 1
    progBar.Value = 0
    progBar.Min = 0
    For cLines = 0 To tLines
        DoEvents
        progBar.Value = cLines + 1
        CurrLine = Trim$(spLines(cLines))
        If IsMethod(CurrLine) = True Then
            procValues = CurrLine & vbNewLine
            procStart = cLines
        ElseIf IsEndOfMethod(CurrLine) = True Then
            procValues = procValues & CurrLine
            procEnd = cLines
            If IsSourceComplex(procValues) = True Then
                newLines = RecodeComplex(procValues, varNames)
                newLines = Code_DimStatements(newLines)
                olds = Code_RemoveBlanks(procValues)
                news = Code_RemoveBlanks(newLines)
                If LCase$(olds) <> LCase$(news) Then
                    tb.AddNew
                    tb!file = CompoA.CodeModule.Parent.FileNames(1)
                    tb!Source = procValues
                    tb!improved = newLines
                    tb.Update
                End If
            End If
            Set varNames = New Collection
        ElseIf IsVariable(CurrLine) = True Then
            procValues = procValues & CurrLine & vbNewLine
            CurrLine = FixDeclaration(CurrLine)
            StrLine = VariableName(CurrLine) & "," & VariableType(CurrLine)
            varNames.Add StrLine
        Else
            procValues = procValues & CurrLine & vbNewLine
        End If
    Next
    progBar.Value = 0
    If tb.RecordCount > 0 Then
        Code_TrimLines vbCp
        CurrLine = CompoA.CodeModule.Lines(1, CompoA.CodeModule.CountOfLines)
        tLines = tb.RecordCount
        tb.MoveFirst
        progBar.Max = tLines
        For cLines = 1 To tLines
            DoEvents
            progBar.Value = cLines
            olds = tb!Source & ""
            news = tb!improved & ""
            CurrLine = Replace$(CurrLine, olds, news, , , vbTextCompare)
            tb.MoveNext
        Next
        progBar.Value = 0
        CompoA.CodeModule.DeleteLines 1, CompoA.CodeModule.CountOfLines
        CompoA.CodeModule.AddFromString CurrLine
    End If
    tb.Close
    Db.Close
    Set tb = Nothing
    Set Db = Nothing
    Err.Clear
End Sub
Private Sub Code_MoveDeclarations(CompoA As VBIDE.CodePane)
    On Error Resume Next
    Caption = "Cool Code: Moving Declarations"
    Dim procStart As Long
    Dim procEnd As Long
    Dim procValues As String
    Dim tLines As Long
    Dim cLines As Long
    Dim CurrLine As String
    Dim newLines As String
    Dim SourceCode As String
    Dim spLines() As String
    Dim Db As DAO.Database
    Dim tb As DAO.Recordset
    Dim olds As String
    Dim news As String
    Do Until Dao_TableExists(Dbase, "Complex") = True
        Dao_CreateTable Dbase, "Complex", "File,Source,Improved", "te,me,me", "255,,", "1"
    Loop
    Set Db = DAO.OpenDatabase(Dbase)
    Set tb = Db.OpenRecordset("Complex")
    tLines = CompoA.CodeModule.CountOfLines
    SourceCode = CompoA.CodeModule.Lines(1, tLines)
    spLines = Split(SourceCode, vbNewLine)
    tLines = UBound(spLines)
    progBar.Max = tLines + 1
    progBar.Value = 0
    progBar.Min = 0
    For cLines = 0 To tLines
        DoEvents
        progBar.Value = cLines + 1
        CurrLine = Trim$(spLines(cLines))
        If IsMethod(CurrLine) = True Then
            procValues = CurrLine & vbNewLine
            procStart = cLines
        ElseIf IsEndOfMethod(CurrLine) = True Then
            procValues = procValues & CurrLine
            procEnd = cLines
            newLines = Code_DimStatements(procValues)
            olds = Code_RemoveBlanks(procValues)
            news = Code_RemoveBlanks(newLines)
            If LCase$(olds) <> LCase$(news) Then
                tb.AddNew
                tb!file = CompoA.CodeModule.Parent.FileNames(1)
                tb!Source = procValues
                tb!improved = newLines
                tb.Update
            End If
        Else
            procValues = procValues & CurrLine & vbNewLine
        End If
    Next
    progBar.Value = 0
    If tb.RecordCount > 0 Then
        Code_TrimLines vbCp
        CurrLine = CompoA.CodeModule.Lines(1, CompoA.CodeModule.CountOfLines)
        tLines = tb.RecordCount
        tb.MoveFirst
        progBar.Max = tLines
        For cLines = 1 To tLines
            DoEvents
            progBar.Value = cLines
            olds = tb!Source & ""
            news = tb!improved & ""
            CurrLine = Replace$(CurrLine, olds, news, , , vbTextCompare)
            tb.MoveNext
        Next
        progBar.Value = 0
        CompoA.CodeModule.DeleteLines 1, CompoA.CodeModule.CountOfLines
        CompoA.CodeModule.AddFromString CurrLine
    End If
    tb.Close
    Db.Close
    Set tb = Nothing
    Set Db = Nothing
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
    DoEvents
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
Private Function Code_RemoveBlanks(ByVal StrValue As String) As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    Dim xLine As String
    xRslt = ""
    xData = Split(StrValue, vbNewLine)
    xTot = UBound(xData)
    For xCnt = 0 To xTot
        DoEvents
        xLine = Trim$(xData(xCnt))
        If Len(xLine) <> 0 Then
            xRslt = xRslt & xLine & vbNewLine
        End If
    Next
    Code_RemoveBlanks = StringRemoveDelim(xRslt, vbNewLine)
    Erase xData
    Err.Clear
End Function
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
        DoEvents
        spStr = Trim$(spLines(spCnt))
        If Left$(spStr, 4) = "Dim " Then
            newDim.Add spStr, spStr
            spLines(spCnt) = ""
            If dimStart = 0 Then dimStart = spCnt
        End If
    Next
    If newDim.Count > 0 Then spStr = MvFromCollection(newDim)
    If dimStart > 0 Then spLines(dimStart) = spStr
    Code_DimStatements = MvFromArray(spLines)
    Err.Clear
End Function
Private Function MvFromArray(Varray() As String, Optional StartPos As Long = 0) As String
    On Error Resume Next
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As String
    totArray = UBound(Varray)
    BldStr = ""
    For i = StartPos To totArray
        strL = Varray(i)
        Select Case i
        Case totArray
            BldStr = BldStr & strL
        Case Else
            BldStr = BldStr & strL & vbNewLine
        End Select
    Next
    MvFromArray = BldStr
    Err.Clear
End Function
Private Function StringParse(RetArray() As String, ByVal strText As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim VarA As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    VarA = VarE + 1
    ReDim RetArray(VarA)
    For varCnt = VarS To VarE
        VarA = varCnt + 1
        RetArray(VarA) = varArray(varCnt)
    Next
    Erase varArray
    StringParse = UBound(RetArray)
    Err.Clear
End Function
Private Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = 0
    If Len(MvSearch) = 0 Or Len(StrLookFor) = 0 Then
        MvSearch = 0
        Err.Clear
        Exit Function
    End If
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(TheFields, StringMv, Delim)
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Erase TheFields
    Err.Clear
End Function
Private Function boolFileExists(ByVal Filename As String) As Boolean
    On Error Resume Next
    boolFileExists = False
    If Len(Filename) = 0 Then
        Err.Clear
        Exit Function
    End If
    boolFileExists = IIf(Dir$(Filename) <> "", True, False)
    Err.Clear
End Function
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    If ArrayTot = 0 Then
        Err.Clear
        Exit Function
    End If
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
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
        Next
    End With
    Db.Close
    Set Db = Nothing
    Err.Clear
End Function
Private Function MvFromCollection(objCollection As Collection) As String
    On Error Resume Next
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        Select Case xCnt
        Case xTot
            sRet = sRet & objCollection.Item(xCnt)
        Case Else
            sRet = sRet & objCollection.Item(xCnt) & vbNewLine
        End Select
    Next
    MvFromCollection = sRet
    Err.Clear
End Function

Sub Initialize(YourVbInstance As VBIDE.VBE, YourCommandBar As Office.CommandBar)
    On Error GoTo Err_Handler:
    Set VBInstance = YourVbInstance
   
    Set ccRemoveBlanks = YourCommandBar.Controls.Add(msoControlButton)
    Set ccTrimLines = YourCommandBar.Controls.Add(msoControlButton)
  
    With ccRemoveBlanks
        .ToolTipText = "Remove blank lines in module"
        .Style = msoButtonWrapCaption
        .State = msoButtonUp
        .Caption = "Blanks"
    End With
   
    With ccTrimLines
        .ToolTipText = "Trim lines in module"
        .Style = msoButtonWrapCaption
        .Caption = "TrimLines"
        .State = msoButtonUp
    End With
   
    'link this button to an event sink so event, that we can write
    'code for the button,that will execute,k when it is clicked
    Set ccRemoveBlanksEvents = VBinst.Events.CommandBarEvents(ccRemoveBlanks)
    Set ccTrimLinesEvents = VBinst.Events.CommandBarEvents(ccTrimLines)
    
    Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & varType(Me) & ".Initialize"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description
    Err.Clear
    Resume Next
End Sub

Sub Destroy()
    On Error Resume Next
    ccRemoveBlanks.Delete
    ccTrimLines.Delete
 
    'delete variables
    Set ccRemoveBlanks = Nothing
    Set ccTrimLines = Nothing
 
    'delete event objects
    Set ccRemoveBlanksEvents = Nothing
    Set ccTrimLinesEvents = Nothing
    
    Set VBInstance = Nothing
End Sub

