VERSION 5.00
Begin VB.Form MyListBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyListBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   495
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1215
      Top             =   1740
   End
   Begin VB.VScrollBar VScroll 
      Height          =   675
      Left            =   1875
      TabIndex        =   0
      Top             =   1290
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   720
      Picture         =   "MyListBox.frx":000C
      Top             =   1335
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "MyListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const MAX_SHOWN_ITEMS As Long = 7
Private Const ITEM_HEIGHT As Long = 17
Private TextOffsetY As Long
Dim arrValues() As String, nValues As Long
Dim iSelectedVal As Long
Dim iScrollOffset As Long, nVisible As Long
Private Sub Form_DblClick()
    On Error Resume Next
    ReplaceCurrentWord GetSelectedText
    Err.Clear
End Sub
Private Sub Form_GotFocus()
    On Error Resume Next
    SetFocusToCodeWindow
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    ' Initialize the text offset
    TextOffsetY = (ITEM_HEIGHT - TextHeight("X")) \ 2
    Err.Clear
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' Hilight the selected item
    SelectItem (Y \ ITEM_HEIGHT) + iScrollOffset
    Err.Clear
End Sub
Public Sub SetSearchWord(sWord As String)
    On Error Resume Next
    Dim i As Long, n As Long
    Dim sPattern As String
    sPattern = sWord & "*"
    n = Len(sWord)
    ' This is a very slow way of checking
    ' This could definately be improved
    For i = 0 To nValues - 1
        If arrValues(i) Like sPattern Then
            SelectItem i
            Err.Clear
            Exit Sub
        End If
        Err.Clear
    Next
    SelectItem -1
    Err.Clear
End Sub
Public Function GetSelectedText() As String
    On Error Resume Next
    If iSelectedVal > -1 Then
        GetSelectedText = arrValues(iSelectedVal)
    End If
    Err.Clear
End Function
Public Sub HandleKeyUp()
    On Error Resume Next
    Select Case iSelectedVal
    Case -1
        SelectItem nValues - 1
    Case 0
    Case Else
        SelectItem iSelectedVal - 1
    End Select
    Err.Clear
End Sub
Public Sub HandleKeyDown()
    On Error Resume Next
    Select Case iSelectedVal
    Case -1
        SelectItem 0
    Case nValues - 1
    Case Else
        SelectItem iSelectedVal + 1
    End Select
    Err.Clear
End Sub
Private Sub Form_Paint()
    On Error Resume Next
    PaintPicture BackBuffer.Image, 0, 0
    Err.Clear
End Sub
Private Sub RedrawAll()
    On Error Resume Next
    If nValues = 0 Then Exit Sub
    Dim i As Long, j As Long
    With BackBuffer
        .Cls
        For i = 0 To nVisible - 1
            j = i + iScrollOffset
            If j = iSelectedVal Then
                BackBuffer.Line (20, ITEM_HEIGHT * i + 1)-Step(ScaleWidth - 20, ITEM_HEIGHT), &H8000000D, BF
                .CurrentX = 24
                .CurrentY = ITEM_HEIGHT * i + TextOffsetY + 1
                .ForeColor = vbWhite
                BackBuffer.Print arrValues(j)
            Else
                .CurrentX = 24
                .CurrentY = ITEM_HEIGHT * i + TextOffsetY + 1
                .ForeColor = vbBlack
                BackBuffer.Print arrValues(j)
            End If
            .PaintPicture Image1.Picture, 2, ITEM_HEIGHT * i + 2
            Err.Clear
        Next
        ' Draw 3d border lines [just like the vb popup window]
        BackBuffer.Line (0, 0)-(ScaleWidth, 0), &H8000000F
        BackBuffer.Line (0, 0)-(0, ScaleHeight), &H8000000F
        BackBuffer.Line (ScaleWidth - 2, 0)-Step(0, ScaleHeight), &H80000010
        BackBuffer.Line (0, ScaleHeight - 2)-Step(ScaleWidth, 0), &H80000010
        BackBuffer.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight), 0
        BackBuffer.Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0), 0
    End With
    Form_Paint
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    BackBuffer.Move 0, 0, ScaleWidth, ScaleHeight
    RedrawAll
    Err.Clear
End Sub
Public Sub SetListValues(sListValues As String)
    On Error Resume Next
    Dim i As Long, w As Long, tmp As Long
    arrValues = Split(Mid$(sListValues, 2), "|")
    nValues = UBound(arrValues) + 1
    iSelectedVal = -1
    ' Find the maximum width of the strings
    For i = 0 To nValues - 1
        tmp = TextWidth(arrValues(i))
        If tmp > w Then w = tmp
        Err.Clear
    Next
    ' Add padding and space for the icon, convert to twips
    Width = (w + 23 + 16) * Screen.TwipsPerPixelX
    ' Check if we need a scrollbar
    If nValues > MAX_SHOWN_ITEMS Then
        Height = (MAX_SHOWN_ITEMS * ITEM_HEIGHT + 3) * Screen.TwipsPerPixelY
        ' Increase width for the scrollbar
        Width = Width + 16 * Screen.TwipsPerPixelX
        ' Initialize the scrollbar parameters
        With VScroll
            .Move ScaleWidth - 18, 1, 16, ScaleHeight - 3
            .Min = 0
            .Max = nValues - MAX_SHOWN_ITEMS
            .SmallChange = 1
            .LargeChange = MAX_SHOWN_ITEMS
            .Value = 0
            .Visible = True
        End With
        nVisible = MAX_SHOWN_ITEMS
    Else
        Height = (nValues * ITEM_HEIGHT + 3) * Screen.TwipsPerPixelY
        VScroll.Visible = False
        nVisible = nValues
    End If
    SelectItem -1
    Err.Clear
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    CheckListBoxNeeded
    If Visible Then SetFocusToCodeWindow
    Err.Clear
End Sub
Private Sub VScroll_Change()
    On Error Resume Next
    iScrollOffset = VScroll.Value
    RedrawAll
    Err.Clear
End Sub
Private Sub VScroll_GotFocus()
    On Error Resume Next
    SetFocusToCodeWindow
    Err.Clear
End Sub
Private Sub VScroll_Scroll()
    On Error Resume Next
    iScrollOffset = VScroll.Value
    RedrawAll
    Err.Clear
End Sub
Private Sub SelectItem(ByVal idx As Long)
    On Error Resume Next
    iSelectedVal = idx
    If iSelectedVal > -1 Then
        If iSelectedVal < iScrollOffset Then iScrollOffset = iSelectedVal
        If iSelectedVal >= iScrollOffset + nVisible Then iScrollOffset = iSelectedVal - nVisible + 1
        If iScrollOffset < 0 Then iScrollOffset = 0
        VScroll.Value = iScrollOffset
    End If
    RedrawAll
    Err.Clear
End Sub
