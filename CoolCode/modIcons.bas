Attribute VB_Name = "modIcons"
Option Explicit
Private Const MAX_PATH As Long = 260
Private Const ILD_TRANSPARENT As Long = &H1                     '  Display transparent
Private Const SHGFI_DISPLAYNAME As Long = &H200                 '  get display name
Private Const SHGFI_EXETYPE As Long = &H2000                    '  return exe type
'Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SHELLICONSIZE As Long = &H4                 '  get shell size icon
Private Const SHGFI_SMALLICON As Long = &H1                     '  get small icon
Private Const SHGFI_SYSICONINDEX As Long = &H4000                '  get system icon index
Private Const SHGFI_TYPENAME As Long = &H400                    '  get type name
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private IFileInfo As SHFILEINFO
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Const IFlags As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE 'Too stuffs, just put it in decs
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Sub tvExtractIcon(Filename As String, picIcon As PictureBox)
    On Error Resume Next
    Dim Icon As Long
    Icon = SHGetFileInfo(Filename, 0&, IFileInfo, Len(IFileInfo), IFlags Or SHGFI_SMALLICON)
    If Icon <> 0 Then
        Set picIcon.Picture = LoadPicture()
        Icon = ImageList_Draw(Icon, IFileInfo.iIcon, picIcon.hDC, 0, 0, ILD_TRANSPARENT)
    End If
    Err.Clear
End Sub
Public Function tvAddIconToIML(ByVal Filename As String, ByVal FType As String, imgList As ImageList, picIcon As PictureBox) As Long
    On Error Resume Next
    ' add an image of a file to an image list
    ' the file type is the extension of the file
    Dim i As Long
    Dim i_Tot As Long
    If IsNumeric(FType) Then
        FType = "XXX"
    End If
    If LCase$(FType) = "exe" Or LCase$(FType) = "ico" Then
        Call tvExtractIcon(Filename, picIcon)
        tvAddIconToIML = imgList.ListImages.Add(, , picIcon.Image).Index
    Else
        i_Tot = imgList.ListImages.Count
        For i = 1 To i_Tot
            If imgList.ListImages(i).Key = FType Then
                tvAddIconToIML = i
                Err.Clear
                Exit Function
            End If
            Err.Clear
        Next
        Call tvExtractIcon(Filename, picIcon)
        tvAddIconToIML = imgList.ListImages.Add(, FType, picIcon.Image).Index
    End If
    Err.Clear
End Function
