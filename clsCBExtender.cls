VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCBExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  ICBExtender  - ImageComboBox Extender Class                                        *
'*                                                                                     *
'*  Created:     May 24, 2006                                                          *
'*  Purpose:     ComboBox Extender Class                                               *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.1.0                                                                 *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*  Credits to:  M. Novaro for the Imagebox class extender (txtCodeId=34778),          *
'*               and Paul Caton, for the great self subclass routines                  *
'*                                                                                     *
'***************************************************************************************

' ~*** Notes ***~

'/~ Couldn't find a combobox with the nice features of the M$ version(Darn that M$!)
'/~ so here you go..
'/~ Added a directory in list browse, through root level folders.
'/~ Added a browse button to the list that invokes a standard dialog.
'/~ Just add a picturebox, and an imagelist to the form, and wire it up.
'/~ You could build on this.. add font, color picker, etc.. and if
'/~ you want to hack the dropdown list some more, you could do things like
'/~ expand the window dynamically, and make it so that all folders could be browsed,
'/~ doing away with the need for directory/drive list couplings.
'/~ Also, styles could be manipulated, scrollbar and button colors/shape..
'/~ It's up to you..

'/~ May 24, 2006 - Ver. 1.1
'/~ Applied a couple of fixes/improvements, and some basic error handling.
'/~ John Underhill (Steppenwolfe) steppenwolfe_2000@yahoo.com

'/~ Properties ~/
'/~ p_Picture                   - set the picturebox
'/~ p_ImageList                 - set the imagelist
'/~ p_AutoComplete              - autocomplete a phrase
'/~ p_AutoSelect                - select the entry
'/~ p_AutoSize                  - size list to largest entry
'/~ p_DirectoryList             - enable directory view
'/~ p_ItemSelected              -
'/~ p_KeyedDescent              - lower list with keystroke
'/~ p_ListWidth                 - set list start width
'/~ p_RestrictItems             - allow only list items

'/~ Routines ~/
'/~ SetReference                - attatch the combobox
'/~ ListDrives                  - select display type

Private Const CB_GETEDITSEL                 As Long = &H140
Private Const CB_LIMITTEXT                  As Long = &H141
Private Const CB_SETEDITSEL                 As Long = &H142
Private Const CB_ADDSTRING                  As Long = &H143
Private Const CB_DELETESTRING               As Long = &H144
Private Const CB_DIR                        As Long = &H145
Private Const CB_GETCOUNT                   As Long = &H146
Private Const CB_GETCURSEL                  As Long = &H147
Private Const CB_GETLBTEXT                  As Long = &H148
Private Const CB_GETLBTEXTLEN               As Long = &H149
Private Const CB_INSERTSTRING               As Long = &H14A
Private Const CB_RESETCONTENT               As Long = &H14B
Private Const CB_FINDSTRING                 As Long = &H14C
Private Const CB_SELECTSTRING               As Long = &H14D
Private Const CB_SETCURSEL                  As Long = &H14E
Private Const CB_SHOWDROPDOWN               As Long = &H14F
Private Const CB_GETITEMDATA                As Long = &H150
Private Const CB_SETITEMDATA                As Long = &H151
Private Const CB_GETDROPPEDCONTROLRECT      As Long = &H152
Private Const CB_SETITEMHEIGHT              As Long = &H153
Private Const CB_GETITEMHEIGHT              As Long = &H154
Private Const CB_SETEXTENDEDUI              As Long = &H155
Private Const CB_GETEXTENDEDUI              As Long = &H156
Private Const CB_GETDROPPEDSTATE            As Long = &H157
Private Const CB_FINDSTRINGEXACT            As Long = &H158
Private Const CB_SETLOCALE                  As Long = &H159
Private Const CB_GETTOPINDEX                As Long = &H15B
Private Const CB_SETTOPINDEX                As Long = &H15C
Private Const CB_GETHORIZONTALEXTENT        As Long = &H15D
Private Const CB_SETHORIZONTALEXTENT        As Long = &H15E
Private Const CB_GETDROPPEDWIDTH            As Long = &H15F
Private Const CB_SETDROPPEDWIDTH            As Long = &H160
Private Const CB_INITSTORAGE                As Long = &H161
Private Const BIF_RETURNONLYFSDIRS          As Integer = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Integer = 2
Private Const MAX_PATH                      As Integer = 260
Private Const BASIC_SHGFI_FLAGS             As Double = _
    &H4 Or &H200 Or &H400 Or &H2000 Or &H4000

Private Type BROWSEINFO
    hwndOwner                               As Long
    pidlRoot                                As Long
    pszDisplayName                          As Long
    lpszTitle                               As Long
    ulFlags                                 As Long
    lpfnCallback                            As Long
    lParam                                  As Long
    iImage                                  As Long
End Type

Private Type SHFILEINFO
    hIcon                                   As Long
    iIcon                                   As Long
    dwAttributes                            As Long
    szDisplayName                           As String * 260
    szTypeName                              As String * 80
End Type

Private Type POINTAPI
    X                                       As Long
    Y                                       As Long
End Type

Private Type RECT
    left                                    As Long
    tOp                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

Public Enum LIST_FORMAT
    Drives_Only = 0
    Special_Folders = 1
End Enum

Private Type SHITEMID
    cb                                      As Long
    abID                                    As Byte
End Type

Private Type ITEMIDLIST
    mkid                                    As SHITEMID
End Type

Public Enum eDirectories
    CSIDL_DESKTOPDIRECTORY = 0
    CSIDL_START_PROGRAMS = 2
    CSIDL_MYDOCUMENTS = 5
    CSIDL_FAVORITES = 6
    CSIDL_STARTUP = 7
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_START_MENU = 11
    CSIDL_MYMUSIC = 13
    CSIDL_MYVIDEO = 14
    CSIDL_DESKTOP = 16
    CSIDL_NETHOOD = 19
    CSIDL_FONTS = 20
    CSIDL_TEMPLATES = 21
    CSIDL_COMMON_STARTMENU = 22
    CSIDL_COMMON_PROGRAMS = 23
    CSIDL_COMMON_STARTUP = 24
    CSIDL_COMMON_DESKTOP = 25
    CSIDL_APPDATA = 26
    CSIDL_PRINTHOOD = 27
    CSIDL_SETTINGS_APPDATA = 28
    CSIDL_COMMON_FAVORITES = 31
    CSIDL_INTERNET_CACHE = 32
    CSIDL_COOKIES = 33
    CSIDL_HISTORY = 34
    CSIDL_COMMON_APPDATA = 35
    CSIDL_WINDOWS = 36
    CSIDL_SYSTEM = 37
    CSIDL_PROGRAM_FILES = 38
    CSIDL_MYPICTURES = 39
    CSIDL_PROFILE = 40
    CSIDL_COMMON_SYSTEM = 42
    CSIDL_COMMON_FILES = 43
    CSIDL_COMMON_TEMPLATES = 45
    CSIDL_COMMON_DOCUMENTS = 46
    CSIDL_COMMON_MUSIC = 53
    CSIDL_COMMON_PICTURES = 54
    CSIDL_COMMON_VIDEO = 55
    CSIDL_RESOURCES = 56
    CSIDL_CD_BURN_AREA = 56
End Enum

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                              ByVal wMsg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function GetVolumeInformation Lib "Kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                                                                                ByVal lpVolumeNameBuffer As String, _
                                                                                                ByVal nVolumeNameSize As Long, _
                                                                                                lpVolumeSerialNumber As Long, _
                                                                                                lpMaximumComponentLength As Long, _
                                                                                                lpFileSystemFlags As Long, _
                                                                                                ByVal lpFileSystemNameBuffer As String, _
                                                                                                ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hDCDest As Long, _
                                                            ByVal X As Long, _
                                                            ByVal Y As Long, _
                                                            ByVal Flags As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, _
                                                                       ByVal csidl As Long, _
                                                                       ByRef ppidl As ITEMIDLIST) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

Public Event eFilePath(sPath As String)

Private WithEvents m_oComboBox              As ImageCombo
Attribute m_oComboBox.VB_VarHelpID = -1
Private m_bAutoSize                         As Boolean
Private m_bAutoComplete                     As Boolean
Private blnAuto                             As Boolean
Private m_bRestrictItems                    As Boolean
Private m_bAutoSelect                       As Boolean
Private m_bExtended                         As Boolean
Private z_ScMem                             As Long
Private z_Sc(64)                            As Long
Private m_lMenuHandle                       As Long
Private m_lHighCount                        As Long
Private m_ExtDrive                          As String
Private m_sSearchPath                       As String
Private m_cSIcon                            As Collection
Private z_Funk                              As Collection
Private m_oImageList                        As Object
Private m_oPicture                          As Object
Private tShInfo                             As SHFILEINFO
Private m_cMem                              As Collection
Private m_eListMem                          As LIST_FORMAT


Public Property Get p_AutoComplete() As Boolean
    p_AutoComplete = m_bAutoComplete
End Property

Public Property Let p_AutoComplete(ByVal PropVal As Boolean)
    m_bAutoComplete = PropVal
End Property

Public Property Get p_AutoSelect() As Boolean
    p_AutoSelect = m_bAutoSelect
End Property

Public Property Let p_AutoSelect(ByVal PropVal As Boolean)
    m_bAutoSelect = PropVal
End Property

Public Property Get p_AutoSize() As Boolean
    p_AutoSize = m_bAutoSize
End Property

Public Property Let p_AutoSize(ByVal PropVal As Boolean)
    m_bAutoSize = PropVal
End Property

Public Property Get p_ImageList() As Object
    Set p_ImageList = m_oImageList
End Property

Public Property Set p_ImageList(ByVal PropVal As Object)
    Set m_oImageList = PropVal
    Set m_oComboBox.ImageList = m_oImageList
End Property

Public Property Get p_KeyedDescent() As Boolean
    p_KeyedDescent = (SendMessageLong(m_oComboBox.hwnd, CB_GETEXTENDEDUI, 0, 0) <> 0)
End Property

Public Property Let p_KeyedDescent(ByVal PropVal As Boolean)
    SendMessageLong m_oComboBox.hwnd, CB_SETEXTENDEDUI, Abs(PropVal), 0
End Property

Private Property Get p_ListWidth() As Long
    p_ListWidth = SendMessage(m_oComboBox.hwnd, CB_GETDROPPEDWIDTH, 0, ByVal 0)
End Property

Private Property Let p_ListWidth(PropVal As Long)
    SendMessage m_oComboBox.hwnd, CB_SETDROPPEDWIDTH, PropVal, ByVal 0
End Property

Public Property Get p_Picture() As Object
    Set p_Picture = m_oPicture
End Property

Public Property Set p_Picture(ByVal PropVal As Object)
    Set m_oPicture = PropVal
End Property

Public Property Get p_RestrictItems() As Boolean
    p_RestrictItems = m_bRestrictItems
End Property

Public Property Let p_RestrictItems(ByVal PropVal As Boolean)
    m_bRestrictItems = PropVal
End Property


Private Sub Class_Initialize()
    Set m_cMem = New Collection
End Sub

Private Sub BrowseForFolder()
'/* standard folder browsing dialog

Dim lList           As Long
Dim sTitle          As String
Dim tBrowseInfo     As BROWSEINFO
Dim sBuffer         As String
Dim sKey            As String

On Error Resume Next

    If m_lHighCount > 0 Then
        ComboRemoveItem m_lHighCount
        m_lHighCount = 0
    End If
    
    '/* title
    sTitle = "Select a Directory to Scan: "
    '/* fill struct
    With tBrowseInfo
        .hwndOwner = m_oComboBox.Container.hwnd
        .lpszTitle = lstrcat(sTitle, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lList = SHBrowseForFolder(tBrowseInfo)
    '/* call dialog
    If lList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lList, sBuffer
        sBuffer = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        sKey = Get_SmallIcon(sBuffer)
        ComboAddItem sBuffer, sKey
        With m_oComboBox
            m_lHighCount = .ComboItems.Count
            .ComboItems.Item(m_lHighCount).Selected = True
            m_sSearchPath = sBuffer
            RaiseEvent eFilePath(m_sSearchPath)
            .Refresh
        End With
    End If
    
On Error GoTo 0
    
End Sub

Private Sub ComboAddItem(ByVal sItem As String, _
                        ByVal sKey As String, _
                        Optional ByVal lIndex As Long, _
                        Optional lIndent As Long)

'/* add an item

    If lIndex > 0 Then
        m_oComboBox.ComboItems.Add lIndex, , sItem, sKey, , lIndent
    Else
        m_oComboBox.ComboItems.Add , , sItem, sKey, , 1
    End If

End Sub

Private Sub ComboRemoveItem(ByVal lIndex As Long)
'/* remove an item

    m_oComboBox.ComboItems.Remove lIndex

End Sub

Public Function Directory_Exists(ByVal sDir As String) As Boolean
'/* can test any, file, any directory..

Dim lRes                As Long
Dim sPath               As String

    '/* prep var
    sPath = String$(255, 0)
    '/* get short path
    lRes = GetShortPathName(sDir, sPath, 254)
    '/* success
    Directory_Exists = lRes > 0

End Function

Private Function Drive_Check() As Collection
'/* get drive strings

Dim lBuffer     As Long
Dim lCount      As Long
Dim sDrives     As String
Dim aDrives()   As String
Dim cTemp       As Collection

On Error Resume Next

    Set cTemp = New Collection
    '//get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrives)
    '//set string len
    sDrives = String$(lBuffer, 0)
    '//get the drive list
    GetLogicalDriveStrings lBuffer, sDrives
    '/* split
    sDrives = left(sDrives, Len(sDrives) - 2)
    aDrives = Split(sDrives, vbNullChar)
    
    '/* loop through drives and add info
    For lCount = 0 To UBound(aDrives)
        If Not LCase$(aDrives(lCount)) = "a:\" Then
            cTemp.Add aDrives(lCount) + Drive_Label(aDrives(lCount))
        End If
    Next lCount
    
    '/* success
    If cTemp.Count > 0 Then
        Set Drive_Check = cTemp
    End If

On Error GoTo 0

End Function

Private Function Drive_Label(ByVal sPath As String) As String
'/* get drive label and fs type

Dim lFlags          As Long
Dim lMaxLen         As Long
Dim lSerial         As Long
Dim sName           As String * 256
Dim sType           As String * 256
Dim sTemp           As String

On Error Resume Next

    '/* test and shorten string
    If Len(sPath) > 3 Then
        sPath = left(sPath, 3)
    End If
    '/* get volume flags
    GetVolumeInformation sPath, sName, Len(sName), lSerial, lMaxLen, lFlags, sType, Len(sType)
    
    sTemp = left(sName, InStr(1, sName, vbNullChar) - 1)
    If Len(sTemp) > 0 Then
        Drive_Label = sTemp
    End If

    sTemp = left(sType, InStr(1, sType, vbNullChar) - 1)
    If Len(sTemp) > 0 Then
        If Len(Drive_Label) > 0 Then
            Drive_Label = "  [" & Drive_Label & " - " & sTemp & "]"
        Else
            Drive_Label = "  [" & sTemp & "]"
        End If
    End If

On Error GoTo 0

End Function

Public Function Get_Folder(SHFlag As eDirectories) As String
'/* get default folder locations

Dim lRes            As Long
Dim sPath           As String
Dim ItemIdL         As ITEMIDLIST

On Error GoTo Handler

    lRes = SHGetSpecialFolderLocation(100, SHFlag, ItemIdL)

    If lRes Then
        Get_Folder = vbNullString
    Else
        sPath = Space$(512)
        lRes = SHGetPathFromIDList(ByVal ItemIdL.mkid.cb, ByVal sPath)
        Get_Folder = left(sPath, InStr(sPath, vbNullChar) - 1)
    End If

Handler:
On Error GoTo 0

End Function

Private Function Get_SmallIcon(ByVal sFile As String) As String
'/* names image keys with file extensions,
'/* for reusable image items in imagelist

Dim hSIcon      As Long
Dim imgObj      As ListImage
Dim sKey        As String

On Error Resume Next

    '/* get associated extension
    '/* get handle to icon
    hSIcon = SHGetFileInfo(sFile, 0&, tShInfo, Len(tShInfo), BASIC_SHGFI_FLAGS Or &H1)
    '/* load icon to picturebox
    If Not hSIcon = 0 Then
        With m_oPicture
            Set .Picture = LoadPicture("")
            .AutoRedraw = True
            ImageList_Draw hSIcon, tShInfo.iIcon, .hdc, 0&, 0&, &H1
            .Refresh
        End With
        '/* test for icon presence in collection
        sKey = m_cSIcon.Item(sFile)
        '/* if not present, add to collection
        If LenB(sKey) = 0 Then
            m_cSIcon.Add 1, sFile
            '/* add icon to image list
            '/* use file extension as image key
            Set imgObj = m_oImageList.ListImages.Add(Key:=sFile, Picture:=m_oPicture.Image)
        End If
        Get_SmallIcon = sFile
    Else
        '/* no icon, use default image
        Get_SmallIcon = "dft"
    End If

On Error GoTo 0

End Function

Public Sub ListDrives(Optional ByVal eFormat As LIST_FORMAT = Drives_Only)
'/* select the items displayed (change to your liking)

Dim vItem       As Variant
Dim sKey        As String
Dim sPath       As String

    m_eListMem = eFormat
    m_oComboBox.ComboItems.Clear
    
    Select Case eFormat
    Case 0
        ComboAddItem "Browse..", "dft"
        For Each vItem In Drive_Check
            sKey = Get_SmallIcon(left(CStr(vItem), 3))
            ComboAddItem vItem, sKey
        Next vItem

    Case 1
        '/* drive list
        For Each vItem In Drive_Check
            sKey = Get_SmallIcon(left(CStr(vItem), 3))
            ComboAddItem vItem, sKey
        Next vItem
        
        ComboAddItem "Browse..", "dft"
        '/* docs
        sPath = Get_Folder(CSIDL_MYDOCUMENTS)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "My Documents", sKey
        End If
        '/* music
        sPath = Get_Folder(CSIDL_MYMUSIC)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "My Music", sKey
        End If
        '/* pictures
        sPath = Get_Folder(CSIDL_MYPICTURES)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "My Pictures", sKey
        End If
        '/* video
        sPath = Get_Folder(CSIDL_MYVIDEO)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "My Video", sKey
        End If
        '/* profile
        sPath = Get_Folder(CSIDL_PROFILE)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "Profile", sKey
        End If
        '/* programs
        sPath = Get_Folder(CSIDL_PROGRAM_FILES)
        If Directory_Exists(sPath) Then
            sKey = Get_SmallIcon(sPath)
            ComboAddItem "Program Files", sKey
        End If
    End Select
    
    '/* select the first item
    m_oComboBox.ComboItems.Item(1).Selected = True
    
End Sub

Private Sub m_oComboBox_Click()
'/* raise event with selected path

Dim sPath   As String

On Error Resume Next

    '/* safety valve
    If m_oComboBox.SelectedItem Is Nothing Then
        m_oComboBox.ComboItems.Clear
        ListDrives m_eListMem
        m_oComboBox.ComboItems.Item(1).Selected = True
        m_oComboBox.Container.Refresh
        Exit Sub
    End If
    
    '/* add folders
    Select Case m_oComboBox.SelectedItem.Text
    Case "Browse.."
        BrowseForFolder
    Case "My Documents"
        sPath = Get_Folder(CSIDL_MYDOCUMENTS) + Chr$(92)
    Case "My Music"
        sPath = Get_Folder(CSIDL_MYMUSIC)
    Case "My Pictures"
        sPath = Get_Folder(CSIDL_MYPICTURES)
    Case "My Video"
        sPath = Get_Folder(CSIDL_MYVIDEO)
    Case "Profile"
        sPath = Get_Folder(CSIDL_PROFILE)
    Case "Program Files"
        sPath = Get_Folder(CSIDL_PROGRAM_FILES)
    Case Else
        If InStr(1, m_oComboBox.SelectedItem.Text, Chr$(58)) > 0 Then
            sPath = left$(m_oComboBox.SelectedItem.Text, 2) + Chr$(92)
        Else
            sPath = m_ExtDrive + m_oComboBox.SelectedItem.Text
        End If
    End Select
    
    m_oComboBox.Container.Refresh
    m_sSearchPath = sPath
    RaiseEvent eFilePath(m_sSearchPath)

On Error GoTo 0

End Sub

Private Sub m_oComboBox_GotFocus()
'/* resize entry

    If m_bAutoSelect Then
        m_oComboBox.SelStart = 0
        m_oComboBox.SelLength = Len(m_oComboBox.Text)
    End If

End Sub

Private Sub m_oComboBox_KeyPress(KeyAscii As Integer)
'/* auto complete

Dim found       As Boolean
Dim sPart       As String
Dim sTemp       As String
Dim bFound      As Boolean
Dim sComp       As String
Dim cbItem      As ComboItem

On Error GoTo Handler

    If Not (m_bAutoComplete Or m_bRestrictItems) Then Exit Sub
    Select Case KeyAscii
        '/* backspace - erase
    Case vbKeyBack
        If m_oComboBox.SelStart <= 1 Then
            m_oComboBox.Text = ""
            KeyAscii = 0
            Exit Sub
        End If
        If Not m_oComboBox.SelLength = 0 Then
            sPart = left$(m_oComboBox.Text, m_oComboBox.SelStart - 1)
        End If
        '/* enter
    Case vbKeyReturn, vbKeyTab
        m_oComboBox.Refresh
        Exit Sub
        '/* chars
    Case 32 Or (KeyAscii > 127 And KeyAscii < 161)
        Exit Sub
    Case Else
        sPart = left$(m_oComboBox.Text, m_oComboBox.SelStart) + Chr$(KeyAscii)
    End Select
    
    '/* loop though list and
    '/* get best match
    If m_bAutoComplete Or m_bRestrictItems Then
        For Each cbItem In m_oComboBox.ComboItems
            sTemp = left$(cbItem.Text, Len(sPart))
            If LCase$(sTemp) Like LCase$(sPart) Then
                If m_bAutoComplete Then
                    With m_oComboBox
                        Set .SelectedItem = cbItem
                        .SelStart = Len(sPart)
                        .SelLength = Len(.Text) - .SelStart
                    End With
                    KeyAscii = 0
                End If
                bFound = True
                Exit For
            End If
        Next cbItem
    End If
    
    '/* item restrict
    If Not bFound Then
        If m_bRestrictItems Then
            KeyAscii = 0
        End If
    Else
       m_oComboBox_Click
    End If
    blnAuto = False

Handler:
On Error GoTo 0

End Sub

Private Sub SizeComboList()
'/* set list size to contents

Dim cbItem          As ComboItem
Dim sText           As String
Dim lHdc            As Long
Dim lLength         As Long
Dim tPnt            As POINTAPI

    lHdc = GetDC(m_oComboBox.Container.hwnd)

    For Each cbItem In m_oComboBox.ComboItems
        sText = cbItem.Text
        GetTextExtentPoint32 lHdc, sText, Len(sText) + 2, tPnt
        If tPnt.X > lLength Then
            lLength = tPnt.X
        End If
    Next cbItem

    lLength = lLength + 64
    SendMessage m_oComboBox.hwnd, CB_SETDROPPEDWIDTH, lLength, ByVal 0

End Sub

Public Sub SetReference(oCombo As ImageCombo)
'/* attatch the class

    Set m_oComboBox = oCombo
    
End Sub

Private Sub m_oComboBox_Dropdown()

    '/* auto size contents
    If m_bAutoSize Then
        SizeComboList
    End If
    
End Sub


