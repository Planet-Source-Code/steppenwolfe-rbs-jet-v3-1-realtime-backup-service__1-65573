Attribute VB_Name = "mSundry"
Option Explicit

Private Const BIF_RETURNONLYFSDIRS          As Integer = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Integer = 2
Private Const MAX_PATH                      As Integer = 260
Private Const SHGFI_ICON                    As Long = &H100
Private Const SHGFI_DISPLAYNAME             As Long = &H200
Private Const SHGFI_TYPENAME                As Long = &H400
Private Const SHGFI_ATTRIBUTES              As Long = &H800
Private Const SHGFI_ICONLOCATION            As Long = &H1000
Private Const SHGFI_EXETYPE                 As Long = &H2000
Private Const SHGFI_SYSICONINDEX            As Long = &H4000
Private Const SHGFI_LINKOVERLAY             As Long = &H8000
Private Const SHGFI_SELECTED                As Long = &H10000
Private Const SHGFI_ATTR_SPECIFIED          As Long = &H20000
Private Const SHGFI_LARGEICON               As Long = &H0
Private Const SHGFI_SMALLICON               As Long = &H1
Private Const SHGFI_OPENICON                As Long = &H2
Private Const SHGFI_SHELLICONSIZE           As Long = &H4
Private Const SHGFI_PIDL                    As Long = &H8
Private Const SHGFI_USEFILEATTRIBUTES       As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL         As Long = &H80
Private Const SMICON_FLAGS                  As Long = _
    SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON

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

'/* large integer structs
Private Type ULong
    Byte1                                   As Byte
    Byte2                                   As Byte
    Byte3                                   As Byte
    Byte4                                   As Byte
End Type

Private Type LargeInt
    LoDWord                                 As ULong
    HiDWord                                 As ULong
    LoDWord2                                As ULong
    HiDWord2                                As ULong
End Type

Public Enum eMediaType
    Removable = 2
    HardDrive = 3
    Remote = 4
    CdRom = 5
    RamDisk = 6
    'DVD = 7 '<< someone want to check this for me?
End Enum

Public Type SHFILEINFO
    hIcon                                   As Long
    iIcon                                   As Long
    dwAttributes                            As Long
    szDisplayName                           As String * MAX_PATH
    szTypeName                              As String * 80
End Type

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                                                                                ByVal lpVolumeNameBuffer As String, _
                                                                                                ByVal nVolumeNameSize As Long, _
                                                                                                lpVolumeSerialNumber As Long, _
                                                                                                lpMaximumComponentLength As Long, _
                                                                                                lpFileSystemFlags As Long, _
                                                                                                ByVal lpFileSystemNameBuffer As String, _
                                                                                                ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                                                                        FreeBytesAvailableToCaller As LargeInt, _
                                                                                        TotalNumberOfBytes As LargeInt, _
                                                                                        TotalNumberOfFreeBytes As LargeInt) As Long

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long


Public Property Get p_ListHandle() As Long

Dim lHandle         As Long
Dim tFileInfo       As SHFILEINFO

    lHandle = SHGetFileInfo(".txt", FILE_ATTRIBUTE_NORMAL, tFileInfo, LenB(tFileInfo), SMICON_FLAGS)
    p_ListHandle = lHandle

End Property

Public Function IconIndex(ByVal vIconKey As Variant, _
                          ByRef tFileInfo As SHFILEINFO) As Long

Dim lFlags          As Long
Dim lResult         As Long

    If IsNumeric(vIconKey) Then
        IconIndex = vIconKey
    Else
        lFlags = SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICON
        lResult = SHGetFileInfo(vIconKey, FILE_ATTRIBUTE_NORMAL, tFileInfo, LenB(tFileInfo), lFlags)
        If Not lResult = 0 Then
            IconIndex = lResult
        End If
    End If
   
End Function

Public Function DriveList() As Collection
'/* list active drives

Dim lBuffer         As Long
Dim lCount          As Long
Dim sDrives         As String
Dim aDrives()       As String
Dim cTemp           As Collection

On Error Resume Next

    Set cTemp = New Collection
    
    '/* get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrives)
    
    '/* set string len
    sDrives = String$(lBuffer, 0)
    
    '/* get the drive list
    GetLogicalDriveStrings lBuffer, sDrives
    
    '/* split
    sDrives = Left$(sDrives, Len(sDrives) - 2)
    aDrives = Split(sDrives, Chr$(0))
    For lCount = 0 To UBound(aDrives)
        If Not LCase$(aDrives(lCount)) = "a:\" Then
            If MediaCheck(aDrives(lCount)) = HardDrive Then
                cTemp.Add aDrives(lCount)
            End If
        End If
    Next lCount
    
    '/* success
    If cTemp.Count > 0 Then
        Set DriveList = cTemp
    End If
    
On Error GoTo 0

End Function

Public Function DriveType(ByVal sPath As String) As String

'/* get drive file system

Dim lFlags          As Long
Dim lMaxLen         As Long
Dim lSerial         As Long
Dim sName           As String * 256
Dim sType           As String * 256
Dim sReturn         As String

On Error Resume Next

    '/* test and shorten string
    If Len(sPath) > 3 Then sPath = Left$(sPath, 3)
    
    '/* get volume flags
    GetVolumeInformation sPath, sName, Len(sName), lSerial, lMaxLen, lFlags, sType, Len(sType)
    sType = Left$(sType, InStr(1, sType, Chr$(0)) - 1)
    sReturn = Left$(sType, InStr(1, sType, Chr$(32)) - 1)
    
    '/* no value
    If LenB(sReturn) = 0 Then
        sReturn = "Unknown"
    End If
    DriveType = sReturn

On Error GoTo 0

End Function

Public Function DriveSize(ByVal sPath As String) As Collection

'/* get drive size|free space

Dim dFreeSpace      As Double
Dim dTotalSpace     As Double
Dim dUsedSpace      As Double
Dim tFreeBytes      As LargeInt
Dim tTotalBytes     As LargeInt
Dim tTotalFree      As LargeInt
Dim cTemp           As New Collection

On Error Resume Next

    Set cTemp = New Collection
    
    '/* fill the structures
    GetDiskFreeSpaceEx sPath, tFreeBytes, tTotalBytes, tTotalFree
    
    '/* free space
    With tFreeBytes
        dFreeSpace = LargeInteger(.HiDWord.Byte1, .HiDWord.Byte2, .HiDWord.Byte3, .HiDWord.Byte4) * 2 ^ 32 + _
            LargeInteger(.LoDWord.Byte1, .LoDWord.Byte2, .LoDWord.Byte3, .LoDWord.Byte4)
    End With
    
    '/* total space
    With tTotalBytes
        dTotalSpace = LargeInteger(.HiDWord.Byte1, .HiDWord.Byte2, .HiDWord.Byte3, .HiDWord.Byte4) * 2 ^ 32 + _
            LargeInteger(.LoDWord.Byte1, .LoDWord.Byte2, .LoDWord.Byte3, .LoDWord.Byte4)
    End With
    
    '/* format
    dUsedSpace = (dTotalSpace - dFreeSpace)
    dUsedSpace = ((dUsedSpace / 1024) / 1024)
    dFreeSpace = ((dFreeSpace / 1024) / 1024)
    dTotalSpace = (dTotalSpace / 1024) / 1024
    
    '/* return
    With cTemp
        .Add FormatNumber(dTotalSpace, 2)
        .Add FormatNumber(dFreeSpace, 2)
    End With
    
    Set DriveSize = cTemp
    
On Error GoTo 0

End Function

Public Function MediaCheck(ByVal sDrive As String) As eMediaType

    Select Case GetDriveType(sDrive)
        Case 2
            MediaCheck = Removable
        Case 3
            MediaCheck = HardDrive
        Case 4
            MediaCheck = Remote
        Case 5
            MediaCheck = CdRom
        Case 6
            MediaCheck = RamDisk
        Case 7
        '    MediaCheck = DVD
    End Select

End Function

Public Function LargeInteger(ByVal Byte1 As Byte, _
                             ByVal Byte2 As Byte, _
                             ByVal Byte3 As Byte, _
                             ByVal Byte4 As Byte) As Double

    LargeInteger = Byte4 * 2 ^ 24 + Byte3 * 2 ^ 16 + Byte2 * 2 ^ 8 + Byte1

End Function

Public Function FolderBrowse(ByVal sTitle As String, _
                             ByVal lHwnd As Long) As String
'/* standard folder browsing dialog

Dim lList           As Long
Dim sBuffer         As String
Dim tBrowseInfo     As BROWSEINFO

On Error Resume Next

    '/* fill struct
    With tBrowseInfo
        .hwndOwner = lHwnd
        .lpszTitle = lstrcat(sTitle, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lList = SHBrowseForFolder(tBrowseInfo)
    '/* call dialog
    If lList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lList, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Not Right$(sBuffer, 1) = Chr$(92) Then
            FolderBrowse = sBuffer + Chr$(92)
        Else
            FolderBrowse = sBuffer
        End If
    End If
    
On Error GoTo 0
    
End Function

