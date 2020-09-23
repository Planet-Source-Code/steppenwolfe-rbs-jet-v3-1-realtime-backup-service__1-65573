VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysIml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
Private Const MAX_PATH                      As Long = 260
Private Const FILE_ATTRIBUTE_NORMAL         As Long = &H80

Private Type SHFILEINFO
    hIcon                                   As Long
    iIcon                                   As Long
    dwAttributes                            As Long
    szDisplayName                           As String * MAX_PATH
    szTypeName                              As String * 80
End Type

Private Const SMICON_FLAGS                  As Long = _
    SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON

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

Public Property Get p_IconIndex(ByVal vIconKey As Variant) As Long

Dim lFlags          As Long
Dim lResult         As Long
Dim tFileInfo   As SHFILEINFO

    If IsNumeric(vIconKey) Then
        p_IconIndex = vIconKey
    Else
        lFlags = SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES
        lResult = SHGetFileInfo(vIconKey, FILE_ATTRIBUTE_NORMAL, tFileInfo, LenB(tFileInfo), lFlags)
        If Not lResult = 0 Then
            p_IconIndex = tFileInfo.iIcon
        End If
    End If
   
End Property
