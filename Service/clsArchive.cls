VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  FRS Jet      File Recovery and Archiving System Verion 1.3                         *
'*                                                                                     *
'*  Created:     January 11, 2006                                                      *
'*  Updated:     June 26, 2006                                                         *
'*  Purpose:     File Monitoring and Archiving Engine                                  *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.4.0                                                                 *
'*  Compile:     Native                                                                *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~ Exposed Functions ~/
'/~ Archive_Add             - add a file to an archive
'/~ Archive_Compress        - compress an archive
'/~ Archive_Create          - build an archive
'/~ Archive_Extract         - extract a file from an archive
'/~ Archive_Inflate         - decompress an archive
'/~ Archive_List            - return a list of files in an archive
'/~ Archive_Remove          - remove a file from an archive

'/~ Exposed Properties ~/
'/~ p_CompRatio             - compression ratio switch
'/~ p_SourceFile            - source file path
'/~ p_CompName              - compressed archive name
'/~ p_DecompName            - decompressed archive name
'/~ p_Compressed            - compression switch
'/~ p_Rebuild               - archive rebuild switch
'/~ p_Remove                - file removal switch

'/~ You may use this code in any way you like, as long as you agree that no warranty or
'/~ liability of any kind is expressed or implied. If you wish you can add a mention to the
'/~ about page of your project/app, ex. Credit to: John Underhill of NSPowertools.com.

'/~ This class could be transported and used in any application that requires
'/~ a compression/archive component. Api is NT only though, (NT4/W2K/XP/2003/Vista).

'/~ John

'/* file constants
Private Const SECTION_ALL_ACCESS As Long = &HF0000 Or _
    &H1 Or &H2 Or &H4 Or &H8 Or &H10

'/* compressed status
Private Enum eArchiveState
    NoArchive = 0
    Compressed = 1
    DeCompressed = 2
End Enum

'/* compression enums
Public Enum eCmpEngine
    cStandard = &H0
    cMaximum = &H100
    cHiber = &H200
End Enum

Public Enum eRatio
    cLow = 0
    cHigh = 1
End Enum

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                            ByVal dwDesiredAccess As Long, _
                                                                            ByVal dwShareMode As Long, _
                                                                            ByVal lpSecurityAttributes As Long, _
                                                                            ByVal dwCreationDisposition As Long, _
                                                                            ByVal dwFlagsAndAttributes As Long, _
                                                                            ByVal hTemplateFile As Long) As Long

Private Declare Function NtCreateSection Lib "ntdll.dll" (Handle As Long, _
                                                          ByVal DesiredAcess As Long, _
                                                          ObjectAttributes As Any, _
                                                          SectionSize As Any, _
                                                          ByVal Protect As Long, _
                                                          ByVal Attributes As Long, _
                                                          ByVal FileHandle As Long) As Long

Private Declare Function NtMapViewOfSection Lib "ntdll.dll" (ByVal Handle As Long, _
                                                             ByVal ProcessHandle As Long, _
                                                             BaseAddress As Long, _
                                                             ByVal ZeroBits As Long, _
                                                             ByVal CommitSize As Long, _
                                                             SectionOffset As Any, _
                                                             ViewSize As Long, _
                                                             ByVal InheritDisposition As Long, _
                                                             ByVal AllocaitonType As Long, _
                                                             ByVal Protect As Long) As Long

Private Declare Function NtUnmapViewOfSection Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                               ByVal Handle As Long) As Long

Private Declare Function NtClose Lib "ntdll.dll" (ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, _
                                                         lpFileSizeHigh As Long) As Long

Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, _
                                                            ByVal liDistanceToMove As Long, _
                                                            ByVal lpNewFilePointer As Long, _
                                                            ByVal dwMoveMethod As Long) As Long

Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                              BaseAddress As Long, _
                                                              regionsize As Long, _
                                                              ByVal FreeType As Long) As Long

Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal CompressionFormatAndEngine As Integer, _
                                                        ByVal UnCompressedBuffer As Long, _
                                                        ByVal UnCompressedBufferSize As Long, _
                                                        ByVal CompressedBuffer As Long, _
                                                        ByVal CompressedBufferSize As Long, _
                                                        ByVal UncompressedChunkSize As Long, _
                                                        FinalCompressedSize As Long, _
                                                        ByVal Workspace As Long) As Long

Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal CompressionFormat As Integer, _
                                                          ByVal UnCompressedBufferPtr As Long, _
                                                          ByVal UnCompressedBufferSize As Long, _
                                                          ByVal CompressedBuffer As Long, _
                                                          ByVal CompressedBufferSize As Long, _
                                                          FinalCompressedSize As Long) As Long

Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal CompressionFormatAndEngine As Integer, _
                                                                     CompressBufferWorkSpaceSize As Long, _
                                                                     CompressFragmentWorkSpaceSize As Long) As Long

Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                                  BaseAddress As Long, _
                                                                  ByVal ZeroBits As Long, _
                                                                  regionsize As Long, _
                                                                  ByVal AllocationType As Long, _
                                                                  ByVal Protect As Long) As Long

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, _
                                                                        ByVal lpNewFileName As String, _
                                                                        ByVal dwFlags As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

'/* archive progress max
Public Event eICompPMax(lMax As Long)
'/* archive progress event
Public Event eICompPTick(lCnt As Long)
'/* compression complete
Public Event eICompComplete()
'/* decompression complete
Public Event eIDecompComplete()
'/* error condition
Public Event eIErrorCond(sErr As String)

'/* class variables
Private m_lWorkspace                As Long
Private m_lFileHandle               As Long
Private m_lCompRatio                As Long
Private m_sSourceFile               As String
Private m_sCompName                 As String
Private m_sDecompName               As String
Private m_bCompressed               As Boolean
Private m_bRemoveFile               As Boolean
Private m_bRebuild                  As Boolean
Private c_CForward                  As Collection
Private c_CReturn                   As Collection

'***~ Properties ~***

'/* compression ratio switch
Public Property Get p_CompRatio() As eRatio
    p_CompRatio = m_lCompRatio
End Property

Public Property Let p_CompRatio(ByVal PropVal As eRatio)
    m_lCompRatio = PropVal
End Property

'/* source file path
Public Property Get p_SourceFile() As String
    p_SourceFile = m_sSourceFile
End Property

Public Property Let p_SourceFile(ByVal PropVal As String)
    m_sSourceFile = PropVal
End Property

'/* compressed file name
Public Property Get p_CompName() As String
    p_CompName = m_sCompName
End Property

Public Property Let p_CompName(ByVal PropVal As String)
    m_sCompName = PropVal
End Property

'/* decompressed file name
Public Property Get p_DecompName() As String
    p_DecompName = m_sDecompName
End Property

Public Property Let p_DecompName(ByVal PropVal As String)
    m_sDecompName = PropVal
End Property

'/* compression switch
Public Property Get p_Compressed() As Boolean
    p_Compressed = m_bCompressed
End Property

Public Property Let p_Compressed(ByVal PropVal As Boolean)
    m_bCompressed = PropVal
End Property

'/* archive rebuild switch
Public Property Get p_Rebuild() As Boolean
    p_Rebuild = m_bRebuild
End Property

Public Property Let p_Rebuild(ByVal PropVal As Boolean)
    m_bRebuild = PropVal
End Property

'/* file removal switch
Public Property Get p_Remove() As Boolean
    p_Remove = m_bRemoveFile
End Property

Public Property Let p_Remove(ByVal PropVal As Boolean)
    m_bRemoveFile = PropVal
End Property

'/* pass list collection into class
Public Property Get p_CForward() As Collection
    Set p_CForward = c_CForward
End Property

Public Property Set p_CForward(ByVal NewCollection As Collection)
    Set c_CForward = NewCollection
End Property

'/* return results collection to caller
Public Property Get p_CReturn() As Collection
    Set p_CReturn = c_CReturn
End Property

Public Property Set p_CReturn(ByVal NewCollection As Collection)
    Set c_CReturn = NewCollection
End Property

Private Sub Class_Initialize()

    Set p_CReturn = New Collection
    Set p_CForward = New Collection

End Sub

Public Function Archive_Add(ByVal sSourceFile As String) As Boolean
'/* add a file to archive

Dim vFile       As Variant

On Error GoTo Handler

    '/* pretest name and path
    If Len(m_sDecompName) = 0 Then GoTo Handler
    If c_CForward.Count = 0 And Len(sSourceFile) = 0 Then GoTo Handler
    
    '/* test entries for existence before processing
    '/* if only one file add to collection and test
    If Not c_CForward.Count = 0 Then
        Prebuild_Archive
    Else
        c_CForward.Add sSourceFile
        Prebuild_Archive
    End If
    
    '/* test archive state
    If Archive_State = Compressed Then
        '/* decompress
        Archive_Inflate
    End If
    
    '/* remove previous listing
    If m_bRebuild Then
        Rebuild_Archive sSourceFile
    End If
    
    '/* add file to archive
    Add_Record
        
    '/* compress archive
    If m_bCompressed Then
        Archive_Compress
    End If
    
    '/* remove original files
    If m_bRemoveFile Then
        For Each vFile In c_CForward
            Kill vFile
        Next vFile
    End If
    
    '/* reset resources
    Class_Cleanup
    '/* success
    Archive_Add = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Add" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_Compress() As Boolean
'/* compression hub

On Error GoTo Handler

    If Len(m_sDecompName) = 0 Or Len(m_sCompName) = 0 Then GoTo Handler
    If Not File_Exists(m_sDecompName) Then GoTo Handler
    If IsEmpty(m_lCompRatio) Then m_lCompRatio = cLow
    If FileLen(m_sDecompName) = 0 Then Exit Function
    If File_Exists(m_sCompName) Then Kill m_sCompName
    Compress_File m_sDecompName, m_sCompName, m_lCompRatio
    If m_bRemoveFile Then Kill m_sDecompName

    '/* success
    Archive_Compress = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Compress" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_Create(ByVal sSourceFile As String) As Boolean
'/* archive build control

Dim vFile   As Variant

On Error GoTo Handler

    '/* path and file check
    If Len(m_sCompName) = 0 Then GoTo Handler
    If c_CForward.Count = 0 And Len(sSourceFile) = 0 Then GoTo Handler
    
    '/* test entries for existence before processing
    '/* if only one file add to collection and test
    If Not c_CForward.Count = 0 Then
        Prebuild_Archive
    Else
        c_CForward.Add sSourceFile
        Prebuild_Archive
    End If
    
    '/* set decompressed name
    If Len(m_sDecompName) = 0 Then
        m_sDecompName = Left$(m_sCompName, Len(m_sCompName) - 3) + "dat"
    End If
    
    '/* build the archive
    Build_Archive
    '/* compress archive
    If m_bCompressed Then
        Archive_Compress
    End If
    
    '/* remove original files
    If m_bRemoveFile Then
        For Each vFile In c_CForward
            Kill vFile
        Next vFile
    End If
    
    '/* reset resources
    Class_Cleanup
    '/* success
    Archive_Create = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Create" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_Extract(ByVal sExtractFile As String, _
                                ByVal sDestFile As String) As Boolean
'/* extract a file from the archive

On Error GoTo Handler

    '/* source check
    If Len(sExtractFile) = 0 Or Len(sDestFile) = 0 Then GoTo Handler
    
    Select Case Archive_State
    Case Compressed
        '/* decompress
        Archive_Inflate
        '/* extract and restore
        Extract_File sExtractFile, sDestFile
    Case DeCompressed
        '/* extract and restore
        Extract_File sExtractFile, sDestFile
    Case Else
        GoTo Handler
    End Select

    '/* rebuild the archive
    If m_bRebuild Then
        Rebuild_Archive sExtractFile
    End If
    
    '/* reset resources
    Class_Cleanup
    '/* success
    Archive_Extract = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Extract" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_Inflate() As Boolean
'/* decompression hub

On Error GoTo Handler

    If Len(m_sCompName) = 0 Or Len(m_sDecompName) = 0 Then GoTo Handler
    If Not Archive_State = Compressed Then GoTo Handler
    If File_Exists(m_sDecompName) Then Kill m_sDecompName
    Decompress_File m_sCompName, m_sDecompName
    Kill m_sCompName
    
    '/* success
    Archive_Inflate = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Inflate" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_List() As Boolean
'/* list files in archive

On Error GoTo Handler

    '/* source check
    Select Case Archive_State
    Case Compressed
        '/* decompress
        Archive_Inflate
        '/* get file names
        List_Files
    Case DeCompressed
        '/* get file names
        List_Files
    Case Else
        GoTo Handler
    End Select
    
    '/* success
    Archive_List = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_List" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Archive_Remove(ByVal sExtractFile As String) As Boolean
'/* remove a file from archive

On Error GoTo Handler

    '/* validity check
    If Len(sExtractFile) = 0 Then GoTo Handler
    '/* state check
    If Archive_State = Compressed Then
        '/* decompress
        Archive_Inflate
    ElseIf Archive_State = NoArchive Then
        GoTo Handler
    End If
    
    '/* remove the file
    Rebuild_Archive sExtractFile

    '/* reset resources
    Class_Cleanup
    '/* success
    Archive_Remove = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Archive_Remove" + Chr$(30) + CStr(Err.Number))

End Function

Private Function Archive_State() As eArchiveState
'/* return current state of archive

    If File_Exists(m_sCompName) Then
        Archive_State = Compressed
    ElseIf File_Exists(m_sDecompName) Then
        Archive_State = DeCompressed
    Else
        Archive_State = NoArchive
    End If

End Function

Public Function Index_Append(ByVal sPath As String, _
                             ByVal sItem As String) As Boolean
'/* add an entry to the index

Dim iRec    As Integer

On Error GoTo Handler

    iRec = FreeFile
    Open sPath For Append As #iRec
    Print #iRec, sItem
    Close #iRec
    
    '/* success
    Index_Append = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Index_Append" + Chr$(30) + CStr(Err.Number))

End Function

Public Function Index_Extract(ByVal sPath As String, _
                              ByVal iItem As Integer) As Boolean
'/* remove an entry from index

Dim sLine       As String
Dim aLine()     As String
Dim i           As Integer
Dim sTemp       As String
Dim iRec        As Integer

On Error GoTo Handler

    iRec = FreeFile
    '/* name temp path
    sTemp = Left$(sPath, Len(sPath) - 3) & "tmp"
    '/* get file data
    Open sPath For Binary As #iRec
    sLine = Input$(LOF(iRec), iRec)
    Close #iRec
    '/* add to array
    aLine = Split(sLine, vbNewLine)
    '/* add all but item number
    Open sTemp For Output As #iRec
        For i = 0 To UBound(aLine)
            If Not i = iItem Then
                Print #iRec, aLine(i)
            End If
        Next i
    Close #iRec
    
    '/* cleanup
    Kill sPath
    File_Move sTemp, sPath
    Kill sTemp
    
    '/* success
    Index_Extract = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eIErrorCond("Index_Extract" + Chr$(30) + CStr(Err.Number))

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<                                  SUNDRY WORKER ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Database_Compact(ByVal sSource As String)
'/* needed because the compression/decompression adds
'/* spaces to the file

Dim sPath   As String
Dim sTmp    As String
Dim iRec    As Integer

On Error GoTo Handler

    iRec = FreeFile
    Open sSource For Binary As #iRec
        sTmp = Space$(LOF(iRec))
        Get #iRec, , sTmp
    Close #iRec
    
    sTmp = Mid$(sTmp, InStrRev(sTmp, Chr(0)))
    Kill sSource
    Open sSource For Output As #iRec
    Print #iRec, sTmp
    Close #iRec
    Exit Sub
    
Handler:
    
End Sub

Public Function File_Exists(ByVal sDir As String) As Boolean
'/* test file

Dim lRes    As Long
Dim sPath   As String

    sPath = String$(260, 0)
    lRes = GetShortPathName(sDir, sPath, 259)
    File_Exists = lRes > 0

End Function

Public Sub File_Move(ByVal sSource As String, _
                     ByVal sDestination As String)

    MoveFileEx sSource, sDestination, &H1

End Sub

Public Sub Make_Directory(ByVal sPath As String)

    MakeSureDirectoryPathExists sPath

End Sub

Public Sub Startup_Check(ByVal sPath As String)
'/* test and create paths

    If Not File_Exists(sPath) Then
        Make_Directory sPath
    End If

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<                                  ARCHIVE WORKER ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Add_Record()
'/* add a file to the archive

Dim vFile       As Variant
Dim lData       As Long
Dim iFName      As Integer
Dim lFData      As Long
Dim lFiles      As Long
Dim sTemp       As String
Dim sFName      As String
Dim sData       As String
Dim i           As Long
Dim sDate       As String
Dim iDate       As Integer
Dim iSize       As Integer
Dim sSize       As String
Dim iRec        As Integer
Dim iRc2        As Integer

On Error GoTo Handler

    iRec = FreeFile
    '/* test for decomp file
    If Not File_Exists(m_sDecompName) Then
        Open m_sDecompName For Append As #iRec
        Close #iRec
    End If

    Open m_sDecompName For Binary As #iRec
    Get #iRec, , lFiles
    sTemp = Left$(m_sDecompName, Len(m_sDecompName) - 3) & "tmp"
    '/* set number of files less 1
    iRc2 = FreeFile
    Open sTemp For Binary As #iRc2
    Put #iRc2, , lFiles + c_CForward.Count

    '/* write original records to temp
    For i = 1 To lFiles
        '/* read name
        Get #iRec, , iFName
        sFName = Space$(iFName)
        Get #iRec, , sFName
        '/* write name
        Put #iRc2, , iFName
        Put #iRc2, , sFName
        '/* read date
        Get #iRec, , iDate
        sDate = Space$(iDate)
        Get #iRec, , sDate
        '/* write date
        Put #iRc2, , iDate
        Put #iRc2, , sDate
        '/* read size
        Get #iRec, , iSize
        sSize = Space$(iSize)
        Get #iRec, , sSize
        '/* write size
        Put #iRc2, , iSize
        Put #iRc2, , sSize
        '/* read data
        Get #iRec, , lFData
        sData = Space$(lFData)
        Get #iRec, , sData
        '/* write file
        Put #iRc2, , lFData
        Put #iRc2, , sData
        DoEvents
    Next i
    Close #iRec

    '/* put new records in archive
    sDate = Format$(Now, ("dd/mmm/yy"))
    '/* add each file to archive
    For Each vFile In c_CForward
        '/* add filename
        iFName = Len(vFile)
        Put #iRc2, , iFName
        Put #iRc2, , CStr(vFile)
        '/* add date
        iDate = Len(sDate)
        Put #iRc2, , iDate
        Put #iRc2, , sDate
        '/* read data
        Open vFile For Binary As #iRec
        sData = Space$(LOF(iRec))
        Get #iRec, , sData
        Close #iRec
        '/* add file size
        lData = Len(sData)
        sSize = Len(sData)
        iSize = Len(sSize)
        Put #iRc2, , iSize
        Put #iRc2, , sSize
        '/* add file data
        Put #iRc2, , lData
        Put #iRc2, , sData
        Close #iRec
        DoEvents
    Next vFile
    Close #iRc2

    '/* kill archive and replace
    Kill m_sDecompName
    File_Move sTemp, m_sDecompName
    Kill sTemp

Exit Sub

Handler:
    Close #iRec
    Close #iRc2

End Sub

Private Sub Build_Archive()
'/* create a binary archive

Dim sPath       As String
Dim iFName      As Integer
Dim lData       As Long
Dim sData       As String
Dim iDate       As Integer
Dim sDate       As String
Dim iSize       As Integer
Dim sSize       As String
Dim lFCount     As Long
Dim vFile       As Variant
Dim lCntr       As Long
Dim iRec        As Integer
Dim iRc2        As Integer

On Error GoTo Handler

    iRec = FreeFile
    sPath = m_sDecompName
    sDate = Format$(Now, ("dd/mmm/yy"))
    '/* max progress
    RaiseEvent eICompPMax(c_CForward.Count)
    Open sPath For Binary As #iRec
    lFCount = c_CForward.Count
    '/* Set number of files
    Put #iRec, , lFCount
    '/* add each file to archive
    For Each vFile In c_CForward
        '/* add filename
        iFName = Len(vFile)
        Put #iRec, , iFName
        Put #iRec, , CStr(vFile)
        '/* add date
        iDate = Len(sDate)
        Put #iRec, , iDate
        Put #iRec, , sDate
        '/* read data
        iRc2 = FreeFile
        Open vFile For Binary As #iRc2
        sData = Space$(LOF(2))
        Get #iRc2, , sData
        Close #iRc2
        '/* add file size
        lData = Len(sData)
        sSize = Len(sData)
        iSize = Len(sSize)
        Put #iRec, , iSize
        Put #iRec, , sSize
        '/* add file data
        Put #iRec, , lData
        Put #iRec, , sData
        '/* progress event
        lCntr = lCntr + 1
        RaiseEvent eICompPTick(lCntr)
        DoEvents
    Next vFile
    Close #iRec

Exit Sub

Handler:

    Close #iRec
    Close #iRc2

End Sub

Private Sub Extract_File(ByVal sExtractFile As String, _
                         ByVal sDestName As String)

Dim iFName      As Integer
Dim lFData      As Long
Dim sData       As String
Dim sFName      As String
Dim lFiles      As Long
Dim i           As Long
Dim sDate       As String
Dim iDate       As Integer
Dim iSize       As Integer
Dim sSize       As String
Dim iRec        As Integer
Dim iRc2        As Integer

On Error GoTo Handler

    iRec = FreeFile
    '/* Open the archive
    Open m_sDecompName For Binary As #iRec
    '/* get number of files
    Get #iRec, , lFiles
    '/* read error
    If lFiles > 10000 Or lFiles < 1 Then GoTo Handler
    For i = 1 To lFiles
        '/* get file name
        Get #iRec, , iFName
        sFName = Space$(iFName)
        Get #iRec, , sFName
        '/* read date
        Get #iRec, , iDate
        sDate = Space$(iDate)
        Get #iRec, , sDate
        '/* read file size
        Get #iRec, , iSize
        sSize = Space$(iSize)
        Get #iRec, , sSize
        '/* read data
        Get #iRec, , lFData
        sData = Space$(lFData)
        Get #iRec, , sData
        '/* match and restore file
        If sFName = sExtractFile Then
            '/* if path not specified, use original
            If sDestName = "" Then
                sDestName = sFName
            End If
            '/* test for file presence
            If File_Exists(sDestName) Then Kill sDestName
            '/* restore file
            iRc2 = FreeFile
            Open sDestName For Output As #iRc2
            Print #iRc2, sData
            Close #iRc2
            Exit For
        End If
        DoEvents
    Next i
    Close #iRec

Exit Sub

Handler:
    Close #iRec
    Close #iRc2

End Sub

Private Sub List_Files()
'/* list files in archive

Dim iFName      As Integer
Dim sFName      As String
Dim sData       As String
Dim lFiles      As Long
Dim lFData      As Long
Dim i           As Long
Dim sDate       As String
Dim iDate       As Integer
Dim iSize       As Integer
Dim sSize       As String
Dim iRec        As Integer

On Error GoTo Handler

    iRec = FreeFile
    '/* Open the archive
    Open m_sDecompName For Binary As #iRec
    '/* get number of files
    Get #iRec, , lFiles
    '/* read error
    If lFiles > 10000 Or lFiles < 1 Then GoTo Handler
    For i = 1 To lFiles
        '/* get file name
        Get #iRec, , iFName
        sFName = Space$(iFName)
        Get #iRec, , sFName

        '/* read date
        Get #iRec, , iDate
        sDate = Space$(iDate)
        Get #iRec, , sDate
        '/* read size
        Get #iRec, , iSize
        sSize = Space$(iSize)
        Get #iRec, , sSize
        '/* pass names and dates to return list
        c_CReturn.Add sFName + Chr$(30) + sDate + Chr$(30) + sSize
        '/* read data
        Get #iRec, , lFData
        sData = Space$(lFData)
        Get #iRec, , sData
        DoEvents
    Next i
    Close #iRec

Exit Sub

Handler:
    Close #iRec

End Sub

Private Sub Rebuild_Archive(ByVal sExtractFile As String)

Dim iFName      As Integer
Dim lFData      As Long
Dim lFiles      As Long
Dim sTemp       As String
Dim sFName      As String
Dim sData       As String
Dim i           As Long
Dim sDate       As String
Dim iDate       As Integer
Dim iSize       As Integer
Dim sSize       As String
Dim bFound      As Boolean
Dim iRec        As Integer
Dim iRc2        As Integer

On Error GoTo Handler

    iRec = FreeFile
    Open m_sDecompName For Binary As #iRec
    Get #iRec, , lFiles
    '/* read error
    If lFiles > 10000 Or lFiles < 1 Then GoTo Handler
    '/* no files left
    If lFiles = 0 Then
        Kill m_sDecompName
        Exit Sub
    End If

    sTemp = Left$(m_sDecompName, Len(m_sDecompName) - 3) & "tmp"
    '/* set number of files less 1
    iRc2 = FreeFile
    Open sTemp For Binary As #iRc2
    Put #iRc2, , lFiles - 1

    For i = 1 To lFiles
        '/* read name
        Get #iRec, , iFName
        sFName = Space$(iFName)
        Get #iRec, , sFName
        '/* write name
        If Not sFName = sExtractFile Then
            Put #iRc2, , iFName
            Put #iRc2, , sFName
        Else
            bFound = True
        End If
        '/* read date
        Get #iRec, , iDate
        sDate = Space$(iDate)
        Get #iRec, , sDate
        '/* write date
        If Not sFName = sExtractFile Then
            Put #iRc2, , iDate
            Put #iRc2, , sDate
        End If
        '/* read size
        Get #iRec, , iSize
        sSize = Space$(iSize)
        Get #iRec, , sSize
        '/* write size
        If Not sFName = sExtractFile Then
            Put #iRc2, , iSize
            Put #iRc2, , sSize
        End If
        '/* read data
        Get #iRec, , lFData
        sData = Space$(lFData)
        Get #iRec, , sData
        '/* write file
        If Not sFName = sExtractFile Then
            Put #iRc2, , lFData
            Put #iRc2, , sData
        End If
        DoEvents
    Next i
    Close #iRc2
    Close #iRec

    '/* kill archive and replace
    If bFound Then
        Kill m_sDecompName
        If (lFiles - 1) > 0 Then
            File_Move sTemp, m_sDecompName
        End If
    End If
    Kill sTemp

Exit Sub

Handler:
    Close #iRec
    Close #iRc2

End Sub

Private Sub Prebuild_Archive()
'/* pre test archive file paths

Dim i As Long

On Error Resume Next

    With c_CForward
        For i = 1 To .Count
            If Not File_Exists(.Item(i)) Then
                .Remove i
            End If
        Next i
    End With

On Error GoTo 0

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<                                  COMPRESSION WORKER ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Compress_File(ByVal sSrce As String, _
                              ByVal sDest As String, _
                              ByVal Ratio As eRatio) As Long
'/* compression hub

Dim iCmpType        As eCmpEngine
Dim lFilePtr        As Long
Dim lFileSize       As Long
Dim lFileHandle     As Long
Dim lMHandle        As Long
Dim lFinalSize      As Long

On Error GoTo Handler

    m_lWorkspace = 0

    Select Case Ratio
        Case 0
            iCmpType = cStandard
        Case 1
            iCmpType = cMaximum
    End Select

    m_lWorkspace = Create_WorkSpace(iCmpType)
    lFilePtr = Open_File(sSrce, lFileSize, lFileHandle, lMHandle)
    Compress_Engine iCmpType, lFilePtr, lFileSize, lFinalSize, sDest, m_lWorkspace
    NtUnmapViewOfSection -1, lFilePtr
    NtClose lFileHandle
    NtClose lMHandle
    Compress_File = 1
    Exit Function
    
Handler:
Compress_File = 0

End Function

Public Function Decompress_File(ByVal sSrce As String, _
                                ByVal sDest As String) As Long
'/* decompression hub

Dim lFilePtr        As Long
Dim lFileSize       As Long
Dim lFileHandle     As Long
Dim lMHandle        As Long
Dim lFinalSize      As Long
Dim lStart          As Long

On Error GoTo Handler

    lFilePtr = Open_File(sSrce, lFileSize, lFileHandle, lMHandle)
    lStart = GetFileSize(lFileHandle, 0)
    DeCompress_Engine lFilePtr, lFileSize, lFinalSize, sDest
    NtUnmapViewOfSection -1, lFilePtr
    NtClose lFileHandle
    NtClose lMHandle
    Decompress_File = 1
    Exit Function
    
Handler:
Decompress_File = 0

End Function


Private Function Open_File(ByVal sFile As String, _
                           lSize As Long, _
                           lHandle As Long, _
                           lMemory As Long) As Long
'/* map a file to memory

Dim lBase   As Long

    lHandle = CreateFile(sFile, &H80000000 Or &H40000000, &H1 Or &H2, 0&, &H4, 0&, 0&)
    If lSize = 0 Then
        lSize = GetFileSize(lHandle, 0)
    End If
    NtCreateSection lMemory, SECTION_ALL_ACCESS, ByVal 0&, lSize, &H4, &H8000000, lHandle
    NtMapViewOfSection lMemory, -1, lBase, 0&, lSize, 0&, lSize, 1, 0, &H4
    Open_File = lBase
    m_lFileHandle = lHandle

End Function

Private Function Create_WorkSpace(eEngine As eCmpEngine) As Long
'/* create a workspace

Dim iCmpType        As Integer
Dim lWorkSize       As Long
Dim lFragSize       As Long

    iCmpType = &H2 Or eEngine
    RtlGetCompressionWorkSpaceSize iCmpType, lWorkSize, lFragSize
    NtAllocateVirtualMemory -1, Create_WorkSpace, 0, lWorkSize, &H1000, &H40

End Function

Private Function Compress_Engine(eEngine As eCmpEngine, _
                                 lCmpBuffer As Long, _
                                 ByVal lUnCmpBuffSize As Long, _
                                 lFinalSize As Long, _
                                 sNewFile As String, _
                                 ByVal lWorkspace As Long) As Long
'/* compression worker

Dim iCmpType            As Integer
Dim lTCmpBuffer         As Long
Dim lTCmpBufferSize     As Long
Dim lHCmpBuffer         As Long
Dim lTHCmpBuffer        As Long

    lTCmpBufferSize = lUnCmpBuffSize * 1.13 + 4
    lTCmpBuffer = Open_File(sNewFile, lTCmpBufferSize, lHCmpBuffer, lTHCmpBuffer)
    iCmpType = &H2 Or eEngine
    Compress_Engine = RtlCompressBuffer(iCmpType, lCmpBuffer, lUnCmpBuffSize, lTCmpBuffer, lTCmpBufferSize, 0&, lFinalSize, lWorkspace)
    NtUnmapViewOfSection -1, lTCmpBuffer
    NtClose lTHCmpBuffer
    SetFilePointer lHCmpBuffer, lFinalSize, 0, 0
    SetEndOfFile lHCmpBuffer
    NtClose lHCmpBuffer
    NtFreeVirtualMemory -1, lWorkspace, 0, &H4000
    CloseHandle m_lFileHandle
    m_lFileHandle = 0

End Function

Private Function DeCompress_Engine(lCmpBuffer As Long, _
                                   lCmpBuffSize As Long, _
                                   lFinalSize As Long, _
                                   sNewFile As String) As Long

'/* decompression worker

Dim lTUnCmpBuffer       As Long
Dim lTUnCmpBuffSize     As Long
Dim lHOBuffer           As Long
Dim lHTOBuffer          As Long

    lTUnCmpBuffSize = lCmpBuffSize * 12.5
    lTUnCmpBuffer = Open_File(sNewFile, lTUnCmpBuffSize, lHOBuffer, lHTOBuffer)
    DeCompress_Engine = RtlDecompressBuffer(&H2, lTUnCmpBuffer, lTUnCmpBuffSize, _
        lCmpBuffer, lCmpBuffSize, lFinalSize)
    NtUnmapViewOfSection -1, lTUnCmpBuffer
    NtClose lHTOBuffer
    SetFilePointer lHOBuffer, lFinalSize, 0, 0
    SetEndOfFile lHOBuffer
    NtClose lHOBuffer
    CloseHandle m_lFileHandle
    m_lFileHandle = 0

End Function

Private Sub Class_Cleanup()

    Set p_CForward = New Collection
    Set p_CReturn = New Collection

End Sub

Private Sub Class_Terminate()

    Set p_CReturn = Nothing
    Set p_CForward = Nothing

End Sub

