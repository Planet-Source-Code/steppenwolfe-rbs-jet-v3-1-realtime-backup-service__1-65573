VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***~ Lean and Mean Win32 ~***

'***************************************************************************************
'*  IFM V1.4     1.4 Indexing and File Monitoring class                                *
'*                                                                                     *
'*  Created:     Febuary 3, 2005                                                       *
'*  Updated:     June 03, 2006                                                         *
'*  Purpose:     Index Update Control Class                                            *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.4                                                                   *
'*  Compile:     PCode                                                                 *
'*  Referenced:  (Member RBS Jet)                                                      *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************


' ***~ Exposed Routines ~***

'/~ Engine_Init                     - load the indexing engine
'/~ Engine_End                      - stop and unload the service


' ***~ Registry Signal And Control Flags ~***

'/~ Control Flags {set externally}:
'/~ Master Flag                     - [mstflg | dword -(bool)]      flag signals a state change
'/~ Strict Termination              - [trmlib | dword -(bool)]      terminate - immediate service stop
'/~ Drive Paths                     - [drvpth | mltsz -(list)]      root of monitored drives
'/~ Filter Paths                    - [pthmon | mltsz -(list)]      list of filtered paths
'/~ Filter Files                    - [flemon | mltsz -(list)]      list of filtered files
'/~ Status Logging                  - [mdlogg | dword -(bool)]      log operational events
'/~ Archive Busy                    - [mdabsy | dword -(bool)]      the archive is engaged
'/~ Backup Timer                    - [bkptmr | dword -(long)]      backup scheduler switch
'/~ Backup Schedule                 - [bkpscd | string -(str)]      initial backup index

'/~ Signal Flags {state notifier}:
'/~ Index Active                    - [idxact | dword -(bool)]      we are alive
'/~ Write Active                    - [idxoft | dword -(bool)]      write is in commit phase


' ***~ Notes ~***

'/~ You may use this code in any way you like, as long as you agree that no warranty or
'/~ liability of any kind is expressed or implied. If you wish you can add a mention to the
'/~ about page of your project/app, ex. Credit to: John Underhill of NSPowertools.com.

'/~ Mandatory Settings:
'/~ c_IDriveList - list of drives to be monitored (registry 'drvpth' msz)

'/~ This one was a little tricky.. I think the key to making a successful directory monitor is that
'/~ it has to be both lean, and flexible. Two concepts that are often mutually exclusive, but I think I
'/~ struck a good balance here. There are a number of cascading filters that allow an early exit during the
'/~ directory change event, that, and using the smallest amount of processing within cycles with a one minute
'/~ interval update check, make the cpu footprint on this library almost negligible.
'/~ I have monitored four partitions simultaneously, while doing large installs, and no system lag
'/~ was evident. There are a series of properties that allow you to change drive path, file extension,
'/~ and file name dynamically, adding or removing them from the filter groups.

'/~ 98/ME - *-Absolutely Not Supported-*
'/~ They simply do not have the necessary libraries to support these operations. You could port this to
'/~ .Net, if you want to attempt backwards compliance, (good luck ;o).

'/~ 2K Server - *-Doubts-*
'/~ While researching the ReadDirectoryChangesW api on msdn, I saw a _vague_ note that alluded, async monitoring was
'/~ not available on 2K server, don't know if that's been patched, maybe someone can let me know? [email please].

'/~ This form of the index engine was heavily modified to work with RBS Jet, so there may still be a few loose
'/~ ends kicking around, I'll get to them in a couple of days. The original version of this was meant to monitor all
'/~ file changes across a maximum of 10 drives simultaneously. When complied as a service, this component has
'/~ performed very well, with almost no cpu footprint, and a very flexible filter mechanism. This could be converted
'/~ back to the directory monitor format with a minimal effort.

'/~ The monitor component properties are controlled by registry flags. A master flag is set by the controller
'/~ application, that signals a state change. The master flag is read every 60 cycles, so changes like
'/~ file additions, or a termination call, are processed on the nadir of a new loop cycle.
'/~ The directory monitor uses the asynchronous read method of ReadDirectoryChangesW api. In this way it
'/~ is possible to monitor up to 10 seperate paths, (drives), simultaneously. The completion key is hung
'/~ on an I/O completion port, using an overlapping directory structure to store results.
'/~ This class also uses a series of cascading filters in the Directory_Change routine, these
'/~ filters help to keep unncessary processing to a minimum, by filtering temp files, invalid entries,
'/~ and duplications, (an unusual quirk of the directory change api..).
'/~ The service element demonstrates the method for creating a true NT service. No ocx file, or form showing
'/~ up in the process list, but a genuine NT service.
'/~ When compiled with options, the service produces almost no cpu impact, can run unattended, and, after
'/~ the kinks are worked out, this could prove to be a handy little app..

'/~ Cheers
'/~ John - For a comment or a job: steppenwolfe_2000@yahoo.com


'* lets rock..
'*************************

'/* archive constants
Private Const COMP_NAME                                 As String = "\rbsarchive.rbc"
Private Const DECOMP_NAME                               As String = "\rbsarchive.rba"

'winnt.h
'FILE_NOTIFY_CHANGE_FILE_NAME                           0x00000001
'FILE_NOTIFY_CHANGE_DIR_NAME                            0x00000002
'FILE_NOTIFY_CHANGE_ATTRIBUTES                          0x00000004
'FILE_NOTIFY_CHANGE_SIZE                                0x00000008
'FILE_NOTIFY_CHANGE_LAST_WRITE                          0x00000010
'FILE_NOTIFY_CHANGE_LAST_ACCESS                         0x00000020
'FILE_NOTIFY_CHANGE_CREATION                            0x00000040
'FILE_NOTIFY_CHANGE_SECURITY                            0x00000100
Private Const FILE_CHANGE                               As Double = _
    &H1 Or &H2
Private Const FILE_MODIFY                               As Double = _
    &H1 Or &H2 Or &H8 Or &H10
Private Const FILE_ATTRCH                               As Double = _
    &H4 Or &H8 Or &H10 Or &H20 Or &H40
Private Const FILE_All                                  As Double = _
    &H1& Or &H2& Or &H4& Or &H8& Or &H10& Or &H20& Or &H40&

'/winnt.h
'FILE_ACTION_ADDED                                      0x00000001
'FILE_ACTION_REMOVED                                    0x00000002
'FILE_ACTION_MODIFIED                                   0x00000003
'FILE_ACTION_RENAMED_OLD_NAME                           0x00000004
'FILE_ACTION_RENAMED_NEW_NAME                           0x00000005
Private Const FILE_LIST                                 As Double = _
    &H1 Or &H2 Or &H4

'winbase.h
'FILE_FLAG_OVERLAPPED                                   0x40000000
'FILE_FLAG_BACKUP_SEMANTICS                             0x02000000
Private Const FILE_ATTRIB                               As Double = _
    &H2000000 Or &H40000000

Private Type OVERLAPPED
    Internal                                            As Long
    InternalHigh                                        As Long
    offset                                              As Long
    OffsetHigh                                          As Long
    hEvent                                              As Long
End Type

Private Type DIRECTORY_INFO
    hDir                                                As Long
    lpszDirName                                         As String * 256
    lpBuffer(4096)                                      As Byte
    dwBufLength                                         As Long
    oOverLapped                                         As OVERLAPPED
    lComplKey                                           As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength                                             As Long
    lpSecurityDescriptor                                As Long
    bInheritHandle                                      As Long
End Type

Private Type FT
    lLD                                                 As Long
    lHD                                                 As Long
End Type

Private Type BHFI
    lFA                                                 As Long
    fCT                                                 As FT
    fLA                                                 As FT
    fLWT                                                As FT
    lVSN                                                As Long
    lFSH                                                As Long
    lFSL                                                As Long
    lNOL                                                As Long
    lFIH                                                As Long
    lFIL                                                As Long
End Type

Private Enum eArchiveState
    NoArchive = 0
    Compressed = 1
    DeCompressed = 2
End Enum

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, _
                                                                     ByRef lpSource As Any, _
                                                                     ByVal iLen As Long)

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function PostQueuedCompletionStatus Lib "kernel32" (ByVal CompletionPort As Long, _
                                                                    lpNumberOfBytesTransferred As Long, _
                                                                    lpCompletionKey As Long, _
                                                                    lpOverlapped As Long) As Long

Private Declare Function GetQueuedCompletionStatus Lib "kernel32" (ByVal CompletionPort As Long, _
                                                                   lpNumberOfBytesTransferred As Long, _
                                                                   lpCompletionKey As Long, _
                                                                   lpOverlapped As OVERLAPPED, _
                                                                   ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                        ByVal dwDesiredAccess As Long, _
                                                                        ByVal dwShareMode As Long, _
                                                                        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                        ByVal dwCreationDisposition As Long, _
                                                                        ByVal dwFlagsAndAttributes As Long, _
                                                                        ByVal hTemplateFile As Long) As Long

Private Declare Function ReadDirectoryChangesW Lib "kernel32" (ByVal hDirectory As Long, _
                                                               lpBuffer As Any, _
                                                               ByVal nBufferLength As Long, _
                                                               ByVal bWatchSubtree As Long, _
                                                               ByVal dwNotifyFilter As Long, _
                                                               lpBytesReturned As Long, _
                                                               lpOverlapped As Any, _
                                                               lpCompletionRoutine As Any) As Long

Private Declare Function CreateIoCompletionPort Lib "kernel32" (ByVal FileHandle As Long, _
                                                                ByVal ExistingCompletionPort As Long, _
                                                                ByVal CompletionKey As Long, _
                                                                ByVal NumberOfConcurrentThreads As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long


Private m_bChange                                       As Boolean
Private m_bFiltered                                     As Boolean
Private m_bLoaded                                       As Boolean
Private m_bLogged                                       As Boolean
Private m_bModified                                     As Boolean
Private m_bPathFilter                                   As Boolean
Private m_bFileFilter                                   As Boolean
Private m_bReset                                        As Boolean
Private m_Cycles                                        As Long
Private m_lFile                                         As Long
Private m_lIoStatus                                     As Long
Private m_lMaxSize                                      As Long
Private m_lResult                                       As Long
Private m_lReturned                                     As Long
Private c_IDriveList                                    As Collection
Private c_WatchPath                                     As Collection
Private c_WatchIndex                                    As Collection
Private c_MatchFile                                     As Collection
Private c_HndlTracker                                   As Collection
Private c_FileFilter                                    As Collection
Private c_PathFilter                                    As Collection
Private c_TmpFilters                                    As Collection
Private m_tDirInfo(10)                                  As DIRECTORY_INFO
Private m_FileInfo                                      As BHFI
Private m_tOverlap                                      As OVERLAPPED
Private m_tSecAttr                                      As SECURITY_ATTRIBUTES
Private cArchive                                        As clsArchive

Public Property Get p_Modified() As Boolean
'/* monitor existing
    p_Modified = m_bModified
End Property

Public Property Let p_Modified(ByVal PropVal As Boolean)
    m_bModified = PropVal
End Property

Public Property Get p_Loaded() As Boolean
'/* running status
    p_Loaded = m_bLoaded
End Property

Public Property Let p_Loaded(ByVal PropVal As Boolean)
    m_bLoaded = PropVal
End Property

Public Property Get p_IDriveList() As Collection
    Set p_IDriveList = c_IDriveList
End Property

Public Property Set p_IDriveList(ByVal NewCollection As Collection)
    Set c_IDriveList = NewCollection
End Property


Private Sub Class_Initialize()
    Set cArchive = New clsArchive
End Sub

'> Control Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : Engine_Init
'* Purpose        : instantiate index engine
'* Inputs         : none
'* Outputs        : none
'*********************************************
Public Sub Engine_Init()
'/* start scanning

    '/* get app settings
    If Not Get_Settings Then
        '/* load failure - log and abort
        Log_Events "The Indexing service could not be started - invalid paramater."
        Exit Sub
    End If

    '/* load storage
    If Not m_bLoaded Then
        If Not Preload_Engine Then
            '/* load failure - log and abort
            Log_Events "The Indexing service could not be started - no drives selected."
            Exit Sub
        End If
    End If
    
    '/* log event
    If m_bLogged Then
        Log_Events "The Indexing service entered the Running State."
    End If
    
    '/* running
    m_bLoaded = True
    '/* set defaults
    Archive_Defaults
    '/* start monitoring engine
    Directory_Watch c_WatchPath

End Sub

'* Name           : Engine_End
'* Purpose        : unload the index engine
'* Inputs         : none
'* Outputs        : none
'*********************************************
Public Sub Engine_End()
'/* end indexing

On Error Resume Next

    '/* active flag
    cLightning.Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxact", 0
    
    '/* reset and unload storage
    Index_Reset
    
    '/* logging and state change
    If m_bLogged Then
        Log_Events "The Indexing service entered the Stop State."
    End If
    
    '/* stopped
    m_bLoaded = False
    
On Error GoTo 0

End Sub


'> Initialization Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function Preload_Engine() As Boolean
'/* initialize storage|classes

On Error GoTo Handler

    '/* init containers
    If Not Storage_Init Then
        '/* unload on failure
        Exit Function
    End If

    '/* success
    Preload_Engine = True

Handler:
On Error GoTo 0

End Function

Private Function Storage_Init() As Boolean
'/* initialize containers

Dim vI      As Variant

On Error GoTo Handler

    '/* get watched drives
    With cLightning
        If .Value_Exists(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth") Then
            If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth") Is Nothing Then
                Set p_IDriveList = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth")
            End If
        End If
    End With
    
    '/* no drives to monitor - terminate
    If c_IDriveList Is Nothing Then
        Engine_End
        Exit Function
    End If
    
    '/* initialize containers
    Set c_WatchPath = New Collection
    Set c_WatchIndex = New Collection
    Set c_HndlTracker = New Collection
    Set c_MatchFile = New Collection
    
    '/* add to drive list
    For Each vI In c_IDriveList
        c_WatchPath.Add Left$(CStr(vI), 3), CStr(Drive_Index(vI))
        c_WatchIndex.Add Drive_Index(vI), Left$(CStr(vI), 3)
    Next vI
    
    '/* success
    Storage_Init = True

Handler:
On Error GoTo 0

End Function

Private Function Drive_Index(ByVal vDrive As Variant) As Long
'/* get the drive index

Dim sDrives     As String
Dim aDrive()    As String
Dim lBuffer     As Long
Dim lCount      As Long

    '/* get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrives)
    '/* set string len
    sDrives = String$(lBuffer, 0)
    '/* get the drive list
    GetLogicalDriveStrings lBuffer, sDrives
    '/* split
    sDrives = LCase$(Left$(sDrives, Len(sDrives) - 2))
    aDrive = Split(sDrives, Chr$(0))
    For lCount = 0 To UBound(aDrive)
        If aDrive(lCount) = LCase$(vDrive) Then
            Drive_Index = lCount + 1
            Exit For
        End If
    Next lCount

End Function


'> Core Processors
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function Directory_Watch(ByRef c_WatchPath As Collection) As Boolean
'/* start monitoring using asynchronous callback
'/* into I/O completion port [ntsbl. pfm constructor]

Dim vI      As Variant
Dim lC      As Long

    If c_WatchPath.Count > 0 Then
        m_tSecAttr.nLength = Len(m_tSecAttr)
        '/* add watched paths
        For Each vI In c_WatchPath
            lC = CLng(c_WatchIndex.Item(CStr(vI)))
            Directory_Track CStr(vI), lC
            '/* create dir structure
            With m_tDirInfo(lC)
                .hDir = c_HndlTracker.Item(CStr(lC))
                .lpszDirName = vI
                .lComplKey = VarPtr(m_tDirInfo(lC))
                '/* create I/O port for async ops
                m_lIoStatus = CreateIoCompletionPort(.hDir, m_lIoStatus, .lComplKey, 0)
                '/* start monitoring
                m_lResult = ReadDirectoryChangesW(.hDir, .lpBuffer(0), _
                    4096, &H1, FILE_MODIFY, m_lReturned, .oOverLapped, ByVal 0&)
            End With
        Next vI
        '/* success
        If Not m_lResult = 0 Then
            Directory_Watch = True
        End If
    End If

End Function

Public Sub Directory_Change()
'/* store file changes

Dim bM      As Boolean
Dim bF()    As Byte
Dim iP      As Integer
Dim lK      As Long
Dim lL      As Long
Dim lO      As Long
Dim lp      As Long
Dim lS      As Long
Dim lT      As Long
Dim lX      As Long
Dim sE      As String
Dim sR      As String
Dim sT      As String
Dim sD      As String
Dim aR()    As String
Dim vI      As Variant

On Error Resume Next

    '/* load completed structure
    m_lResult = GetQueuedCompletionStatus(m_lIoStatus, m_lReturned, lK, m_tOverlap, 100)
    
    '/* get directory
    If Not m_lResult = 0 Then
        For iP = 1 To c_WatchPath.Count
            If m_tDirInfo(iP).lComplKey = lK Then
                Exit For
            End If
        Next iP
        
        Do
            '/* offset number
            With m_tDirInfo(iP)
                CopyMemory lO, .lpBuffer(0 + lX), 4
                '/* type number
                CopyMemory lT, .lpBuffer(4 + lX), 4
                '/* name length
                CopyMemory lL, .lpBuffer(8 + lX), 4
            End With
            
            '/* test for null value
            If lL = 0 Then
                lp = 0
                GoTo skip
            End If
            
            '/* byte array
            ReDim bF(0 To lL - 1) As Byte
            CopyMemory bF(0), m_tDirInfo(iP).lpBuffer(12 + lX), lL
            '/* convert to string
            sR = bF
            '/* our file path
            sR = Trim$(m_tDirInfo(iP).lpszDirName) & sR
            
            '/* filter 1: length filter
            If LenB(sR) < 4 Then GoTo skip
        
            '/* filter 2: proper file name
            If Not InStr(1, LCase$(sR), Chr$(46)) > 0 Then GoTo skip
            
            '/* filter 3: pass extension through temp filter
            sE = Mid$(LCase$(sR), InStrRev(sR, Chr$(46)) + 1)
            sT = c_TmpFilters.Item(sE)
            If LenB(sT) > 0 Then GoTo skip
                
            bM = False
            '/* filter 4: file filter
            If m_bFileFilter Then
                For Each vI In c_FileFilter
                    If LCase$(sR) = LCase$(vI) Then
                        bM = True
                        GoTo match
                    End If
                Next vI
            End If
            
            '<< good spot for a breakpoint
            '/* filter 5: path filter
            If m_bPathFilter Then
                For Each vI In c_PathFilter
                    aR = Split(LCase$(vI), Chr$(30))
                    If InStrB(1, LCase$(sR), aR(0)) > 0 Then
                        If Right$(LCase$(sR), 4) = aR(1) Then
                            bM = True
                        End If
                    End If
                Next vI
            End If
match:
            If bM Then
                '/ modified file
                If lT = 3 Then
                    '/* filter 6: duplication filter
                    sD = ""
                    sD = c_MatchFile.Item(sR)
                    If LenB(sD) = 0 Then
                        '/* add entry
                        c_MatchFile.Add sR, sR
                        '/* flag index write
                        m_bChange = True
                    End If
                End If
            End If
skip:
            '/* test for completion
            lX = lX + lO
            lp = m_tDirInfo(iP).lpBuffer(lX)
            '/* safe exit counter
            lS = lS + 1
            If lS > 500 Then
                lS = 0
                Exit Do
            End If
            '/*** IMPORTANT: unrem doevents if debugging ***/
            ' DoEvents
        Loop While (lp > 0)
        
        '/* test for changes
        m_lResult = ReadDirectoryChangesW(m_tDirInfo(iP).hDir, m_tDirInfo(iP).lpBuffer(0), _
            4096, &H1, FILE_MODIFY, m_lReturned, m_tDirInfo(iP).oOverLapped, ByVal 0&)
    End If
    
    '/* get flag changes/status every minute
    m_Cycles = m_Cycles + 1
    If m_Cycles = 60 Then
        m_Cycles = 0
        Operational_State
    End If

On Error GoTo 0

End Sub

Private Sub Directory_Track(ByVal sPath As String, _
                            ByVal lKey As Long)

'/* add path handle to tracker

    m_lFile = CreateFile(sPath, &H1, FILE_LIST, m_tSecAttr, &H3, FILE_ATTRIB, 0)
    c_HndlTracker.Add m_lFile, CStr(lKey)

End Sub


'> Status Tracking
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Operational_State()
'/* test the master flag for changes
   
    '/* master flag
    If cLightning.Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg") = 1 Then
        Flag_Check
    End If
    
    '/* process changes
    If m_bChange Then
        m_bChange = False
        Archive_Update
    End If

End Sub

Public Sub Flag_Check()
'/* act on change flags atmexc
    
    If cLightning.Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "trmlib") = 1 Then
        '/* reset flags
        Flag_Reset
        '/* stop and unload service
        mMain.EngineEnd
        Exit Sub
    End If

    '/* get changes
    Get_Settings
    Flag_Reset

End Sub

Public Function Get_Settings() As Boolean
'/* load settings

    '/* active list check
    If Not cLightning.Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth").Count = 0 Then
        m_bActivate = True
    End If
    '/* file filter
    m_bFileFilter = cLightning.Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "flemon") > 0
    If m_bFileFilter Then Filter_File
    '/* temp filters
    Filter_Ignore
    '/* path filter
    m_bPathFilter = cLightning.Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "pthmon") > 0
    If m_bPathFilter Then Filter_Path
    '/* logging
    m_bLogged = cLightning.Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "mdlogg") > 0
    '/* success
    Get_Settings = True

End Function

Private Sub Flag_Active(ByVal bActive As Boolean)
'/* load default flags

    With cLightning
        Select Case bActive
        Case True
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxact", 1
        Case Else
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxact", 0
        End Select
    End With

End Sub

Private Sub Flag_Commited(ByVal bActive As Boolean)
'/* write phase is active

    With cLightning
        Select Case bActive
        Case True
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxoft", 1
        Case Else
            .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxoft", 0
        End Select
    End With

End Sub

Private Sub Flag_Reset()
'/* reset of [local] flags on termination

    With cLightning
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "mstflg", 0
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "trmlib", 0
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxoft", 0
        .Write_DWord HKEY_LOCAL_MACHINE, m_sRegPath, "idxact", 0
    End With

End Sub

'> Filter Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Filter_Path()
'/* get monitored paths

On Error Resume Next

    With cLightning
        If .Value_Exists(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext") Then
            If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext") Is Nothing Then
                Set c_PathFilter = New Collection
                Set c_PathFilter = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltext")
            Else
                m_bPathFilter = False
            End If
        Else
            m_bPathFilter = False
        End If
    End With

On Error GoTo 0

End Sub

Private Sub Filter_File()
'/* fetch monitored files

    With cLightning
        '/* add file filter
        If .Value_Exists(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle") Then
            If Not .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle") Is Nothing Then
                Set c_FileFilter = New Collection
                Set c_FileFilter = .Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "fltfle")
            Else
                m_bFileFilter = False
            End If
        Else
            m_bFileFilter = False
        End If
    End With

End Sub

Private Sub Filter_Ignore()
'/* ignore temp files

    Set c_TmpFilters = New Collection
    With c_TmpFilters
        .Add 1, "---"
        .Add 1, "??~"
        .Add 1, "?~?"
        .Add 1, "$$$"
        .Add 1, "$db"
        .Add 1, "_mp"
        .Add 1, "~mp"
        .Add 1, "db$"
        .Add 1, "ol_"
        .Add 1, "bak"
        .Add 1, "btr"
        .Add 1, "chk"
        .Add 1, "data"
        .Add 1, "datap"
        .Add 1, "dd"
        .Add 1, "dir"
        .Add 1, "dmp"
        .Add 1, "evt"
        .Add 1, "ftg"
        .Add 1, "gid"
        .Add 1, "ini"
        .Add 1, "lnk"
        .Add 1, "log"
        .Add 1, "logt"
        .Add 1, "map"
        .Add 1, "mtx"
        .Add 1, "old"
        .Add 1, "pf"
        .Add 1, "pip"
        .Add 1, "prf"
        .Add 1, "temp"
        .Add 1, "tmp"
        .Add 1, "ver"
        .Add 1, "wpl"
    End With

End Sub

Private Function Filter_Duplicates(ByVal sItem As String, _
                                   ByVal iIndex As Integer) As Boolean

'/* I have seen duplications on every implementation of the api,
'/* for every add, 2 adds and 1 remove are produced,
'/* hence the need for this filter..

Dim vI      As Variant

On Error Resume Next

    If Not LenB(sItem) = 0 Then
        For Each vI In c_MatchFile
            If CStr(vI) = sItem Then
                Filter_Duplicates = True
                Exit For
            End If
        Next vI
    End If
    
On Error GoTo 0

End Function


'> Archive Engine
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Archive_Defaults()
'/* set archiver default paths

    With cArchive
        .p_CompName = App.Path + COMP_NAME
        .p_DecompName = App.Path + DECOMP_NAME
        .p_CompRatio = cLow
    End With

End Sub

Private Sub Archive_Update()
'/* get backup schedule

Dim vI          As Variant
Dim lInt        As Long
Dim sScd        As String
Dim lDif        As Long

    With cLightning
        '/* archive is busy
        If .Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "mdabsy") = 1 Then Exit Sub
        lInt = .Read_DWord(HKEY_LOCAL_MACHINE, m_sRegPath, "bkptmr")
        sScd = .Read_String(HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd")
    End With
    
    '/* default
    If lInt = 0 Then lInt = 1
    
    Select Case lInt
    '/* realtime
    Case 1
        Archive_Commit
    '/* 12 hr
    Case 2
        lDif = DateDiff("h", sScd, Now)
        If lDif > 12 Then
            Archive_Commit
            With cLightning
                .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd"
                .Write_String HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd", CStr(Now)
            End With
        Else
            '/* keep testing interval
            m_bChange = True
        End If
    '/* 24 hr
    Case 3
        lDif = DateDiff("h", sScd, Now)
        If lDif > 24 Then
            Archive_Commit
            With cLightning
                .Delete_Value HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd"
                .Write_String HKEY_LOCAL_MACHINE, m_sRegPath, "bkpscd", CStr(Now)
            End With
        Else
            m_bChange = True
        End If
    End Select

End Sub

Private Function Archive_Commit() As Boolean
'/* write the file to the archive

Dim vI      As Variant

    With cArchive
        .p_Rebuild = True
        For Each vI In c_MatchFile
            .Archive_Add CStr(vI)
        Next vI
        .p_Remove = True
        .Archive_Compress
    End With
    
    '/* reset storage
    Set c_MatchFile = New Collection

End Function

Private Function Archive_Compressed() As eArchiveState
'/* return current state of archive

    With cArchive
        If .File_Exists(App.Path + COMP_NAME) Then
            Archive_Compressed = Compressed
        ElseIf .File_Exists(App.Path + DECOMP_NAME) Then
            Archive_Compressed = DeCompressed
        Else
            Archive_Compressed = NoArchive
        End If
    End With

End Function

'> Support Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function File_Exists(ByVal sDir As String) As Boolean
'/* test file presence

Dim lR      As Long
Dim sP      As String

    sP = String$(260, 0)
    lR = GetShortPathName(sDir, sP, 259)
    File_Exists = lR > 0

End Function

'> Logging and Shutdown
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Index_Reset()
'/* reset monitoring engine

Dim vI      As Variant

On Error Resume Next

    '/* reset port status
    PostQueuedCompletionStatus m_lIoStatus, 0, 0, ByVal 0&
    
    '/* close directory handles
    For Each vI In c_HndlTracker
        CloseHandle CLng(vI)
    Next vI
    
    '/* close port
    CloseHandle m_lIoStatus
    m_lIoStatus = 0
    
    '/* reset storage
    Storage_Cleanup

On Error GoTo 0

End Sub

Private Sub Storage_Cleanup()
'/* empty all storage

On Error Resume Next
    
    '/* conditional reset
    If Not m_bReset Then
        Set c_IDriveList = Nothing
    End If
    
    '/* reset storage containers
    Set c_TmpFilters = Nothing
    Set c_PathFilter = Nothing
    Set c_FileFilter = Nothing
    Set c_MatchFile = Nothing
    Set c_WatchPath = Nothing
    Set c_WatchIndex = Nothing
    Set c_HndlTracker = Nothing

On Error GoTo 0

End Sub

Private Sub Log_Events(ByVal sEvent As String)
'/* log events

On Error Resume Next

    '/* write to log
    Open App.Path & "\hypidx.log" For Append As #1
    Print #1, CStr(Now) + vbTab + sEvent
    Close #1
    
On Error GoTo 0

End Sub

Private Sub Class_Terminate()
'/* cleanup if loaded

    If m_bLoaded Then
        Engine_End
        Set cArchive = Nothing
    End If

End Sub


