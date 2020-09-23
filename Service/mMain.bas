Attribute VB_Name = "mMain"
Option Explicit

'/~ original service component demo - Sergey Merzlikin (@codeguru)

Private Const SERVICE_NAME              As String = "nspidx"
Private Const INFINITE                  As Long = -1&
Private Const WAIT_TIMEOUT              As Long = 258&
Private Const VER_PLATFORM_WIN32_NT     As Long = 2&

Private Type OSVERSIONINFO
    dwOSVersionInfoSize                 As Long
    dwMajorVersion                      As Long
    dwMinorVersion                      As Long
    dwBuildNumber                       As Long
    dwPlatformId                        As Long
    szCSDVersion(1 To 128)              As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, _
                                                                      ByVal lpText As String, _
                                                                      ByVal lpCaption As String, _
                                                                      ByVal wType As Long) As Long

Private m_bNTService                    As Boolean
Private m_btServiceName()               As Byte
Private m_lSvcNamePtr                   As Long
Private m_lStopPending                  As Long
Private m_lStopEvent                    As Long
Private m_lStartEvent                   As Long
Private m_lSvcStatus                    As Long
Private m_tSvcState                     As SERVICE_STATUS
Public m_bActivate                      As Boolean
Public m_bTerminate                     As Boolean
Public m_sAppPath                       As String
Public m_sRegPath                       As String
Public cMonitor                         As clsMonitor
Public cLightning                       As clsLightning


Private Sub Main()
'/* service entry point

Dim lHandle         As Long
Dim lHObj(0 To 1)   As Long

    '/* one instance
    If App.PrevInstance Then Exit Sub
    
    '/* os check
    If Not IsWin32 Then
        MessageBox 0&, "This Application is only compatable with NT4/Windows 2000/XP or Vista operating systems.", _
            App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
        Exit Sub
    End If

    '/* create events
    m_lStopEvent = CreateEvent(0, 1, 0, vbNullString)
    m_lStopPending = CreateEvent(0, 1, 0, vbNullString)
    m_lStartEvent = CreateEvent(0, 1, 0, vbNullString)
    m_btServiceName = StrConv(SERVICE_NAME, vbFromUnicode)
    m_lSvcNamePtr = VarPtr(m_btServiceName(LBound(m_btServiceName)))
    
    '/* load events
    lHandle = StartAsService
    lHObj(0) = lHandle
    lHObj(1) = m_lStartEvent
    
    '/* start service
    m_bNTService = WaitForMultipleObjects(2&, lHObj(0), 0&, INFINITE) = 1&
    If Not m_bNTService Then
        CloseHandle lHandle
        '/*** IMPORTANT: rem the next 2 lines if debugging ***/
        MessageBox 0&, "This Application must be started as a Service.", _
            App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
        GoTo Handler
    End If
    
    '/* state change
    SetServiceState SERVICE_RUNNING
    '/* get settings
    EngineSettings
    '/* start monitor
    EngineStart
    
    '/* service loop
    Do
        '/* strict termination call
        If m_bTerminate Then Exit Do
        '/*** IMPORTANT: unrem doevents if debugging ***/
         DoEvents
        '/* check for file changes
        If m_bActivate Then
            cMonitor.Directory_Change
        Else
            '/* test monitor status
            EngineActive
        End If
    Loop While WaitForSingleObject(m_lStopPending, 1000&) = WAIT_TIMEOUT
    
    '/* unload indexer
    If Not m_bTerminate Then
        EngineEnd
    End If
    
    '/* set service state
    SetServiceState SERVICE_STOPPED
    '/* unload
    SetEvent m_lStopEvent
    WaitForSingleObject lHandle, INFINITE
    CloseHandle lHandle

Handler:
    '/* cleanup
    CloseHandle m_lStopEvent
    CloseHandle m_lStartEvent
    CloseHandle m_lStopPending

End Sub

Private Sub EngineActive()
'/* activate monitor

    If Not cLightning.Read_MultiCN(HKEY_LOCAL_MACHINE, m_sRegPath, "drvpth").Count = 0 Then
        m_bActivate = True
        '/* start the monitor
        cMonitor.Engine_Init
    End If
        
End Sub

Private Sub EngineSettings()
    
    With App
        m_sAppPath = .Path + Chr$(92)
        m_sRegPath = "Software\" + App.ProductName + "\Index"
    End With

End Sub

Private Sub EngineStart()
    
    '/* create class instances
    If cLightning Is Nothing Then
        Set cLightning = New clsLightning
    End If
    If cMonitor Is Nothing Then
        Set cMonitor = New clsMonitor
    End If

    '/* termination flag
    m_bTerminate = False

End Sub

Public Sub EngineEnd()
    
    If Not cMonitor Is Nothing Then
        '/* destroy class instance
        Set cMonitor = Nothing
    End If
    '/* destroy registry class
    If Not cLightning Is Nothing Then
        Set cLightning = Nothing
    End If
    '/* termination flag
    m_bTerminate = True

End Sub

Private Function IsWin32() As Boolean
'/* os ver check

Dim tOSVer As OSVERSIONINFO

    tOSVer.dwOSVersionInfoSize = LenB(tOSVer)
    GetVersionEx tOSVer
    IsWin32 = tOSVer.dwPlatformId = VER_PLATFORM_WIN32_NT

End Function

Private Function AddressPointer(ByVal lPointer As Long) As Long
'/* pointer function

    AddressPointer = lPointer

End Function

Private Function StartAsService() As Long
'/* launch dispatcher thread

Dim lThreadId As Long

    StartAsService = CreateThread(0&, 0&, AddressOf ServiceThread, 0&, 0&, lThreadId)

End Function

Private Sub ServiceThread(ByVal lDummy As Long)
'/* service thread init

Dim ServiceTableEntry As SERVICE_TABLE

    With ServiceTableEntry
        .lpServiceName = m_lSvcNamePtr
        .lpServiceProc = AddressPointer(AddressOf ServiceMain)
    End With
    StartServiceCtrlDispatcher ServiceTableEntry

End Sub

Private Sub ServiceMain(ByVal dwArgc As Long, _
                        ByVal lpszArgv As Long)
'/* service params

    With m_tSvcState
        .dwServiceType = SERVICE_WIN32_OWN_PROCESS
        .dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_SHUTDOWN
        .dwWin32ExitCode = 0&
        .dwServiceSpecificExitCode = 0&
        .dwCheckPoint = 0&
        .dwWaitHint = 0&
    End With
    
    m_lSvcStatus = RegisterServiceCtrlHandler(SERVICE_NAME, AddressOf ServiceHandler)
    SetServiceState SERVICE_START_PENDING
    SetEvent m_lStartEvent
    WaitForSingleObject m_lStopEvent, INFINITE

End Sub
   
Private Sub ServiceHandler(ByVal lControl As Long)
'/* shudown event

    Select Case lControl
        Case SERVICE_CONTROL_SHUTDOWN, SERVICE_CONTROL_STOP
            SetServiceState SERVICE_STOP_PENDING
            SetEvent m_lStopPending
        Case Else
            SetServiceState
    End Select

End Sub

Private Sub SetServiceState(Optional ByVal NewState As SERVICE_STATE = 0&)
'/* service status

    If NewState <> 0& Then m_tSvcState.dwCurrentState = NewState
    SetServiceStatus m_lSvcStatus, m_tSvcState

End Sub
