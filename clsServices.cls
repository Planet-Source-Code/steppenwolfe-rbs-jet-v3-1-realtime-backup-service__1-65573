VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SVC_NAME                                  As String = "rbsidx.exe"
Private Const ERROR_MORE_DATA                           As Integer = 234
Private Const SC_MANAGER_ENUMERATE_SERVICE              As Long = &H4
Private Const SERVICE_STATE_ALL                         As Long = &H3
Private Const SERVICE_KERNEL_DRIVER                     As Long = &H1
Private Const SERVICE_FILE_SYSTEM_DRIVER                As Long = &H2
Private Const SERVICE_ADAPTER                           As Long = &H4
Private Const SERVICE_RECOGNIZER_DRIVER                 As Long = &H8
Private Const SERVICE_WIN32_OWN_PROCESS                 As Long = &H10
Private Const SERVICE_WIN32_SHARE_PROCESS               As Long = &H20
Private Const SERVICE_INTERACTIVE_PROCESS               As Long = &H100
Private Const SERVICE_STOPPED                           As Long = &H1
Private Const SERVICE_START_PENDING                     As Long = &H2
Private Const SERVICE_STOP_PENDING                      As Long = &H3
Private Const SERVICE_RUNNING                           As Long = &H4
Private Const SERVICE_CONTINUE_PENDING                  As Long = &H5
Private Const SERVICE_PAUSE_PENDING                     As Long = &H6
Private Const SERVICE_PAUSED                            As Long = &H7
Private Const SC_MANAGER_CONNECT                        As Long = &H1
Private Const SC_MANAGER_CREATE_SERVICE                 As Long = &H2
Private Const SERVICE_BOOT_START                        As Long = &H0
Private Const SERVICE_SYSTEM_START                      As Long = &H1
Private Const SERVICE_AUTO_START                        As Long = &H2
Private Const SERVICE_DEMAND_START                      As Long = &H3
Private Const SERVICE_DISABLED                          As Long = &H4
Private Const SERVICE_ERROR_NORMAL                      As Long = &H1
Private Const SERVICE_NO_CHANGE                         As Long = &HFFFFFFFF
Private Const SERVICE_QUERY_CONFIG                      As Long = &H1
Private Const SERVICE_CHANGE_CONFIG                     As Long = &H2
Private Const SERVICE_QUERY_STATUS                      As Long = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS              As Long = &H8
Private Const SERVICE_PAUSE_CONTINUE                    As Long = &H40
Private Const SERVICE_INTERROGATE                       As Long = &H80
Private Const SERVICE_USER_DEFINED_CONTROL              As Long = &H100
Private Const SERVICE_CSTART                            As Long = &H10
Private Const SERVICE_CSTOP                             As Long = &H20
Private Const SERVICE_CONTROL_STOP                      As Long = &H1
Private Const SERVICE_CONTROL_PAUSE                     As Long = &H2
Private Const SERVICE_CONTROL_CONTINUE                  As Long = &H3
Private Const SERVICE_CONTROL_INTERROGATE               As Long = &H4
Private Const SERVICE_CONTROL_SHUTDOWN                  As Long = &H5
Private Const SERVICE_CONTROL_PARAMCHANGE               As Long = &H6
Private Const SERVICE_CONTROL_NETBINDADD                As Long = &H7
Private Const SERVICE_CONTROL_NETBINDREMOVE             As Long = &H8
Private Const SERVICE_CONTROL_NETBINDENABLE             As Long = &H9
Private Const SERVICE_CONTROL_NETBINDDISABLE            As Long = &HA
Private Const STANDARD_RIGHTS_REQUIRED                  As Long = &HF0000

Private Const SERVICE_ALL_ACCESS                        As Double = _
        (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or _
        SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or _
        SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_CSTART Or _
        SERVICE_CSTOP Or SERVICE_PAUSE_CONTINUE Or _
        SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_WIN32                             As Long = _
        SERVICE_WIN32_OWN_PROCESS Or SERVICE_WIN32_SHARE_PROCESS
Private Const SERVICE_DRIVER                            As Long = _
        SERVICE_KERNEL_DRIVER Or SERVICE_FILE_SYSTEM_DRIVER Or _
        SERVICE_RECOGNIZER_DRIVER
Private Const SERVICE_TYPE_ALL                          As Long = _
        SERVICE_WIN32 Or SERVICE_ADAPTER Or SERVICE_DRIVER Or _
        SERVICE_INTERACTIVE_PROCESS

'/* service status
Private Type SERVICE_STATUS
    dwServiceType                                       As Long
    dwCurrentState                                      As Long
    dwControlsAccepted                                  As Long
    dwWin32ExitCode                                     As Long
    dwServiceSpecificExitCode                           As Long
    dwCheckPoint                                        As Long
    dwWaitHint                                          As Long
End Type

Private Type ENUM_SERVICE_STATUS
   lpServiceName As Long
   lpDisplayName As Long
   ServiceStatus As SERVICE_STATUS
End Type

Private Type SERVICE_DESCRIPTION
    lpDescription                               As String
End Type

'/* description
Private Enum InfoLevel
    SERVICE_CONFIG_DESCRIPTION = 1&
    SERVICE_CONFIG_FAILURE_ACTIONS = 2&
End Enum

'/* startup type
Public Enum eServiceStartType
    START_BOOT = &H0
    START_SYSTEM = &H1
    START_AUTO = &H2
    START_DEMAND = &H3
    START_DISABLED = &H4
End Enum

'/* service api
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal strMachineName As String, _
                                                                                  ByVal strDBName As String, _
                                                                                  ByVal lAccessReq As Long) As Long

Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, _
                                                                              ByVal strServiceName As String, _
                                                                              ByVal lAccessReq As Long) As Long

Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, _
                                                                                ByVal lNumServiceArgs As Long, _
                                                                                ByVal strArgs As String) As Boolean

Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, _
                                                            ByVal lControlCode As Long, _
                                                            lpServiceStatus As SERVICE_STATUS) As Boolean

Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hHandle As Long) As Boolean

Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, _
                                                                lpServiceStatus As SERVICE_STATUS) As Boolean

Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, _
                                                                                              ByVal dwServiceType As Long, _
                                                                                              ByVal dwStartType As eServiceStartType, _
                                                                                              ByVal dwErrorControl As Long, _
                                                                                              ByVal lpBinaryPathName As String, _
                                                                                              ByVal lpLoadOrderGroup As String, _
                                                                                              ByVal lpdwTagId As Long, _
                                                                                              ByVal lpDependencies As String, _
                                                                                              ByVal lpServiceStartName As String, _
                                                                                              ByVal lpPassword As String, _
                                                                                              ByVal lpDisplayName As String) As Boolean

Private Declare Function EnumServicesStatus Lib "advapi32" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, _
                                                                                        ByVal dwServiceType As Long, _
                                                                                        ByVal dwServiceState As Long, _
                                                                                        lpServices As Any, _
                                                                                        ByVal cbBufSize As Long, _
                                                                                        pcbBytesNeeded As Long, _
                                                                                        lpServicesReturned As Long, _
                                                                                        lpResumeHandle As Long) As Long

Private Declare Function CreateService Lib "advapi32" Alias "CreateServiceA" (ByVal hSCManager As Long, _
                                                                              ByVal lpServiceName As String, _
                                                                              ByVal lpDisplayName As String, _
                                                                              ByVal dwDesiredAccess As Long, _
                                                                              ByVal dwServiceType As Long, _
                                                                              ByVal dwStartType As Long, _
                                                                              ByVal dwErrorControl As Long, _
                                                                              ByVal lpBinaryPathName As String, _
                                                                              ByVal lpLoadOrderGroup As String, _
                                                                              ByVal lpdwTagId As String, _
                                                                              ByVal lpDependencies As String, _
                                                                              ByVal lp As String, _
                                                                              ByVal lpPassword As String) As Long

Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long

Private Declare Function ChangeServiceConfig2 Lib "advapi32.dll" Alias "ChangeServiceConfig2W" (ByVal hService As Long, _
                                                                                                ByVal dwInfoLevel As InfoLevel, _
                                                                                                lpInfo As Any) As Boolean

Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, _
                                                  ByVal Ptr As Long) As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long



'~*** Control Routines ***~
Public Function Service_Start() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long

On Error GoTo Handler

    '/* get handle to service manager
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    '/* get service handle
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)
    '/* start service
    lResult = StartService(lHService, 0&, 0&)
    '/* success
    If Not lResult = 0 Then
        Service_Start = True
    End If
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function

Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Stop() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim tStatus             As SERVICE_STATUS
Dim lResult             As Long

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)
    '/* stop service
    lResult = ControlService(lHService, SERVICE_CONTROL_STOP, tStatus)
    '/* success
    If Not lResult = 0 Then
        Service_Stop = True
    End If
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Pause() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim svcStatus           As SERVICE_STATUS
Dim lResult             As Long

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)
    '/* pause service
    lResult = ControlService(lHService, SERVICE_CONTROL_PAUSE, svcStatus)
    '/* success
    If Not lResult = 0 Then
        Service_Pause = True
    End If
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager

On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Continue() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim svcStatus           As SERVICE_STATUS
Dim lResult             As Long

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)
    '/* resume service
    lResult = ControlService(lHService, SERVICE_CONTROL_CONTINUE, svcStatus)
    '/* success
    If Not lResult = 0 Then
        Service_Continue = True
    End If
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager

On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_StartUp(ByRef svcStartType As eServiceStartType) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_CHANGE_CONFIG)
    '/* change startup type
    lResult = ChangeServiceConfig(lHService, SERVICE_NO_CHANGE, _
    svcStartType, SERVICE_NO_CHANGE, vbNullString, vbNullString, 0&, _
    vbNullString, vbNullString, vbNullString, vbNullString)
    '/* success
    If Not lResult = 0 Then
        Service_StartUp = True
    End If
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_State() As Long

Dim lHManager           As Long
Dim lHService           As Long
Dim sState              As String
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_QUERY_STATUS)
    '/* query service status
    QueryServiceStatus lHService, tStatus
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    '/* return service state
    '/* 0 - not exist, 1 - stopped, 2 - paused, 3 - waiting, 4 - running
    '/* 5 - shutting down 6 - param change 7 - binding
    Service_State = tStatus.dwCurrentState

On Error GoTo 0
Exit Function

Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Install() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tStatus             As SERVICE_STATUS
Dim sSvcPath            As String
Dim sAccount            As String

On Error GoTo Handler

    sSvcPath = App.Path + "\Service\rbsidx.exe"
    sAccount = "LocalSystem"
    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    lResult = CreateService(lHManager, SVC_NAME, "RBS Jet", SERVICE_ALL_ACCESS, _
        SERVICE_WIN32_OWN_PROCESS, SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL, _
        sSvcPath, vbNullString, vbNullString, vbNullString, sAccount, vbNullString)
    
    '/* success
    If Not lResult = 0 Then
        Service_Install = True
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Uninstall() As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tStatus             As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)

    '/* stop the service
    ControlService lHService, SERVICE_CONTROL_STOP, tStatus
    '/* delete service
    lResult = DeleteService(lHService)
    
    '/* success
    If Not lResult = 0 Then
        Service_Uninstall = True
    Else
        GoTo Handler
    End If
    
    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function

Public Function Service_Desc(ByVal sDescription As String) As Boolean

Dim lHManager           As Long
Dim lHService           As Long
Dim lResult             As Long
Dim tDesc               As SERVICE_DESCRIPTION

On Error GoTo Handler

    '/* get app and service handles
    lHManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lHService = OpenService(lHManager, SVC_NAME, SERVICE_ALL_ACCESS)
    sDescription = StrConv(sDescription, vbUnicode) + Chr$(0)
    tDesc.lpDescription = sDescription
    lResult = ChangeServiceConfig2(lHService, SERVICE_CONFIG_DESCRIPTION, tDesc)
    
    '/* success
    If Not lResult = 0 Then
        Service_Desc = True
    Else
        GoTo Handler
    End If

    '/* cleanup
    CloseServiceHandle lHService
    CloseServiceHandle lHManager
    
On Error GoTo 0
Exit Function
    
Handler:
    If Not lHService = 0 Then CloseServiceHandle lHService
    If Not lHManager = 0 Then CloseServiceHandle lHManager

End Function
