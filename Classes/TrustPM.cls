VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrustPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://docs.microsoft.com/en-us/windows/win32/tbs/tbs-functions
Private Const TBS_SUCCESS                  As Long = 0 '0 (0x0) ' The function was successful.
Private Const TBS_E_BAD_PARAMETER          As Long = &H80284002 '2150121474 (0x80284002)'    One or more parameter values are not valid.
Private Const TBS_E_INTERNAL_ERROR         As Long = &H80284001 '2150121473 (0x80284001)'    An internal software error occurred.
Private Const TBS_E_INVALID_CONTEXT_PARAM  As Long = &H80284007 '2150121479 (0x80284007)'    A context parameter that is not valid was passed when attempting to create a TBS context.
Private Const TBS_E_INVALID_OUTPUT_POINTER As Long = &H80284003 '2150121475 (0x80284003)'    A specified output pointer is not valid.
Private Const TBS_E_SERVICE_DISABLED       As Long = &H80284010 '2150121488 (0x80284010)'    The TBS service has been disabled.
Private Const TBS_E_SERVICE_NOT_RUNNING    As Long = &H80284008 '2150121480 (0x80284008)'    The TBS service is not running and could not be started.
Private Const TBS_E_SERVICE_START_PENDING  As Long = &H8028400B '2150121483 (0x8028400B)'    The TBS service has been started but is not yet running.
Private Const TBS_E_TOO_MANY_TBS_CONTEXTS  As Long = &H80284009 '2150121481 (0x80284009)'    A new context could not be created because there are too many open contexts.
Private Const TBS_E_TPM_NOT_FOUND          As Long = &H8028400F '2150121487 (0x8028400F)'    A compatible Trusted Platform Module (TPM) Security Device cannot be found on this computer.
Private Const TBS_CONTEXT_VERSION_ONE      As Long = 1
'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-hgsa/9a647388-21bc-4451-9bc2-e70a35cf61a4
Private Const TPM_VERSION_UNKNOWN          As Long = 0
Private Const TPM_VERSION_12               As Long = 1
Private Const TPM_VERSION_20               As Long = 2

Private Const TPM_IFTYPE_UNKNOWN           As Long = &H0        'Unknown interface type.
Private Const TPM_IFTYPE_1                 As Long = &H1        'TPM 1.2 interface type that uses port-mapped or memory-mapped I/O.
Private Const TPM_IFTYPE_TRUSTZONE         As Long = &H2        'TPM 2.0 TrustZone interface.
Private Const TPM_IFTYPE_HW                As Long = &H3        'TPM 2.0 hardware interface.
Private Const TPM_IFTYPE_EMULATOR          As Long = &H4        'TPM 2.0 software emulator interface.

'Define TBS_RESULT    Alias Long
'Define PTBS_HCONTEXT Alias Long
'typedef struct tdTBS_CONTEXT_PARAMS {
'  UINT32 version;
'} TBS_CONTEXT_PARAMS, *PTBS_CONTEXT_PARAMS;

Private Type TBS_CONTEXT_PARAMS
    version As Long
End Type

'typedef struct tdTBS_CONTEXT_PARAMS2 {
'  UINT32 version;
'  union {
'    struct {
'      UINT32 requestRaw : 1;
'      UINT32 includeTpm12 : 1;
'      UINT32 includeTpm20 : 1;
'    };
'    UINT32 asUINT32;
'  };
'} TBS_CONTEXT_PARAMS2, *PTBS_CONTEXT_PARAMS2;
Private Type TBS_CONTEXT_PARAMS2
    version      As Long
    requestRaw   As Long ' = 1
    includeTpm12 As Long ' = 1
    includeTpm20 As Long ' = 1
End Type

'TBS_RESULT Tbsi_Context_Create(
'  PCTBS_CONTEXT_PARAMS pContextParams,
'  PTBS_HCONTEXT phContext
');
Private Declare Function Tbsi_Context_Create Lib "tbs" (ByVal pContextParams As Long, phContext_out As Long) As Long

'TBS_RESULT Tbsip_Context_Close(
'  TBS_HCONTEXT hContext
');
Private Declare Function Tbsip_Context_Close Lib "tbs" (ByVal hContext As Long) As Long


'typedef struct _TPM_DEVICE_INFO {
'  UINT32 structVersion;
'  UINT32 tpmVersion;
'  UINT32 tpmInterfaceType;
'  UINT32 tpmImpRevision;
'} TPM_DEVICE_INFO, *PTPM_DEVICE_INFO;

Private Type TPM_DEVICE_INFO
    structVersion    As Long
    tpmVersion       As Long
    tpmInterfaceType As Long
    tpmImpRevision   As Long
End Type

'TBS_RESULT Tbsi_GetDeviceInfo(
'  UINT32 Size,
'  PVOID Info
');
Private Declare Function Tbsi_GetDeviceInfo Lib "tbs" (ByVal Size As Long, ByRef pInfo_out As TPM_DEVICE_INFO) As Long

Private m_TBS_E         As Long

Private m_ContextParam1 As TBS_CONTEXT_PARAMS
Private m_ContextParam2 As TBS_CONTEXT_PARAMS2
Private m_ContextHandle As Long

Private m_TpmDeviceInfo As TPM_DEVICE_INFO

Public Function Error_ToStr() As String
    Dim s As String
    Select Case m_TBS_E
    Case TBS_E_BAD_PARAMETER:          s = "One or more parameter values are not valid."
    Case TBS_E_INTERNAL_ERROR:         s = "An internal software error occurred."
    Case TBS_E_INVALID_CONTEXT_PARAM:  s = "A context parameter that is not valid was passed when attempting to create a TBS context."
    Case TBS_E_INVALID_OUTPUT_POINTER: s = "A specified output pointer is not valid."
    Case TBS_E_SERVICE_DISABLED:       s = "The TBS service has been disabled."
    Case TBS_E_SERVICE_NOT_RUNNING:    s = "The TBS service is not running and could not be started."
    Case TBS_E_SERVICE_START_PENDING:  s = "The TBS service has been started but is not yet running."
    Case TBS_E_TOO_MANY_TBS_CONTEXTS:  s = "A new context could not be created because there are too many open contexts."
    Case TBS_E_TPM_NOT_FOUND:          s = "A compatible Trusted Platform Module (TPM) Security Device cannot be found on this computer."
    End Select
    Error_ToStr = s
End Function

Public Property Get ContextHandle() As Long
    ContextHandle = m_ContextHandle
End Property

Private Sub Class_Terminate()
    If m_ContextHandle Then Tbsip_Context_Close m_ContextHandle
End Sub

Public Function CheckContext1() As Boolean
    m_ContextParam1.version = TBS_CONTEXT_VERSION_ONE
    m_TBS_E = Tbsi_Context_Create(VarPtr(m_ContextParam1), m_ContextHandle)
    CheckContext1 = (m_TBS_E = TBS_SUCCESS)
End Function

Public Function CheckContext2() As Boolean
    With m_ContextParam2
        .version = TPM_VERSION_20 'TBS_CONTEXT_VERSION_ONE
        .requestRaw = 1
        .includeTpm12 = 1
        .includeTpm20 = 1
    End With
    m_TBS_E = Tbsi_Context_Create(VarPtr(m_ContextParam2), m_ContextHandle)
    CheckContext2 = (m_TBS_E = TBS_SUCCESS)
End Function

Public Function CheckDeviceInfo() As Boolean
    If m_TpmDeviceInfo.structVersion = 0 Then
        m_TpmDeviceInfo.structVersion = TPM_VERSION_20
    End If
    m_TBS_E = Tbsi_GetDeviceInfo(LenB(m_TpmDeviceInfo), m_TpmDeviceInfo)
    CheckDeviceInfo = (m_TBS_E = TBS_SUCCESS)
End Function

Public Property Get DeviceInfo_StructVer() As Long
    DeviceInfo_StructVer = m_TpmDeviceInfo.structVersion
End Property
Public Property Let DeviceInfo_StructVer(ByVal Value As Long)
    m_TpmDeviceInfo.structVersion = Value
End Property

Public Property Get DeviceInfo_TPMVersion() As Long
    DeviceInfo_TPMVersion = m_TpmDeviceInfo.tpmVersion
End Property
Public Property Get DeviceInfo_TPMInterfaceType() As Long
    DeviceInfo_TPMInterfaceType = m_TpmDeviceInfo.tpmInterfaceType
End Property
Public Property Get DeviceInfo_TPMImpRevision() As Long
    DeviceInfo_TPMImpRevision = m_TpmDeviceInfo.tpmImpRevision
End Property

Public Function DeviceInfo_ToStr() As String
    Dim s As String
    With m_TpmDeviceInfo
        s = s & "TpmDeviceInfo{" & vbNewLine
        s = s & "    StructVer    : " & .structVersion & vbNewLine
        s = s & "    TpmVersion   : " & .tpmVersion & vbNewLine
        s = s & "    InterfaceType: " & .tpmInterfaceType & vbNewLine
        s = s & "    ImpRevision  : " & .tpmImpRevision & vbNewLine
        s = s & "}"
    End With
    DeviceInfo_ToStr = s
End Function
