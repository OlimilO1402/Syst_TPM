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
Private Const TBS_SUCCESS                  As Long = 0 '0 (0x0) '    The function was successful.
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
'Define TBS_RESULT    Alias Long
'Define PTBS_HCONTEXT Alias Long
'typedef struct tdTBS_CONTEXT_PARAMS {
'  UINT32 version;
'} TBS_CONTEXT_PARAMS, *PTBS_CONTEXT_PARAMS;

Private Type TBS_CONTEXT_PARAMS
    version As Long
End Type

'TBS_RESULT Tbsi_Context_Create(
'  PCTBS_CONTEXT_PARAMS pContextParams,
'  PTBS_HCONTEXT phContext
');
Private Declare Function Tbsi_Context_Create Lib "tbs" (pContextParams As TBS_CONTEXT_PARAMS, phContext_out As Long) As Long

Private m_ver    As TBS_CONTEXT_PARAMS
Private m_Handle As Long
Private m_TBS_E  As Long

Private Sub Class_Initialize()
    m_ver.version = TBS_CONTEXT_VERSION_ONE
    m_TBS_E = Tbsi_Context_Create(m_ver, m_Handle)
    If m_TBS_E = TBS_SUCCESS Then Exit Sub
End Sub

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

Public Property Get Handle() As Long
    Handle = m_Handle
End Property
