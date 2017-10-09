Attribute VB_Name = "modDebugToken"
Option Explicit
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const SE_DEBUG_NAME As String = "SeDebugPrivilege"

Type LARGE_INTEGER
LowPart As Long
HighPart As Long
End Type
Type LUID
LowPart As Long
HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type

Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type


Public Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function OpenProcessToken Lib "advapi32.dll" _
(ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long

Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias _
"LookupPrivilegeValueA" (ByVal lpSystemName As String, _
ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
(ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Public Declare Function AdjustTokenPrivileges1 Lib "advapi32.dll" _
Alias "AdjustTokenPrivileges" _
(ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
PreviousState As Long, ReturnLength As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessID As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" _
(ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Function AdvanceDEBUGToken(Optional ByVal value As Boolean = True) As Boolean
Dim hProcess As Long
Dim hToken As Long ' Handle to your process token.
Dim lPrivilege As Long ' Privilege to enable/disable
Dim iPrivilegeflag As Boolean ' Flag whether to enable/disable ' the privilege of concern.
Dim lResult As Long ' Result call of various APIs.
' get our current process handle
hProcess = GetCurrentProcess
lResult = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY, hToken)
If (lResult = 0) Then
Exit Function
CloseHandle (hToken)
Exit Function
End If
lResult = SetPrivilege(hToken, SE_DEBUG_NAME, value)
If (lResult = 0) Then
CloseHandle (hToken)
Exit Function
End If
CloseHandle (hProcess)
CloseHandle (hToken)
AdvanceDEBUGToken = True
End Function

' The SetPrivilege function will accept a handle to a token, a
' privilege, and a flag to either enable/disable that privilege. The
' function will attempt to perform the desired action upon the token
' returning TRUE if it succeeded, or FALSE if it failed.
Private Function SetPrivilege(hToken As Long, Privilege As String, _
bSetFlag As Boolean) As Boolean

Dim Tp As TOKEN_PRIVILEGES ' Used in getting the current token privileges
Dim TPPrevious As TOKEN_PRIVILEGES ' Used in setting the new token privileges
Dim LUID As LUID ' Stores the Local Unique ' Identifier - refer to MSDN
Dim cbPrevious As Long ' Previous size of the TOKEN_PRIVILEGES structure
Dim lResult As Long ' Result of various API calls
' Grab the size of the TOKEN_PRIVILEGES structure, used in making the API calls.
cbPrevious = Len(Tp)
' Grab the LUID for the request privilege.
lResult = LookupPrivilegeValue("", Privilege, LUID)
' If LoopupPrivilegeValue fails, the return result will be zero Test to make sure that the call succeeded.
If (lResult = 0) Then
SetPrivilege = False
Exit Function
End If
Tp.PrivilegeCount = 1
Tp.Privileges(0).pLuid = LUID
Tp.Privileges(0).Attributes = 0
' You need to acquire the current privileges first
lResult = AdjustTokenPrivileges(hToken, False, Tp, Len(Tp), TPPrevious, cbPrevious)
' If AdjustTokenPrivileges fails, the return result is zero, ' test for success.
If (lResult = 0) Then
SetPrivilege = False
Exit Function
End If
' Now you can set the token privilege information ' to what the user is requesting.
TPPrevious.PrivilegeCount = 1
TPPrevious.Privileges(0).pLuid = LUID

' either enable or disable the privilege,
' depending on what the user wants.
Select Case bSetFlag
Case True: TPPrevious.Privileges(0).Attributes = _
TPPrevious.Privileges(0).Attributes Or _
(SE_PRIVILEGE_ENABLED)
Case False: TPPrevious.Privileges(0).Attributes = _
TPPrevious.Privileges(0).Attributes Xor _
(SE_PRIVILEGE_ENABLED And _
TPPrevious.Privileges(0).Attributes)
End Select

' Call adjust the token privilege information.
lResult = AdjustTokenPrivileges1(hToken, False, TPPrevious, _
cbPrevious, ByVal &O0, ByVal 0&)

' Determine your final result of this function.
If (lResult = 0) Then
' You were not able to set the privilege on this token.
SetPrivilege = False
Else
' You managed to modify the token privilege
SetPrivilege = True
End If

End Function

