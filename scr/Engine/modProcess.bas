Attribute VB_Name = "modProcess"
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" _
      Alias "CreateToolhelp32Snapshot" ( _
      ByVal lFlags As Long, _
      ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" _
      Alias "Process32First" ( _
      ByVal hsnapshot As Long, _
      uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" _
      Alias "Process32Next" ( _
      ByVal hsnapshot As Long, _
      uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" ( _
      ByVal hProcess As Long, _
      ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" ( _
      ByVal dwDesiredAccess As Long, _
      ByVal bInheritHandle As Long, _
      ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Public Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + _
  TH32CS_SNAPmodule
Public Const PROCESS_TERMINATE = 1
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_VM_READ = &H10
Public Const MAX_PATH As Integer = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF

'define PROCESSENTRY32 structure

Public Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type
'########### PUBLIC VARIABLE #############
Public ListProcess() As PROCESSENTRY32
Public CntProcess
'########### PUBLIC VARIABLE #############

Public Sub RefreshProcess()

Dim TheLoop As Long
Dim Proc As PROCESSENTRY32
Dim Snap As Long
 
ReDim ListProcess(0) As PROCESSENTRY32
CntProcess = 1

Snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
Proc.dwSize = Len(Proc)
TheLoop = ProcessFirst(Snap, Proc)
ListProcess(0) = Proc
TheLoop = ProcessNext(Snap, Proc)
    
    While TheLoop <> 0
        ReDim Preserve ListProcess(CntProcess) As PROCESSENTRY32
        CntProcess = CntProcess + 1
        ListProcess(CntProcess - 1) = Proc
        TheLoop = ProcessNext(Snap, Proc)
    Wend
    
    CloseHandle Snap
End Sub

Public Function KillProcessByPID(ByVal PID As Long) As Boolean
Dim hProcess As Long
Dim uExitCode As Long

hProcess = OpenProcess(PROCESS_TERMINATE + PROCESS_QUERY_INFORMATION, False, PID)
If hProcess = 0 Then GoTo TerminateFailed
Call GetExitCodeProcess(hProcess, uExitCode)
If TerminateProcess(hProcess, uExitCode) = False Then GoTo TerminateFailed
KillProcessByPID = True

Exit Function
TerminateFailed:
KillProcessByPID = False
End Function

Function ProcessPathByPID(PID As Long) As String
'Return path to the executable from PID
'http://support.microsoft.com/default.aspx?scid=kb;en-us;187913
Dim cbNeeded As Long
Dim Modules(1 To 200) As Long
Dim ret As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long

hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
            
If hProcess <> 0 Then
    ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
    If ret <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = 500
        ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
        ProcessPathByPID = Left(ModuleName, ret)
    End If
End If

ret = CloseHandle(hProcess)
End Function

Public Function GetFileNameFromPath(ByVal FilePath As String)
Dim pos As Byte
pos = InStr(StrReverse(FilePath), "\")
If pos = 0 Then Exit Function
GetFileNameFromPath = Right(FilePath, pos - 1)
End Function

Public Function KillProcessByPath(ByVal FilePath As String) As Boolean
Dim i As Byte

Call RefreshProcess
For i = 1 To CntProcess
    If ProcessPathByPID(ListProcess(i - 1).th32ProcessID) = FilePath Then
        KillProcessByPID (ListProcess(i - 1).th32ProcessID)
        KillProcessByPID2 (ListProcess(i - 1).th32ProcessID)
    End If
Next
End Function

Public Function SuspendResumeProcess(ByVal Procid As Long, ByVal SuspendResume As Boolean) As Boolean
Dim hsnapshot As Long
Dim htthread As Long
Dim pthread As Boolean
Dim pt As THREADENTRY32

SuspendResumeProcess = False
hsnapshot = CreateToolhelpSnapshot(TH32CS_SNAPthread, 0)
pt.dwSize = Len(pt)
pthread = Thread32First(hsnapshot, pt)

While pthread
    If pt.th32OwnerProcessID = Procid Then
        htthread = OpenThread(THREAD_ALL_ACCESS, 0, pt.th32ThreadID)
        If htthread <> 0 Then
            If SuspendResume Then SuspendThread (htthread) Else ResumeThread (htthread)
            CloseHandle htthread
            SuspendResumeProcess = True
        End If
    End If
    pthread = Thread32Next(hsnapshot, pt)
Wend

CloseHandle hsnapshot
End Function
