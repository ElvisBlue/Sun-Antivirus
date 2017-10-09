Attribute VB_Name = "modDetectHook"
Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Public Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long        'LPBYTE
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long


Public Const CREATE_SUSPENDED = &H4

Private OpenProcessCheck(5) As Byte
Private ZwOpenProcessCheck(5) As Byte
Private TerminateProcessCheck(5) As Byte

Public Sub IntCheckSumByte()

End Sub

Public Function IsBeingHookedOpenProcess(ByVal Restore As Byte) As Boolean

End Function

Public Function IsBeingHookedTerminateProcess(ByVal Restore As Byte) As Boolean

End Function


