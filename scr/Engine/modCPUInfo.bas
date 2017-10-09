Attribute VB_Name = "modCPUInfo"
Option Explicit


'=============================================================================================================
'
' modCPUInfo Module
' -----------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : July 14, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : NOTHING
'
' Description : This module is meant to give you easy access to the vitals of the computer and operating
'               system (OS) that the program is running on.  This module gives you easy access to the following
'               things:
'                 - The Type of OS          (Win95, WinNT, etc)
'                 - The OS Version          (3.51, 4.0, etc)
'                 - The OS Build            (1381, etc)
'                 - The OS ServicePack      (Win95B, WinNT SP6, etc)
'                 - The # of CPUs installed (1, 2, etc)
'                 - The Type of CPU         (386, 486, Pentium, etc)
'                 - The CPU Speed           (350 MHz, etc)  -  Not always available
'                 - The Installed RAM       (16MB, etc)
'
' Example Use :
'
'  If GetSysInfo = False Then
'    Exit Sub
'  End If
'
'  MsgBox "Your computer information :" & Chr(13) & Chr(13) & _
'         "The Type of OS = Windows " & OS_Type & Chr(13) & _
'         "The OS Version = " & OS_Version & Chr(13) & _
'         "The OS Build = " & OS_Build & Chr(13) & _
'         "The OS ServicePack = " & OS_ServicePack & Chr(13) & _
'         "The # of CPUs installed = " & CPU_Count & Chr(13) & _
'         "The Type of CPU = " & CPU_Type & Chr(13) & _
'         "The CPU Speed = " & CPU_Speed & Chr(13) & _
'         "The Installed RAM = " & CPU_RAM
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Type for getting OS information
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long         ' Specifies the size, in bytes, of this data structure. Set this member to sizeof(OSVERSIONINFO) before calling the GetVersionEx function.
  dwMajorVersion      As Long         ' Identifies the major version number of the operating system. For example, for Windows NT version 3.51, the major version number is 3; and for Windows NT version 4.0, the major version number is 4.
  dwMinorVersion      As Long         ' Identifies the minor version number of the operating system. For example, for Windows NT version 3.51, the minor version number is 51; and for Windows NT version 4.0, the minor version number is 0.
  dwBuildNumber       As Long         ' Windows 95 : Identifies the build number of the operating system in the low-order word. The high-order word contains the major and minor version numbers.
                                      ' Windows NT : Identifies the build number of the operating system.
  dwPlatformId        As Long         ' Identifies the operating system platform. (See Constants)
  szCSDVersion        As String * 128 ' Windows 95 : Contains a null-terminated string that provides arbitrary additional information about the operating system.
                                      ' Windows NT : Contains a null-terminated string, such as "Service Pack 3", that indicates the latest Service Pack installed on the system. If no Service Pack has been installed, the string is empty.
End Type

' Type for getting CPU information
Private Type SYSTEM_INFO
  dwOemID                     As Long ' An obsolete member that is retained for compatibility with previous versions of Windows NT. Beginning with Windows NT 3.51 and the initial release of Windows 95, use the wProcessorArchitecture branch of the union.
                                      ' Windows 95: The system always sets this member to zero, the value defined for PROCESSOR_ARCHITECTURE_INTEL.
  dwPageSize                  As Long ' Specifies the system’s processor architecture (See Constants)
  lpMinimumApplicationAddress As Long ' Pointer to the lowest memory address accessible to applications and dynamic-link libraries (DLLs).
  lpMaximumApplicationAddress As Long ' Pointer to the highest memory address accessible to applications and DLLs.
  dwActiveProcessorMask       As Long ' Specifies a mask representing the set of processors configured into the system. Bit 0 is processor 0; bit 31 is processor 31.
  dwNumberOrfProcessors       As Long ' Specifies the number of processors in the system.
  dwProcessorType             As Long ' Windows 95 : Specifies the type of processor in the system (See Constants)
                                      ' Windows NT : This member is no longer relevant, but is retained for compatibility with Windows 95 and previous versions of Windows NT. Use the wProcessorArchitecture, wProcessorLevel, and wProcessorRevision members to determine the type of processor.
  dwAllocationGranularity     As Long ' Specifies the granularity with which virtual memory is allocated. For example, a VirtualAlloc request to allocate 1 byte will reserve an address space of dwAllocationGranularity bytes. This value was hard coded as 64K in the past, but other hardware architectures may require different values.
  dwReserved                  As Long ' Reserved for future use.
End Type

' Type for getting RAM / Memory information
Private Type MEMORYSTATUS
  dwLength                    As Long ' Indicates the size of the structure. The calling process should set this member prior to calling GlobalMemoryStatus.
  dwMemoryLoad                As Long ' Specifies a number between 0 and 100 that gives a general idea of current memory utilization, in which 0 indicates no memory use and 100 indicates full memory use.
  dwTotalPhys                 As Long ' Indicates the total number of bytes of physical memory.
  dwAvailPhys                 As Long ' Indicates the number of bytes of physical memory available.
  dwTotalPageFile             As Long ' Indicates the total number of bytes that can be stored in the paging file. Note that this number does not represent the actual physical size of the paging file on disk.
  dwAvailPageFile             As Long ' Indicates the number of bytes available in the paging file.
  dwTotalVirtual              As Long ' Indicates the total number of bytes that can be described in the user mode portion of the virtual address space of the calling process.
  dwAvailVirtual              As Long ' Indicates the number of bytes of unreserved and uncommitted memory in the user mode portion of the virtual address space of the calling process.
End Type

' OS Type Enumerations
Public Enum OSTypes
  OS_Unknown = 0     ' "Unknown"
  OS_Win32 = 32      ' "Win 32"
  OS_Win95 = 95      ' "Windows 95"
  OS_Win98 = 98      ' "Windows 98"
  OS_WinNT_351 = 351 ' "Windows NT 3.51"
  OS_WinNT_40 = 40   ' "Windows NT 4.0"
  OS_Win2000 = 2000  ' "Windows 2000"
End Enum

' Registry Key Enumerations
Private Enum RegistryKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_CURRENT_USER = &H80000001
  HKEY_DYN_DATA = &H80000006
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_USERS = &H80000003
End Enum

' General Constants
Private Const ERROR_SUCCESS = 0&

' SYSTEM_INFO.dwProcessorType Constants
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000        ' Windows NT / R4101 & R3910 for Windows CE
Private Const PROCESSOR_ALPHA_21064 = 21064      ' Windows NT
Private Const PROCESSOR_PPC_601 = 601            ' Windows NT
Private Const PROCESSOR_PPC_603 = 603            ' Windows NT
Private Const PROCESSOR_PPC_604 = 604            ' Windows NT
Private Const PROCESSOR_PPC_620 = 620            ' Windows NT
Private Const PROCESSOR_HITACHI_SH3 = 10003      ' Windows CE
Private Const PROCESSOR_HITACHI_SH3E = 10004     ' Windows CE
Private Const PROCESSOR_HITACHI_SH4 = 10005      ' Windows CE
Private Const PROCESSOR_MOTOROLA_821 = 821       ' Windows CE
Private Const PROCESSOR_SHx_SH3 = 103            ' Windows CE
Private Const PROCESSOR_SHx_SH4 = 104            ' Windows CE
Private Const PROCESSOR_STRONGARM = 2577         ' Windows CE - 0xA11
Private Const PROCESSOR_ARM720 = 1824            ' Windows CE - 0x720
Private Const PROCESSOR_ARM820 = 2080            ' Windows CE - 0x820
Private Const PROCESSOR_ARM920 = 2336            ' Windows CE - 0x920
Private Const PROCESSOR_ARM_7TDMI = 70001        ' Windows CE

' OSVERSIONINFO.dwPlatformId Constants
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' RegQueryValueEx Constants
Private Const REG_BINARY = 3              ' Binary data in any form.
Private Const REG_DWORD = 4               ' A 32-bit number.
Private Const REG_DWORD_LITTLE_ENDIAN = 4 ' A 32-bit number in little-endian format (same as REG_DWORD). In little-endian format, the most significant byte of a word is the high-order byte. This is the most common format for computers running Windows NT and Windows 95.
Private Const REG_DWORD_BIG_ENDIAN = 5    ' A 32-bit number in big-endian format. In big-endian format, the most significant byte of a word is the low-order byte.
Private Const REG_EXPAND_SZ = 2           ' A null-terminated string that contains unexpanded references to environment variables (for example, “%PATH%”). It will be a Unicode or ANSI string depending on whether you use the Unicode or ANSI functions.
Private Const REG_LINK = 6                ' A Unicode symbolic link.
Private Const REG_MULTI_SZ = 7            ' An array of null-terminated strings, terminated by two null characters.
Private Const REG_NONE = 0                ' No defined value type.
Private Const REG_RESOURCE_LIST = 8       ' A device-driver resource list.
Private Const REG_SZ = 1                  ' A null-terminated string. It will be a Unicode or ANSI string depending on whether you use the Unicode or ANSI functions.

' Declare Variables
Public OS_Type          As OSTypes
Public OS_Version       As String
Public OS_Build         As String
Public OS_ServicePack   As String
Public CPU_Count        As Long
Public CPU_Type         As String
Public CPU_Speed        As String
Public CPU_RAM          As String

' Windows API Declarations
Private Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long

' Function that gets all the information about the OS and the CPU and stores it in variables
Public Function GetSysInfo() As Boolean
On Error Resume Next
  
  Dim SysInfo As SYSTEM_INFO
  Dim Memory As MEMORYSTATUS
  Dim MyCounter As Long
  Dim TempSpeed As Long
  Dim LowOrder As Integer
  Dim HighOrder As Integer
  Dim Number_Order As Long
  
  ' Get the OS information to determine how to get the CPU information
  If GetOSInfo = False Then
    Exit Function
  End If
  
  ' Get the CPU information
  GetSystemInfo SysInfo
  
  With SysInfo
    
    ' Get how many CPUs the system has
    CPU_Count = .dwNumberOrfProcessors
    
    ' Get what kind of CPU it is
    Select Case .dwProcessorType
      Case PROCESSOR_INTEL_386
        CPU_Type = "Intel 80386"
      Case PROCESSOR_INTEL_486
        CPU_Type = "Intel 80486"
      Case PROCESSOR_INTEL_PENTIUM
        CPU_Type = "Intel Pentium"
      Case PROCESSOR_MIPS_R4000
        CPU_Type = "MIPS R4000"
      Case PROCESSOR_ALPHA_21064
        CPU_Type = "ALPHA 21064"
      Case PROCESSOR_PPC_601
        CPU_Type = "PPC 601"
      Case PROCESSOR_PPC_603
        CPU_Type = "PPC 603"
      Case PROCESSOR_PPC_604
        CPU_Type = "PPC 604"
      Case PROCESSOR_PPC_620
        CPU_Type = "PPC 620"
      Case PROCESSOR_HITACHI_SH3
        CPU_Type = "HITACHI SH3"
      Case PROCESSOR_HITACHI_SH3E
        CPU_Type = "HITACHI SH3E"
      Case PROCESSOR_HITACHI_SH4
        CPU_Type = "HITACHI SH4"
      Case PROCESSOR_MOTOROLA_821
        CPU_Type = "MOTOROLA 821"
      Case PROCESSOR_SHx_SH3
        CPU_Type = "SHx SH3"
      Case PROCESSOR_SHx_SH4
        CPU_Type = "SHx SH4"
      Case PROCESSOR_STRONGARM
        CPU_Type = "STRONGARM"
      Case PROCESSOR_ARM720
        CPU_Type = "ARM 720"
      Case PROCESSOR_ARM820
        CPU_Type = "ARM 820"
      Case PROCESSOR_ARM920
        CPU_Type = "ARM 920"
      Case PROCESSOR_ARM_7TDMI
        CPU_Type = "ARM 7TDMI"
      Case Else
        CPU_Type = "Unknown - " & CStr(.dwProcessorType)
    End Select
  End With
  
  ' Find Processor Speed (if it's available)
  TempSpeed = GetRegDWORD(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "~MHz")
  CPU_Speed = CStr(TempSpeed) & " MHz"
  
  ' Get how much total ram is on the user's computer
  Memory.dwLength = Len(Memory)
  GlobalMemoryStatus Memory
  CPU_RAM = FormatFileSize(Memory.dwTotalPhys)
  
  GetSysInfo = True
  
End Function

' Function to gets the windows information only
Public Function GetOSInfo() As Boolean
On Error GoTo ErrorTrap
  
  Dim OSinfo As OSVERSIONINFO
  Dim PID As String
  
  ' Setup the variable to be passed
  With OSinfo
    .dwOSVersionInfoSize = Len(OSinfo) '148
    .szCSDVersion = Space(128)
  End With
  
  ' Get the OS info and store it in the OSInfo variable
  If GetVersionEx(OSinfo) = 0 Then
    MsgBox "An error occured while trying to get the OS version and information." & Chr(13) & "Click OK to continue.", vbOKOnly + vbExclamation, "  Error  -  GetVersionEx"
    GetOSInfo = False
    Exit Function
  End If
  
  ' Take the information retrieved and store it in variables that can be easily used
  With OSinfo
    Select Case .dwPlatformId
      Case VER_PLATFORM_WIN32s
        PID = "Win 32"
        OS_Type = OS_Win32
      Case VER_PLATFORM_WIN32_WINDOWS
        If .dwMinorVersion = 0 Then
          PID = "Windows 95"
          OS_Type = OS_Win95
        ElseIf .dwMinorVersion = 10 Then
          PID = "Windows 98"
          OS_Type = OS_Win98
        End If
      Case VER_PLATFORM_WIN32_NT
        If .dwMajorVersion = 3 Then
          PID = "Windows NT 3.51"
          OS_Type = OS_WinNT_351
        ElseIf .dwMajorVersion = 4 Then
          PID = "Windows NT 4.0"
          OS_Type = OS_WinNT_40
        ElseIf .dwMajorVersion = 5 Then
          PID = "Windows 2000"
          OS_Type = OS_Win2000
        End If
      Case Else
        PID = "Unknown"
        OS_Type = OS_Unknown
    End Select
    
    OS_Version = CStr(.dwMajorVersion) & "." & CStr(.dwMinorVersion)
    OS_Build = CStr(.dwBuildNumber)
    OS_ServicePack = Trim(.szCSDVersion)
    
    If InStr(OS_ServicePack, Chr(0)) = 0 Then
      If OS_ServicePack = "" Then
        OS_ServicePack = ""
      End If
    Else
      OS_ServicePack = Left(OS_ServicePack, InStr(OS_ServicePack, Chr(0)) - 1)
    End If
    
  End With
  
  GetOSInfo = True
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then
    Resume Next
  ElseIf Err.Number = 20 Then
    Resume Next
  Else
    MsgBox Err.Source & " caused the following error while getting the OS version:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    GetOSInfo = False
    Exit Function
  End If
  
End Function

' Function that retrieves a STRING from the Windows Registry
Private Function GetRegString(ByVal TheKey As RegistryKeys, ByVal TheSubKey As String, ByVal TheValue As String) As String
On Error Resume Next
  
  Dim KeyHandle As Long
  Dim TempString As String
  Dim BufferSize As Long
  Dim KeyType As Long
  
  ' Open the key
  RegOpenKey TheKey, TheSubKey, KeyHandle
  
  ' Find the type of key and the size of the registry entry
  RegQueryValueEx KeyHandle, TheValue, 0&, KeyType, ByVal 0&, BufferSize
  If KeyType = REG_SZ Then
    TempString = String(BufferSize, " ")
    
    ' Get the value of the registry entry
    If RegQueryValueEx(KeyHandle, TheValue, 0&, 0&, ByVal TempString, BufferSize) = ERROR_SUCCESS Then
      If InStr(TempString, Chr(0)) <> 0 Then
        GetRegString = Left(TempString, InStr(TempString, Chr(0)) - 1)
      Else
        GetRegString = TempString
      End If
    End If
  End If
  
  ' Close the key
  RegCloseKey KeyHandle
  
End Function

' Function that retrieves a LONG from the Windows Registry
Private Function GetRegDWORD(ByVal TheKey As RegistryKeys, ByVal TheSubKey As String, ByVal TheValue As String) As Long
On Error Resume Next

  Dim KeyType As Long
  Dim TempLong As Long
  Dim KeyHandle As Long
  
  ' Open the key
  RegOpenKey TheKey, TheSubKey, KeyHandle
  
  ' Get the registry entry and type
  If RegQueryValueEx(KeyHandle, TheValue, 0&, KeyType, TempLong, 4) = ERROR_SUCCESS Then
    If KeyType = REG_DWORD Then
      GetRegDWORD = TempLong
    End If
  End If
  
  ' Close the key
  RegCloseKey KeyHandle
  
End Function

' Function that formats the file size in bytes into KB/MB/GB
Private Function FormatFileSize(ByVal TheSize_BYTEs As Long) As String
On Error Resume Next
  
  Const KB As Long = 1024
  Const MB As Long = KB * KB
  Dim FormatSoFar As String
  
  ' Return size of file in kilobytes.
  If TheSize_BYTEs = -1 Then
    FormatFileSize = "0.0KB (0 bytes)"
  ElseIf TheSize_BYTEs < KB Then
    FormatSoFar = Format(TheSize_BYTEs, "#,##0") & " bytes"
  Else
    Select Case TheSize_BYTEs \ KB
      Case Is < 10
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.00") & "KB"
      Case Is < 100
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.0") & "KB"
      Case Is < 1000
        FormatSoFar = Format(TheSize_BYTEs / KB, "0.0") & "KB"
      Case Is < 10000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.00") & "MB"
      Case Is < 100000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.0") & "MB"
      Case Is < 1000000
        FormatSoFar = Format(TheSize_BYTEs / MB, "0.0") & "MB"
      Case Is < 10000000
        FormatSoFar = Format(TheSize_BYTEs / MB / KB, "0.00") & "GB"
    End Select
    FormatSoFar = FormatSoFar & " (" & Format(TheSize_BYTEs, "#,##0") & " bytes)"
  End If
  
  FormatFileSize = FormatSoFar
  
End Function

