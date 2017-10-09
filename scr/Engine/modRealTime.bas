Attribute VB_Name = "modRealTime"
Public Type RealTime_List
    PID As Long
    Detected As Infected
End Type

'########### PUBLIC VARIABLE #############
Public RealTimeList(100) As RealTime_List
Public RealTimeCnt As Byte
Public InjectedProcess As PROCESSENTRY32
'########### PUBLIC VARIABLE #############

Dim NullProcess As PROCESSENTRY32


Public Sub DetectInjectedProcess()

'Detect svchost.exe injection
'Some malware/crypter try to inject its code to system process
'almost is RunPE Injection
'and almost is svchost.exe

Dim i As Byte
Dim ServicesPID As Long

Call modProcess.RefreshProcess
i = 1
While (ServicesPID = 0) And i < modProcess.CntProcess
    If UCase(StripNULL(modProcess.ListProcess(i).szExeFile)) = "SERVICES.EXE" Then ServicesPID = modProcess.ListProcess(i).th32ProcessID
    i = i + 1
Wend

If ServicesPID = 0 Then GoTo ExitF

For i = 1 To (CntProcess - 1)
    If UCase(StripNULL(modProcess.ListProcess(i).szExeFile)) = "SVCHOST.EXE" Then
        If modProcess.ListProcess(i).th32ParentProcessID <> ServicesPID Then
            InjectedProcess = modProcess.ListProcess(i)
            Exit Sub
        End If
    End If
Next

ExitF:
InjectedProcess = NullProcess
End Sub

Public Function DetectMalwareProcess()
Dim i As Byte
Dim FilePath As String
Dim Result As ScanResult
RealTimeCnt = 0
Call modProcess.RefreshProcess
For i = 0 To (CntProcess - 1)
    FilePath = modProcess.ProcessPathByPID(ListProcess(i).th32ProcessID)
    Result = modScanning.ScanFile(FilePath)
    If Result.Clean = False Then
        'Suspend All Malware Process
        'Call SuspendResumeProcess(ListProcess(i).th32ProcessID, True)
        RealTimeList(RealTimeCnt).Detected.FilePath = FilePath
        RealTimeList(RealTimeCnt).Detected.VirusName = Result.VirusName
        RealTimeList(RealTimeCnt).PID = modProcess.ListProcess(i).th32ProcessID
        RealTimeCnt = RealTimeCnt + 1
    End If
Next
End Function

Public Function DeleteLastRTItem(ByVal DelMalware As Boolean) As Boolean
On Error GoTo ExitF

If RealTimeCnt = 0 Then GoTo ExitF
If DelMalware = True Then
    If modRemoveFile.KillFile(RealTimeList(RealTimeCnt - 1).Detected.FilePath) = False Then GoTo ExitF
End If
RealTimeCnt = RealTimeCnt - 1
Exit Function


ExitF:
DeleteItem = False
End Function

Public Function SetQuarantineLastRTItem() As Boolean
    Call AddQuarantine(RealTimeList(RealTimeCnt - 1).Detected.FilePath, RealTimeList(RealTimeCnt - 1).Detected.VirusName)
    DeleteLastRTItem (False)
End Function
