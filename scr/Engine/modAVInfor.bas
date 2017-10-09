Attribute VB_Name = "modAVInfor"
Public Enum MalwareHandle
    RemoveMalware = 0
    MoveToQuarantine = 1
    Wait = 2
End Enum

Public Type AV_Infor
    IsScanning As Boolean
    RealTime As Boolean
    RTAutoKill As Boolean
    ScanAutoKill As MalwareHandle
    ScanPE As Boolean
    StartUp As Boolean
    ProtectMyself As Boolean
    DetectSusProc As Boolean
    AutoScanUSB As Boolean
End Type

'########### PUBLIC VARIABLE #############
Public AvVersion As String
Public AVSettings As AV_Infor
'########### PUBLIC VARIABLE #############

Public Sub LoadSetting()
On Error GoTo DefaultSetting
Dim Data As String
Dim Temp1() As String

ReDim ScanList(0) As Infected
If Dir(App.Path & "\Settings.Sun") = vbNullString Then GoTo DefaultSetting
Data = ReadAllText(App.Path & "\Settings.Sun")
Temp1 = Split(Data, " ")

With AVSettings
    .AutoScanUSB = NumToBool(Temp1(0))
    .DetectSusProc = NumToBool(Temp1(1))
    .ProtectMyself = NumToBool(Temp1(2))
    .RealTime = NumToBool(Temp1(3))
    .RTAutoKill = NumToBool(Temp1(4))
    .ScanAutoKill = Temp1(5)
    .ScanPE = NumToBool(Temp1(6))
    .StartUp = NumToBool(Temp1(7))
End With
Exit Sub



DefaultSetting:
With AVSettings
    .AutoScanUSB = True
    .DetectSusProc = True
    .ProtectMyself = False
    .RealTime = True
    .RTAutoKill = False
    .ScanAutoKill = Wait
    .ScanPE = True
    .StartUp = False
End With
End Sub

Public Sub SaveSetting()
Dim Data As String
With AVSettings
    Data = BoolToNum(.AutoScanUSB) & " " & _
            BoolToNum(.DetectSusProc) & " " & _
            BoolToNum(.ProtectMyself) & " " & _
            BoolToNum(.RealTime) & " " & _
            BoolToNum(.RTAutoKill) & " " & _
            .ScanAutoKill & " " & _
            BoolToNum(.ScanPE) & " " & _
            BoolToNum(.StartUp)
End With
WriteAllText Data, App.Path & "\Settings.Sun"
End Sub

Public Sub AV_Int()
'Check if running
If App.PrevInstance = True Then End

'Load AV Settings
Call LoadSetting
Call LoadQuarantineItem

'Check for registry run
Call GetStartUpList

'Check for AV Sign file
Call LoadSign

'Set AV Process
AdvanceDEBUGToken (True)

'Enable Self Protection

'Other Action
AvVersion = "Alpha"
End Sub

Public Function BoolToNum(ByVal sVar As Boolean) As Byte
If sVar = True Then BoolToNum = 1
End Function

Public Function NumToBool(ByVal sVar As Byte) As Boolean
If sVar = 1 Then NumToBool = True
End Function

Public Sub ExitAV()
Call SaveSetting
Call SaveQuarantineItem
If AVSettings.StartUp = True Then
    Call AddStartUpProgram("Sun AntiVirus", App.Path & "\" & App.EXEName & ".exe")
Else
    Call DeleteStartUpByFilePath(App.Path & "\" & App.EXEName & ".exe")
End If
End
End Sub
