Attribute VB_Name = "modStartUp"
Public Enum StartUp_Type
    RegStart = 0
    StartUpFolder = 1
    Winini = 2
End Enum

Public Type StartUp_List
    StartUpType As StartUp_Type
    ExePath As String
    KeyName As String 'For Registry StartUp
    KeyLocation As Long 'For Registry StartUp
End Type

Public StartUpList() As StartUp_List
Public CntStartUp As Byte

Public Sub GetStartUpList()
CntStartUp = 0
Call GetRegistryStartUp
Call GetFolderStartUp
End Sub

Public Sub GetRegistryStartUp()

Dim ret As Variant
Dim NullVariant As Variant
Dim StartUpKey As String
Dim i As Byte
Dim Temp As String

StartUpKey = "software\microsoft\windows\currentversion\run"

ret = GetAllValues(HKEY_CURRENT_USER, StartUpKey)
If IsArrayInitialized(ret) Then
For i = 0 To UBound(ret)
    ReDim Preserve StartUpList(CntStartUp) As StartUp_List
    With StartUpList(CntStartUp)
        .StartUpType = RegStart
        Temp = ret(i, 0)
        .ExePath = GetSettingString(HKEY_CURRENT_USER, StartUpKey, Temp)
        .KeyName = Temp
        .KeyLocation = HKEY_CURRENT_USER
    End With
    CntStartUp = CntStartUp + 1
Next
End If


ret = GetAllValues(HKEY_LOCAL_MACHINE, StartUpKey)
If IsArrayInitialized(ret) = True Then
For i = 0 To UBound(ret)
    ReDim Preserve StartUpList(CntStartUp) As StartUp_List
    With StartUpList(CntStartUp)
        .StartUpType = RegStart
        Temp = ret(i, 0)
        .ExePath = GetSettingString(HKEY_LOCAL_MACHINE, StartUpKey, Temp)
        .KeyName = Temp
        .KeyLocation = HKEY_LOCAL_MACHINE
    End With
    CntStartUp = CntStartUp + 1
Next
End If


End Sub

Private Function GetStartUpPath() As String
Dim mShell
Set mShell = CreateObject("WScript.Shell")
GetStartUpPath = mShell.SpecialFolders("Startup")
End Function

Public Sub GetFolderStartUp()
Dim fso As New FileSystemObject
Dim fld As Folder
Dim fil As File
Dim StartPath As String

StartPath = GetStartUpPath
Set fld = fso.GetFolder(StartPath)
For Each fil In fld.Files
    If LCase(fil.Name) <> "desktop.ini" Then
        ReDim Preserve StartUpList(CntStartUp) As StartUp_List
        With StartUpList(CntStartUp)
            .StartUpType = StartUpFolder
            .ExePath = StartPath & "\" & fil.Name
        End With
        CntStartUp = CntStartUp + 1
    End If
Next
Set fil = Nothing
Set fld = Nothing
Set fso = Nothing
End Sub

Public Sub GetWinINIStartUp()
'On Window vista,7,8,10 this file was removed

End Sub

Public Sub AddStartUpProgram(ByVal kName As String, ByVal FilePath As String)
Call SaveSettingString(HKEY_CURRENT_USER, "software\microsoft\windows\currentversion\run", kName, FilePath)
Call GetStartUpList 'Refresh StartUp List
End Sub

Public Sub DeleteStartUpByID(ByVal ID As Long)
Dim i As Byte
Select Case StartUpList(ID).StartUpType
    Case 0 'Registry
        Call DeleteValue(StartUpList(ID).KeyLocation, "software\microsoft\windows\currentversion\run", StartUpList(ID).KeyName)
    Case 1 'StartUp Folder
        Call KillFile(StartUpList(ID).ExePath)
    Case 2 'WinINI
End Select

For i = ID To (CntStartUp - 2)
    StartUpList(i) = StartUpList(i + 1)
Next
ReDim Preserve StartUpList(CntStartUp - 2)
CntStartUp = CntStartUp - 1
End Sub


Public Sub DeleteStartUpByFilePath(ByVal FilePath As String)
Dim Temp As String
Dim i As Byte

If CntStartUp = 0 Then Exit Sub
Temp = UCase(FilePath)
For i = 0 To (CntStartUp - 1)
    If InStr(Temp, UCase(StartUpList(i).ExePath)) <> 0 Then
        DeleteStartUpByID (i)
    End If
Next
End Sub
