Attribute VB_Name = "modQuarantine"
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Public Type Quarantine_Item
    OrginalPath As String
    EncryptedPath As String
    DateTime As String
    VirusName As String
End Type

Public ListQuarantine() As Quarantine_Item
Public CntQuarantine As Integer

Public Sub LoadQuarantineItem()
'On Error Resume Next
Dim strSave As String
Dim Temp() As String
Dim Temp2() As String
Dim i As Integer

If Dir(App.Path & "\Quarantine", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Quarantine\")
End If

strSave = modOther.ReadAllText(App.Path & "\Quarantine.Sun")
strSave = StripVBCrLf(strSave)
If strSave = vbNullString Then Exit Sub
Temp = Split(strSave, "%t%t")
CntQuarantine = UBound(Temp) + 1
ReDim ListQuarantine(UBound(Temp)) As Quarantine_Item
For i = 0 To UBound(Temp)
    Temp2 = Split(Temp(i), "%s%s")
    With ListQuarantine(i)
        .DateTime = Temp2(0)
        .EncryptedPath = Temp2(1)
        .OrginalPath = Temp2(2)
        .VirusName = Temp2(3)
    End With
Next
End Sub

Public Sub SaveQuarantineItem()
Dim strSave As String
Dim Temp As String

Dim i As Integer

For i = 0 To (CntQuarantine - 1)
    Temp = ListQuarantine(i).DateTime & "%s%s" & _
            ListQuarantine(i).EncryptedPath & "%s%s" & _
            ListQuarantine(i).OrginalPath & "%s%s" & _
            ListQuarantine(i).VirusName
    strSave = strSave & Temp
    If i <> CntQuarantine - 1 Then
        strSave = strSave & "%t%t"
    End If
Next
Call modOther.WriteAllText(strSave, App.Path & "\Quarantine.Sun")
End Sub

Public Function DeleteQuarantine(ByVal ID As Integer) As Boolean
Dim i As Integer
modRemoveFile.KillFile (ListQuarantine(ID).EncryptedPath)
For i = ID To (CntQuarantine - 2)
    ListQuarantine(i) = ListQuarantine(i + 1)
Next
CntQuarantine = CntQuarantine - 1
If CntQuarantine <> 0 Then ReDim Preserve ListQuarantine(CntQuarantine - 1)
End Function

Public Function AddQuarantine(ByVal Path As String, ByVal VirusName As String) As Boolean
Dim Temp As String
Dim OldSize As Long
'On Error GoTo ErrHandler


Temp = App.Path & "\Quarantine\" & RandomString(20) & ".zip"
If ShellZip(Path, Temp) = False Then
    AddQuarantine = False
End If

'While FileExists(Temp) = False
'    Call Sleep(100)
'Wend

OldSize = 0
FileCheck:
While OldSize <> GetFileSize(Temp) Or OldSize = 0
    OldSize = GetFileSize(Temp)
    Call Sleep(100)
Wend


ReDim Preserve ListQuarantine(CntQuarantine)

With ListQuarantine(CntQuarantine)
    .DateTime = Date
    .OrginalPath = Path
    .VirusName = VirusName
    .EncryptedPath = Temp
End With

modRemoveFile.KillFile (Path)
CntQuarantine = CntQuarantine + 1
AddQuarantine = True
    
End Function

Public Function RestoreQuarantine(ByVal ID As Integer) As Boolean
Dim RestoreFolder As String
Dim PathFile As String

PathFile = ListQuarantine(ID).OrginalPath
RestoreFolder = Left(PathFile, InStrRev(PathFile, "\"))
Call ShellUnzip(ListQuarantine(ID).EncryptedPath, RestoreFolder)
Call DeleteQuarantine(ID)
End Function

