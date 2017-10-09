Attribute VB_Name = "modSign"
Public Type Virus_Sig
    Name As String
    Type As Byte
    MD5Hash As String
End Type

Public Type Virus_Data
    SigCnt As Long
    LastUpdate As String
End Type

Public Type VirusSig_Group
    SigCnt As Long
    VirusSig(0 To 5000) As Virus_Sig
End Type

Public Const TROJAN = 0
Public Const WORM = 1
Public Const BACKDOOR = 2
Public Const ADWARE = 3
Public Const RANSOMWARE = 4
Public Const DOWNLOADER = 5
Public Const CRYPTER = 6

'########### PUBLIC VARIABLE #############
Public VirusSigGroup(15) As VirusSig_Group 'From 0 to F
Public VirusData As Virus_Data
'########### PUBLIC VARIABLE #############

Public Sub LoadSign()
    Dim Data As String
    Dim Temp1() As String
    Dim Temp2() As String
    Dim FirstValue As Byte
    
    Data = ReadAllText(App.Path & "\Sun.sig") 'Read Sign File
    If Data = vbNullString Then
        MsgBox "Virus Database not found!", vbExclamation, "Warning"
        Exit Sub
    End If
    Temp1 = Split(Data, vbCrLf)
    VirusData.LastUpdate = Temp1(0)
    VirusData.SigCnt = UBound(Temp1)
    
    For i = 1 To VirusData.SigCnt
        Temp2 = Split(Temp1(i), " | ") ' Struct: MD5 | VirusName | Type
        FirstValue = Val("&H" & Left(Temp2(0), 1))
        
        With VirusSigGroup(FirstValue)
            .SigCnt = .SigCnt + 1
            .VirusSig(.SigCnt - 1).MD5Hash = Temp2(0)
            .VirusSig(.SigCnt - 1).Name = Temp2(1)
            '.VirusSig(.SigCnt - 1).Type = Val(Temp2(2))
        End With
        
    Next
End Sub
