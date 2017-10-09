Attribute VB_Name = "modScanning"
Public Type Infected
    VirusName As String
    VirusType As Byte
    FilePath As String
End Type

Public Type ScanResult
    Clean As Boolean
    VirusType As Byte
    VirusName As String
End Type

'########### PUBLIC VARIABLE #############
Public ScanList() As Infected 'Khong tinh phan tu o vi tri 0
Public ScanFiles As Variant
Public TotalFiles As Variant
Public CurrentFile As String
'########### PUBLIC VARIABLE #############

Dim StopScanNow As Boolean

Private Sub AddToInfected(ByVal VirusName As String, ByVal VirusPath As String, ByVal VirusType As Byte)
    Dim iNum As Long
    
    'Khong tinh phan tu o vi tri 0
    iNum = UBound(ScanList)
    ReDim Preserve ScanList(iNum + 1) As Infected
    With ScanList(iNum + 1)
        .FilePath = VirusPath
        .VirusName = VirusName
        .VirusType = VirusType
    End With
    
End Sub

Public Function ScanFile(ByVal FilePath As String) As ScanResult
Dim sMD5 As New MD5Hash
Dim FileHash As String
Dim FirstValue As Byte

FileHash = sMD5.HashFile(FilePath)
FirstValue = Val("&H" & Left(FileHash, 1))

If VirusSigGroup(FirstValue).SigCnt = 0 Then GoTo FileClean 'There is no Sign for scanning

For i = 0 To (VirusSigGroup(FirstValue).SigCnt - 1)
    If FileHash = VirusSigGroup(FirstValue).VirusSig(i).MD5Hash Then 'File is malware
        With ScanFile
          .Clean = False
          .VirusName = VirusSigGroup(FirstValue).VirusSig(i).Name
          .VirusType = VirusSigGroup(FirstValue).VirusSig(i).Type
        End With
        Exit Function
    End If
Next

FileClean:
'File is clean
ScanFile.Clean = True
End Function

Public Sub ScanFolder(ByVal FolderPath As String, ByVal PE_check As Boolean) 'failed
'Su dung thuat toan de quy de liet ke cac file trong folder va scan
On Error GoTo ExitSub


Dim fso As New FileSystemObject
Dim fil As File
Dim FSfolder As Folder
Dim subfolder As Folder
Dim sRet As ScanResult
Dim Temp As Long

For Each fil In fso.GetFolder(FolderPath).Files
    If StopScanNow = True Then Exit Sub
    'Call Doevents in IDE, In complied EXE use CreateThread for more effect (May crash)
    CurrentFile = FolderPath & fil.Name
    
    If IsPE(CurrentFile) = False And PE_check = True Then GoTo SkipScan
    
    sRet = ScanFile(CurrentFile)
    If sRet.Clean = False Then
        AddToInfected sRet.VirusName, CurrentFile, sRet.VirusType
    End If
SkipScan:
    ScanFiles = ScanFiles + 1
    DoEvents
Next

Set FSfolder = fso.GetFolder(FolderPath)
 
For Each subfolder In FSfolder.SubFolders
    If StopScanNow = True Then Exit Sub
    ScanFolder subfolder & "\", PE_check
Next subfolder

Set FSfolder = Nothing

ExitSub:
End Sub

Public Sub StartScan(ByVal FolderPath As String, ByVal PE_check)
    StopScanNow = False
    ScanFiles = 0
    TotalFiles = GetTotalFileInFolder(FolderPath)
    ReDim ScanList(0) As Infected
    Call ScanFolder(FolderPath, PE_check)
End Sub

Public Sub StopScan()
    StopScanNow = True
End Sub

Public Function GetTotalFileInFolder(ByVal FolderPath As String) As Long
Dim fso As New FileSystemObject
Dim FSfolder As Folder
Dim subfolder As Folder
On Error GoTo ExitSub

Set FSfolder = fso.GetFolder(FolderPath)

GetTotalFileInFolder = GetTotalFileInFolder + FSfolder.Files.Count
For Each subfolder In FSfolder.SubFolders
    GetTotalFileInFolder = GetTotalFileInFolder + GetTotalFileInFolder(subfolder & "\")
    DoEvents
Next subfolder

ExitSub:
End Function

Public Function IsPE(ByVal FilePath As String) As Boolean
On Error Resume Next 'If File Path is invalid

   Dim ret As Boolean
   Dim f As Long
   Dim magic As String * 2

   ret = False
   f = FreeFile()
   Open FilePath For Binary As #f
   Get #f, , magic
   Close #f

   If magic = "MZ" Then ret = True
   IsPE = ret

End Function
