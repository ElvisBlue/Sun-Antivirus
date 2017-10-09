Attribute VB_Name = "modOther"
Option Explicit
 
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
 
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
 
Public Function FileExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function

Public Function GetFileSize(ByVal Path As String) As Long
Dim fso As New FileSystemObject
    Dim f As File
    'Get a reference to the File object.
    If fso.FileExists(Path) Then
        Set f = fso.GetFile(Path)
        GetFileSize = f.Size
    Else
        GetFileSize = 0
    End If

End Function

Public Function RandomString(cb As Integer) As String

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Public Function ReadAllText(strFile As String) As String

On Error GoTo errorhandler
Dim intFile As Integer

intFile = FreeFile
Open strFile For Input As #intFile
    ReadAllText = Input(LOF(intFile), #intFile)
Close #intFile
Exit Function

errorhandler:
    ReadAllText = vbNullString
    Exit Function
End Function

Public Function WriteAllText(ByVal Txt As String, ByVal FilePath As String) As Boolean
On Error GoTo ExitF
Dim iFileNo As Integer
iFileNo = FreeFile
Open FilePath For Output As #iFileNo
    Print #iFileNo, Txt
Close #iFileNo
WriteAllText = True
Exit Function
ExitF:
WriteAllText = False
End Function

Public Function StripNULL(ByVal Txt As String) As String
Dim pos As Byte

If Txt = vbNullString Then Exit Function
pos = InStr(1, Txt, vbNullChar)
If pos = 0 Then Exit Function
StripNULL = Left$(Txt, pos - 1)
End Function

Public Function StripVBCrLf(ByVal Txt As String) As String
Dim pos As Byte

If Txt = vbNullString Then Exit Function
pos = InStr(1, Txt, vbCrLf)
If pos = 0 Then Exit Function
StripVBCrLf = Left$(Txt, pos - 1)
End Function

Public Function IsArrayInitialized(arr) As Boolean
Dim rv As Long
On Error GoTo ErrHandle

rv = UBound(arr)
IsArrayInitialized = True
Exit Function

ErrHandle:
IsArrayInitialized = False
End Function
