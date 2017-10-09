Attribute VB_Name = "modRemoveFile"

Public Function KillFile(ByVal FilePath As String) As Boolean
On Error GoTo ExitF
KillProcessByPath (FilePath)
SetAttr FilePath, vbNormal
Kill FilePath
KillFile = True
Exit Function

ExitF:
KillFile = False
End Function

