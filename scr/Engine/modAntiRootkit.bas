Attribute VB_Name = "modAntiRootkit"
'Idea to detect hiden process in user mode was taken from gianghoplus
'Orginal source code on Congdongcviet in C

Public Type HidenWindow
    hwnd As Long
    PID As Long
    Title As String
    ClassName As String
    IsHidenWindow As Boolean
End Type

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Const GW_HWNDNEXT = 2
Public Const WM_CLOSE = &H10
Public Const WM_QUIT = &H12
Public Const WM_DESTROY = &H2

'########### PUBLIC VARIABLE #############
Public CntHidenWindow As Long
Public hItem() As HidenWindow
Public Cnt_hItem As Long
'########### PUBLIC VARIABLE #############

Public Sub GetHiden()
Dim chWnd As Long
Dim cPID As Long
Dim HidenProcess As Boolean
Dim i As Integer

ReDim hItem(0) As HidenWindow 'Bo phan tu o vi tri 0
Cnt_hItem = 0

Call RefreshProcess
chWnd = FindWindow(vbNullString, vbNullString)

While chWnd <> 0
    If GetParent(chWnd) = 0 Then
        GetWindowThreadProcessId chWnd, cPID
        HidenProcess = True
        
        For i = 0 To (CntProcess - 1)
            If ListProcess(i).th32ProcessID = cPID Then HidenProcess = False
        Next
        
        If cPID = 0 Or HidenProcess = True Then
            Cnt_hItem = Cnt_hItem + 1
            ReDim Preserve hItem(Cnt_hItem) As HidenWindow
            hItem(Cnt_hItem).PID = cPID
            hItem(Cnt_hItem).hwnd = chWnd
        End If
    End If
    
    chWnd = GetWindow(chWnd, GW_HWNDNEXT)
Wend

End Sub

Public Sub KillProcessByPID2(ByVal PID As Long)
Dim chWnd As Long
Dim cPID As Long
Dim i As Integer
Dim IsContinue As Boolean

IsContinue = False
chWnd = FindWindow(vbNullString, vbNullString)

While chWnd <> 0
    If GetParent(chWnd) = 0 Then
        GetWindowThreadProcessId chWnd, cPID
        
        If cPID = PID Then
            IsContinue = True
            SendMessage chWnd, WM_CLOSE, 0, 0
            SendMessage chWnd, WM_QUIT, 0, 0
            SendMessage chWnd, WM_DESTROY, 0, 0
        End If
    End If
    
    chWnd = GetWindow(chWnd, GW_HWNDNEXT)
Wend
    If IsContinue = True Then KillProcessByPID2 (PID)
End Sub


