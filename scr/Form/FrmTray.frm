VERSION 5.00
Begin VB.Form FrmTray 
   Caption         =   "Tray"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "FrmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Main 
      Caption         =   "Main"
      Begin VB.Menu Show 
         Caption         =   "&Show AntiVirus"
      End
      Begin VB.Menu Tool 
         Caption         =   "&Window Tool"
         Begin VB.Menu Tool1 
            Caption         =   "&Window Task Manager"
         End
         Begin VB.Menu Tool2 
            Caption         =   "&Regedit"
         End
         Begin VB.Menu Tool3 
            Caption         =   "&CMD"
         End
      End
      Begin VB.Menu A 
         Caption         =   "-"
      End
      Begin VB.Menu RT_Protection 
         Caption         =   "RealTime Protection"
      End
      Begin VB.Menu StartUp 
         Caption         =   "StartUp"
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
FrmAbout.Show
End Sub

Private Sub Exit_Click()
Call ExitAV
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Dim sFilter As String
    
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            FrmSun.Show ' show form
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
            RT_Protection.Checked = AVSettings.RealTime
            StartUp.Checked = AVSettings.StartUp
            Me.PopupMenu Main
        Case WM_RBUTTONDBLCLK
        Case WM_MOUSEMOVE
    End Select
End Sub

Private Sub RT_Protection_Click()
Call FrmSun.lblrealtime_Click
End Sub

Private Sub Show_Click()
    FrmSun.Show
End Sub

Private Sub StartUp_Click()
    If AVSettings.StartUp = True Then
        AVSettings.StartUp = False
    Else
        AVSettings.StartUp = True
    End If
End Sub

Private Sub Tool1_Click()
On Error Resume Next
Shell "taskmgr.exe", vbNormalFocus
End Sub

Private Sub Tool2_Click()
On Error Resume Next
Shell "regedit.exe", vbNormalFocus
End Sub

Private Sub Tool3_Click()
On Error Resume Next
Shell "cmd.exe", vbNormalFocus
End Sub


