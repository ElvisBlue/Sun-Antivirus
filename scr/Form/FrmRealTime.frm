VERSION 5.00
Begin VB.Form FrmRealTime 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   14715
   ClientTop       =   7500
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin sun.Abutton cmdok 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picauto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1800
      ScaleHeight     =   1335
      ScaleWidth      =   3975
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
      Begin VB.PictureBox Picask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   3975
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton opleave 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Do nothing"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opquarantine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Move to Quarantine"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton ophandle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Delete File / Terminate Process"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Do you want to?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   3135
         End
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theat has been removed!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
   End
   Begin sun.Abutton cmdcan 
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmpopup 
      Interval        =   10
      Left            =   240
      Top             =   2640
   End
   Begin VB.Label lbldetectinfor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detected"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblprocname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Process name"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   120
      Picture         =   "FrmRealTime.frx":0000
      Top             =   600
      Width           =   1560
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00B17F3C&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblcaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "RealTime Protection: Malware Process found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B17F3C&
      Height          =   3615
      Left            =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "FrmRealTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum PopUpState
    GoUp = 0
    Stay = 1
    GoDown = 2
End Enum

Enum Detected_Item
    Malware = 0
    SusProc = 1
End Enum

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Dim RTFormState As PopUpState
Public DetectedItem As Detected_Item

Private Sub cmdcan_Click()
RTFormState = GoDown
End Sub

Private Sub cmdok_Click()
If AVSettings.RTAutoKill = False Then
    If ophandle.Value = True Then
        DeleteLastRTItem (True)
    End If
    
    If opquarantine.Value = True Then
        Call SetQuarantineLastRTItem
    End If
End If

Call cmdcan_Click
End Sub

Private Sub Form_Load()
Call SetUpForm
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width
RTFormState = GoUp
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
End Sub

Private Sub tmpopup_Timer()
If RTFormState = GoUp Then
    Me.Top = Me.Top - 40
    If Me.Top <= Screen.Height - Me.Height Then RTFormState = Stay
End If

If RTFormState = GoDown Then
    Me.Top = Me.Top + 40
    If Me.Top >= Screen.Height Then Unload Me
End If

End Sub

Private Sub SetUpForm()

If DetectedItem = Malware Then 'If Malware Process Detected
    lblcaption.Caption = "RealTime Protection: Malware Process found"
    lblprocname.Caption = RealTimeList(RealTimeCnt - 1).Detected.FilePath
    lblprocname.ToolTipText = lblprocname.Caption
    lbldetectinfor.Caption = RealTimeList(RealTimeCnt - 1).Detected.VirusName
    If AVSettings.RTAutoKill = True Then
        Picask.Visible = False
        DeleteLastRTItem (True)
        lblMsg.Caption = "Theat has been removed!"
    Else
        Picask.Visible = True
    End If
ElseIf DetectedItem = SusProc Then
    lblcaption.Caption = "RealTime Protection: Suspect Process found"
    Picask.Visible = False
    lblMsg.Caption = "Suspect Process has been terminated"
    lbldetectinfor.Caption = "Suspect Process"
    lblprocname.Caption = modRealTime.InjectedProcess.szExeFile
    Call KillProcessByPID(modRealTime.InjectedProcess.th32ProcessID)
End If
End Sub
