VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUN Antivirus - ALPHA Version"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSun.frx":57E2
   ScaleHeight     =   7455
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmRealTime 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11160
      Top             =   0
   End
   Begin VB.Frame FraStt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8400
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
      Begin VB.Label lblver 
         BackStyle       =   0  'Transparent
         Caption         =   "ALPHA Testing"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lbllic 
         BackStyle       =   0  'Transparent
         Caption         =   "FREE VERSION"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbldbver 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed to"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Database"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FraMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8055
      Begin VB.Label lblrtstt 
         BackStyle       =   0  'Transparent
         Caption         =   "*Realtime protection is enabled*"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   4920
         TabIndex        =   85
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         MouseIcon       =   "FrmSun.frx":6DC2
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Image Image10 
         Height          =   585
         Left            =   4130
         Picture         =   "FrmSun.frx":6F14
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "User Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         MouseIcon       =   "FrmSun.frx":805A
         MousePointer    =   99  'Custom
         TabIndex        =   83
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Image Image9 
         Height          =   585
         Left            =   240
         Picture         =   "FrmSun.frx":81AC
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label lblupdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         MouseIcon       =   "FrmSun.frx":8760
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Image Image8 
         Height          =   705
         Left            =   4080
         Picture         =   "FrmSun.frx":88B2
         Top             =   3360
         Width           =   705
      End
      Begin VB.Label lblstartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Startup List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         MouseIcon       =   "FrmSun.frx":8EB7
         MousePointer    =   99  'Custom
         TabIndex        =   81
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   240
         Picture         =   "FrmSun.frx":9009
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Registry Fixer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         MouseIcon       =   "FrmSun.frx":94F2
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Image Image6 
         Height          =   795
         Left            =   4080
         Picture         =   "FrmSun.frx":9644
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblprocess 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Process Explorer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         MouseIcon       =   "FrmSun.frx":9C95
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Image Image5 
         Height          =   585
         Left            =   240
         Picture         =   "FrmSun.frx":9DE7
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lblsetting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         MouseIcon       =   "FrmSun.frx":A3D2
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   4080
         Picture         =   "FrmSun.frx":A524
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblquarantine 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quarantine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   960
         MouseIcon       =   "FrmSun.frx":AA08
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Image Image3 
         Height          =   660
         Left            =   240
         Picture         =   "FrmSun.frx":AB5A
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblrealtime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realtime Protection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         MouseIcon       =   "FrmSun.frx":C24C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   4200
         Picture         =   "FrmSun.frx":C39E
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanning Now"
         DragIcon        =   "FrmSun.frx":C810
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   960
         MouseIcon       =   "FrmSun.frx":C962
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   240
         Picture         =   "FrmSun.frx":CAB4
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame FraRegistry 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   49
      Top             =   1320
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4200
         TabIndex        =   80
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4200
         TabIndex        =   79
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   4200
         TabIndex        =   78
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   4200
         TabIndex        =   77
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   4200
         TabIndex        =   76
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   4200
         TabIndex        =   75
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   4200
         TabIndex        =   74
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4200
         TabIndex        =   73
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   4200
         TabIndex        =   72
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   4200
         TabIndex        =   71
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   4200
         TabIndex        =   70
         Top             =   4320
         Width           =   2415
      End
      Begin sun.Abutton cmlchkallreg 
         Height          =   375
         Left            =   3720
         TabIndex        =   67
         Top             =   5280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Check All"
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
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   4200
         TabIndex        =   66
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   65
         Top             =   0
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   64
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   63
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   62
         Top             =   4320
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   61
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   60
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   59
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   58
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   57
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   56
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Updating..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Folder Options Menu"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Hidden Files And Folders"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fix Registry Is Disabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox ckregistry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fix Task Manager Is Disabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   0
         Width           =   2415
      End
      Begin sun.Abutton cmdunchkallreg 
         Height          =   375
         Left            =   5160
         TabIndex        =   68
         Top             =   5280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Uncheck All"
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
      Begin sun.Abutton cmdfixreg 
         Height          =   375
         Left            =   6600
         TabIndex        =   69
         Top             =   5280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Fix"
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
   End
   Begin VB.Frame FraScanning 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   8175
      Begin VB.Timer tmscan 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3120
         Top             =   4680
      End
      Begin sun.Abutton cmsbrow 
         Height          =   375
         Left            =   7080
         TabIndex        =   26
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Browse"
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
      Begin VB.TextBox txtpath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   6735
      End
      Begin VB.Timer tmprg 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2640
         Top             =   4680
      End
      Begin sun.ProgressBar PrgScan 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   450
         BrushStyle      =   0
         Color           =   12937777
         Style           =   1
         Color2          =   12937777
      End
      Begin MSComctlLib.ListView lstscan 
         Height          =   2655
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin sun.Abutton cmdscan 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Start"
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
      Begin sun.Abutton cmdremove 
         Height          =   375
         Left            =   3000
         TabIndex        =   27
         Top             =   5280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Remove All"
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
      Begin sun.Abutton cmdmovqua 
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   5280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Move All To Quarantine"
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
      Begin VB.Label lblfile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ready to scan!"
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   3960
         Width           =   6255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanning"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblmal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Left            =   6720
         TabIndex        =   22
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Malware found"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblnfile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Files scanned"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   1215
      End
   End
   Begin VB.Frame FraQuarantine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      TabIndex        =   86
      Top             =   1320
      Visible         =   0   'False
      Width           =   8295
      Begin sun.Abutton cmddelqua 
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   5280
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
         Caption         =   "Delete this file"
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
      Begin MSComctlLib.ListView LstQuarantine 
         Height          =   4695
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   8281
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin sun.Abutton cmdrestore 
         Height          =   375
         Left            =   6240
         TabIndex        =   89
         Top             =   5280
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
         Caption         =   "Restore this file"
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
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quarantine List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Frame FraSetting 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   8295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   2415
         TabIndex        =   46
         Top             =   2400
         Width           =   2415
         Begin VB.OptionButton oprt_auto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Auto Solve Problem"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton opaskuser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Asking User"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CheckBox cksusprocess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Also Detect Suspect Process"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox ckprotect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Protect Antivirus From Malware (Not Recommend)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   3975
      End
      Begin VB.CheckBox ckrunstart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Antivirus On Window Startup (Recommend)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox ckautoUSB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto Scan USB"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2055
      End
      Begin VB.OptionButton opwait 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wait For User"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1335
      End
      Begin VB.OptionButton opmove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Move To Quarantine Without Asking"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2895
      End
      Begin VB.OptionButton opremove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Remove Malware Without Asking"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4440
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CheckBox ckpecheck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Only Scan PE File"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When Problem Found"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Real Time Protection Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "When Malware Found"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanning Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3000
         Width           =   2655
      End
   End
   Begin VB.Frame FraCPUInfor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      TabIndex        =   94
      Top             =   1320
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox txtCPUInfor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   120
         Width           =   8055
      End
   End
   Begin VB.Frame FraStartUp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   0
      TabIndex        =   91
      Top             =   1320
      Visible         =   0   'False
      Width           =   8295
      Begin MSComctlLib.ListView LstStartUp 
         Height          =   5175
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   9128
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "StartUp List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Label lblback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<<<Back to control panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmSun.frx":CFEC
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblstt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your computer is safe now!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Image ImgSafe 
      Height          =   1890
      Left            =   9000
      Picture         =   "FrmSun.frx":D13E
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUN Antivirus                         "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Protect your computer from riskware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin VB.Image ImgWarn 
      Height          =   2040
      Left            =   8640
      Picture         =   "FrmSun.frx":E3E4
      Top             =   1560
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "FrmSun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFra As Frame

Private Sub ckautoUSB_Click()
AVSettings.AutoScanUSB = NumToBool(ckautoUSB.Value)
End Sub

Private Sub ckpecheck_Click()
AVSettings.ScanPE = NumToBool(ckpecheck.Value)
End Sub

Private Sub ckprotect_Click()
AVSettings.ProtectMyself = NumToBool(ckprotect.Value)
End Sub

Private Sub ckrunstart_Click()
AVSettings.StartUp = NumToBool(ckrunstart.Value)
End Sub

Private Sub cksusprocess_Click()
AVSettings.DetectSusProc = NumToBool(cksusprocess.Value)
End Sub

Private Sub cmddelqua_Click()
If LstQuarantine.ListItems.Count = 0 Then Exit Sub
DeleteQuarantine (LstQuarantine.SelectedItem.Index - 1)
MsgBox "Deleted File!", vbInformation, "Successful"
Call LoadQuarantineToLst
End Sub

Private Sub cmdmovqua_Click()
Dim TotalVirus As Long
Dim i As Long

TotalVirus = UBound(ScanList)
If TotalVirus = 0 Then
    MsgBox "No malware found", , "_Error"
    Exit Sub
End If

For i = 1 To TotalVirus
    Call AddQuarantine(ScanList(i).FilePath, ScanList(i).VirusName)
Next
ChangeAVSTT (True)
lstscan.ListItems.Clear

MsgBox "All Malware Have Been Moved To Quarantine", vbInformation, "Sucessful!"

ReDim ScanList(0) As Infected
End Sub

Private Sub cmdremove_Click()
Dim TotalVirus As Long
Dim RemovedVirus As Long
Dim i As Long

TotalVirus = UBound(ScanList)
If TotalVirus = 0 Then
    MsgBox "No malware found", , "_Error"
    Exit Sub
End If

For i = 1 To TotalVirus
    If KillFile(ScanList(i).FilePath) = True Then RemovedVirus = RemovedVirus + 1
Next
ChangeAVSTT (True)
lstscan.ListItems.Clear
If RemovedVirus = TotalVirus Then
    MsgBox "Total Viruses:" & TotalVirus & vbCrLf & _
            "Cleaned Viruses:" & RemovedVirus & vbCrLf & _
            "All Malware Removed!", vbInformation, "Sucessful!"
Else
    MsgBox "Total Viruses:" & TotalVirus & vbCrLf & _
            "Cleaned Viruses:" & RemovedVirus & vbCrLf & _
            "Antivirus Can Not Clean All Virus. Please Report To Author" & vbCrLf, vbExclamation, "Failed"
End If

ReDim ScanList(0) As Infected
End Sub

Private Sub cmdrestore_Click()
If LstQuarantine.ListItems.Count = 0 Then Exit Sub
RestoreQuarantine (LstQuarantine.SelectedItem.Index - 1)
MsgBox "Restored File!", vbInformation, "Successful"
Call LoadQuarantineToLst
End Sub

Private Sub cmdscan_Click()
If AVSettings.IsScanning = False Then
    If UBound(ScanList) > 0 Then
        If MsgBox("Start scan will skip old scan result. Are you sure?", vbExclamation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    ChangeAVSTT (True)
    If txtpath.Text = vbNullString Then
        MsgBox "Please Choose a Folder to Scan", vbCritical, "_Error"
        Exit Sub
    End If
    cmdscan.Caption = "Stop"
    tmscan.Enabled = True
Else
    StopScan
    tmscan.Enabled = True
End If
End Sub

Private Sub cmsbrow_Click()
txtpath.Text = BrowseForFolder(Me.hwnd, "Please select a folder to scan")
End Sub

Private Sub Form_Load()
Call AV_Int
Call IntMainForm
Call SetSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1 'No Exit
Call SystrayOn(FrmTray, "Sun AntiVirus")
FrmSun.Hide
End Sub

Private Sub Label14_Click()
FraCPUInfor.Visible = True
lblback.Visible = True
CurrentFra.Visible = False
Set CurrentFra = FraCPUInfor
If GetSysInfo = False Then
    txtCPUInfor.Text = "Failed to get Information!"
Else
    txtCPUInfor.Text = "Your computer information :" & vbCrLf & vbCrLf & _
         "The OS Version = " & OS_Version & vbCrLf & _
         "The OS Build = " & OS_Build & vbCrLf & _
         "The OS ServicePack = " & OS_ServicePack & vbCrLf & _
         "The # of CPUs installed = " & CPU_Count & vbCrLf & _
         "The Type of CPU = " & CPU_Type & vbCrLf & _
         "The CPU Speed = " & CPU_Speed & vbCrLf & _
         "The Installed RAM = " & CPU_RAM
End If
End Sub

Public Sub Label15_Click()
FrmAbout.Show
End Sub

Private Sub lblback_Click()
lblback.Visible = False
CurrentFra.Visible = False
Set CurrentFra = FraMain
FraMain.Visible = True

End Sub

Private Sub lblprocess_Click()
FrmProcess.Show
End Sub

Private Sub lblquarantine_Click()
'Load Quarantine ListView
Call LoadQuarantineToLst

FraQuarantine.Visible = True
lblback.Visible = True
CurrentFra.Visible = False
Set CurrentFra = FraQuarantine
End Sub

Public Sub lblrealtime_Click()
'FrmRealTime.Show
If AVSettings.RealTime = False Then
    lblrtstt.Caption = "*Realtime protection is enabled*"
    lblrtstt.ForeColor = &HFF00&
    AVSettings.RealTime = True
    TmRealTime.Enabled = True
Else
    lblrtstt.Caption = "*Realtime protection is disabled*"
    lblrtstt.ForeColor = &HFF&
    AVSettings.RealTime = False
    TmRealTime.Enabled = False
End If
End Sub

Private Sub lblregistry_Click()
MsgBox "Registry Fixer is not available for Alphal Version", vbInformation, "Information"
'FraRegistry.Visible = True
'lblback.Visible = True
'CurrentFra.Visible = False
'Set CurrentFra = FraRegistry
End Sub

Private Sub lblscan_Click()
FraScanning.Visible = True
lblback.Visible = True
CurrentFra.Visible = False
Set CurrentFra = FraScanning
End Sub


Private Sub lblsetting_Click()
FraSetting.Visible = True
lblback.Visible = True
CurrentFra.Visible = False
Set CurrentFra = FraSetting
Call SetSettings
End Sub

Private Sub lblstartup_Click()
FraStartUp.Visible = True
lblback.Visible = True
CurrentFra.Visible = False
Set CurrentFra = FraStartUp
Call LoadStartUpList 'Test
End Sub

Private Sub lblupdate_Click()
MsgBox "Update is not available for Alphal version", vbInformation, "Information"
End Sub

Private Sub opaskuser_Click()
AVSettings.RTAutoKill = False
End Sub

Private Sub opmove_Click()
AVSettings.ScanAutoKill = MoveToQuarantine
End Sub

Private Sub opremove_Click()
AVSettings.ScanAutoKill = RemoveMalware
End Sub

Private Sub oprt_auto_Click()
AVSettings.RTAutoKill = True
End Sub

Private Sub opwait_Click()
AVSettings.ScanAutoKill = Wait
End Sub

Private Sub tmprg_Timer()
'This timer will update scan progress
'On Error Resume Next

Dim iNum As Long
Dim i As Byte

If ScanFiles > 0 Then
lblfile.Caption = CurrentFile
PrgScan.Value = (ScanFiles / TotalFiles) * 100
End If

lblnfile.Caption = ScanFiles
iNum = UBound(ScanList)
lblmal.Caption = iNum
If iNum <> 0 And ScanFiles <> 0 Then
    If lstscan.ListItems.Count < iNum Then
        ChangeAVSTT (False)
        For i = (lstscan.ListItems.Count + 1) To iNum
            Call AddToVirusList(ScanList(i).VirusName, ScanList(i).FilePath)
        Next
    End If
End If
End Sub

Private Sub StartScanning()
    'Int value
    
    AVSettings.IsScanning = True
    FrmSun.tmprg.Enabled = True 'Timer update scanning progress
    lblfile.Caption = "Please wait!"
    lstscan.ListItems.Clear
    PrgScan.Value = 0
    
    'Scan Thread
    If Right(txtpath.Text, 1) <> "\" Then
        Call StartScan(txtpath.Text & "\", AVSettings.ScanPE)
    Else
        Call StartScan(txtpath.Text, AVSettings.ScanPE)
    End If
    
    'Scanning done
    FrmSun.tmprg.Enabled = False
    AVSettings.IsScanning = False
    lblfile.Caption = "Scan finished"
    lblnfile.Caption = ScanFiles
    lblmal.Caption = UBound(ScanList)
    PrgScan.Value = 100
    cmdscan.Caption = "Start"
    Call LoadInfectedToLst
    
    If AVSettings.ScanAutoKill = MoveToQuarantine Then
        Call cmdmovqua_Click
    ElseIf AVSettings.ScanAutoKill = RemoveMalware Then
        Call cmdremove_Click
    End If
End Sub

Private Sub tmscan_Timer()
Call StartScanning
tmscan.Enabled = False
End Sub

Private Sub TmRealTime_Timer()
'Scan 5s 1 lan

If FormCount("FrmRealTime") > 0 Then Exit Sub
Call DetectMalwareProcess
If RealTimeCnt <> 0 Then
    FrmRealTime.DetectedItem = Malware
    FrmRealTime.Show
ElseIf AVSettings.DetectSusProc = True Then
    Call DetectInjectedProcess
    If modRealTime.InjectedProcess.th32ProcessID <> 0 Then
        FrmRealTime.DetectedItem = SusProc
        FrmRealTime.Show
    End If
End If
End Sub

Public Sub IntMainForm()
FrmSun.lbldbver.Caption = VirusData.LastUpdate
FrmSun.Caption = "SUN Antivirus - ALPHA"
Set CurrentFra = FraMain

With FrmSun.lstscan
    .View = lvwReport
    .Arrange = lvwNone
    .LabelEdit = lvwManual
    .HideColumnHeaders = False
    .HideSelection = False
    .LabelWrap = False
    .MultiSelect = False
    .Enabled = True
    .AllowColumnReorder = True
    .Checkboxes = False
    .FlatScrollBar = False
    .FullRowSelect = True
    .GridLines = True
    .HotTracking = False
    .HoverSelection = False
    .Sorted = True
    .SortKey = 0
    .SortOrder = lvwAscending 'lvwDescending
    .ColumnHeaders.Add , , "Virus Name", .Width / 4
    .ColumnHeaders.Add , , "File Path", .Width - .Width / 4
End With


With FrmSun.LstQuarantine
    .View = lvwReport
    .Arrange = lvwNone
    .LabelEdit = lvwManual
    .HideColumnHeaders = False
    .HideSelection = False
    .LabelWrap = False
    .MultiSelect = False
    .Enabled = True
    .AllowColumnReorder = True
    .Checkboxes = False
    .FlatScrollBar = False
    .FullRowSelect = True
    .GridLines = True
    .HotTracking = False
    .HoverSelection = False
    .Sorted = True
    .SortKey = 0
    .SortOrder = lvwAscending 'lvwDescending
    .ColumnHeaders.Add , , "Date", .Width / 6
    .ColumnHeaders.Add , , "Virus Name", .Width / 6
    .ColumnHeaders.Add , , "File Path", .Width / 2 + .Width / 6
End With

With FrmSun.LstStartUp
    .View = lvwReport
    .Arrange = lvwNone
    .LabelEdit = lvwManual
    .HideColumnHeaders = False
    .HideSelection = False
    .LabelWrap = False
    .MultiSelect = False
    .Enabled = True
    .AllowColumnReorder = True
    .Checkboxes = False
    .FlatScrollBar = False
    .FullRowSelect = True
    .GridLines = True
    .HotTracking = False
    .HoverSelection = False
    .Sorted = True
    .SortKey = 0
    .SortOrder = lvwAscending 'lvwDescending
    .ColumnHeaders.Add , , "Type", .Width / 4
    .ColumnHeaders.Add , , "File Path", .Width - .Width / 4
End With

Call SystrayOn(FrmTray, "Sun AntiVirus")
FrmSun.Hide
Call PopupBalloon(FrmTray, "Written by Elvis", "SUN Antivirus - " & AvVersion)
Call Sleep(1500)
Call RemoveBalloon(FrmTray)
End Sub

Public Sub ChangeAVSTT(ByVal Safe As Boolean)

If Safe = True Then
    FrmSun.ImgSafe.Visible = True
    FrmSun.ImgWarn.Visible = False
    FrmSun.lblstt.Caption = "Your computer is safe now!"
    FrmSun.lblstt.ForeColor = &HFF00&
Else
    FrmSun.ImgSafe.Visible = False
    FrmSun.ImgWarn.Visible = True
    FrmSun.lblstt.Caption = "Threat(s) found!"
    FrmSun.lblstt.ForeColor = &HFF&
End If

End Sub

Private Sub SetSettings()
With AVSettings
    ckrunstart.Value = BoolToNum(.StartUp)
    ckprotect.Value = BoolToNum(.ProtectMyself)
    cksusprocess.Value = BoolToNum(.DetectSusProc)
    ckpecheck.Value = BoolToNum(.ScanPE)
    ckautoUSB.Value = BoolToNum(.AutoScanUSB)
    If .RTAutoKill = True Then
        oprt_auto.Value = True
    Else
        opaskuser.Value = True
    End If
    
    Select Case .ScanAutoKill
        Case 0
            opremove.Value = True
        Case 1
            opmove.Value = True
        Case 2
            opwait.Value = True
    End Select
End With

If AVSettings.RealTime = True Then
    lblrtstt.Caption = "*Realtime protection is enabled*"
    lblrtstt.ForeColor = &HFF00&
    TmRealTime.Enabled = True
Else
    lblrtstt.Caption = "*Realtime protection is disabled*"
    lblrtstt.ForeColor = &HFF&
    TmRealTime.Enabled = False
End If
End Sub

Private Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function

Private Sub AddToQuarantineList(sData As Quarantine_Item)
Dim li As ListItem

With FrmSun.LstQuarantine
    Set li = .ListItems.Add(, , sData.DateTime)
    li.SubItems(1) = sData.VirusName
    li.SubItems(2) = sData.OrginalPath
End With
End Sub

Private Sub LoadQuarantineToLst()
Dim i As Byte
LstQuarantine.ListItems.Clear
If modQuarantine.CntQuarantine = 0 Then Exit Sub
For i = 0 To (modQuarantine.CntQuarantine - 1)
    Call AddToQuarantineList(modQuarantine.ListQuarantine(i))
Next
End Sub

Private Sub AddToVirusList(ByVal VirusName As String, ByVal Path As String)
Dim li As ListItem

With FrmSun.lstscan
    Set li = .ListItems.Add(, , VirusName)
    li.SubItems(1) = Path
End With
End Sub

Private Sub LoadInfectedToLst()
Dim i As Byte
lstscan.ListItems.Clear
If UBound(ScanList) = 0 Then Exit Sub
ChangeAVSTT (False)
For i = 1 To UBound(ScanList)
    Call AddToVirusList(ScanList(i).VirusName, ScanList(i).FilePath)
Next
End Sub

Private Sub AddStartUpList(ByVal stType As StartUp_Type, ByVal ExePath As String)
Dim li As ListItem
Dim strType As String

Select Case stType
    Case 0
        strType = "Registry"
    Case 1
        strType = "Startup Folder"
    Case 2
        strType = "Win.INI"
End Select

With FrmSun.LstStartUp
    Set li = .ListItems.Add(, , strType)
    li.SubItems(1) = ExePath
End With
End Sub

Private Sub LoadStartUpList()
Dim i As Byte
FrmSun.LstStartUp.ListItems.Clear
Call GetStartUpList
If CntStartUp = 0 Then Exit Sub
For i = 0 To (CntStartUp - 1)
    Call AddStartUpList(StartUpList(i).StartUpType, StartUpList(i).ExePath)
Next
End Sub
