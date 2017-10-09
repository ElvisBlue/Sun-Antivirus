VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProcess 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Process Explorer"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "FrmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkhiden 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Try to detect hidden process"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin sun.Abutton cmdrefresh 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Refresh"
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
   Begin MSComctlLib.ListView lstprocess 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin sun.Abutton cmdkill 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Kill Process"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Process Explorer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdkill_Click()
Dim Selected As Byte
Dim PID As Long
'If KillProcessByPID(Text1.Text) = False Then MsgBox "Failed to terminate process", vbCritical, "Failed"
Selected = lstprocess.SelectedItem.Index
PID = Val(lstprocess.ListItems.Item(Selected).ListSubItems.Item(1).Text)
If MsgBox("Are you sure?", vbYesNo, "Process Killer") = vbYes Then
    Call KillProcessByPID(PID)
    Call KillProcessByPID2(PID)
    Call Sleep(500)
    Call cmdrefresh_Click
End If
End Sub

Private Sub cmdrefresh_Click()
Dim i As Byte
Dim j As Byte
Dim Added As Boolean
Dim ProccessName As String

lstprocess.ListItems.Clear
Call RefreshProcess
For i = 1 To CntProcess
    Call AddToProcessList(ListProcess(i - 1).th32ProcessID, ListProcess(i - 1).szExeFile, vbBlack)
Next

If chkhiden.Value = 1 Then
    Call GetHiden
    For i = 1 To Cnt_hItem
        Added = False
        For j = 1 To i - 1
            If hItem(j).PID = hItem(i).PID Then Added = True
        Next
        If Added = False Then
            ProcessName = GetFileNameFromPath(ProcessPathByPID(hItem(i).PID))
            If ProcessName = vbNullString Then ProcessName = "<Unknown Process>"
            Call AddToProcessList(hItem(i).PID, ProcessName, vbRed)
        End If
    Next
End If
End Sub

Private Sub Form_Load()
Call IntProcessForm
Call cmdrefresh_Click
End Sub

Public Sub AddToProcessList(ByVal PID As Long, ByVal ProcessName As String, ByVal Color As OLE_COLOR)
Dim li As ListItem

With FrmProcess.lstprocess
    Set li = .ListItems.Add(, , ProcessName)
    li.SubItems(1) = PID
    li.ForeColor = Color
End With
End Sub

Public Sub IntProcessForm()
FrmProcess.Picture = FrmSun.Picture

With FrmProcess.lstprocess
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
    .ColumnHeaders.Add , , "Process Name", .Width - 1000
    .ColumnHeaders.Add , , "PID", 700
End With

End Sub
