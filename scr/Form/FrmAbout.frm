VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5055
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmtext 
      Enabled         =   0   'False
      Left            =   720
      Top             =   4800
   End
   Begin VB.Timer tmwn 
      Interval        =   50
      Left            =   120
      Top             =   4800
   End
   Begin VB.PictureBox Piccre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   1
      Top             =   1800
      Width           =   5055
   End
   Begin sun.Abutton cmdok 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
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
   Begin VB.Label lblsn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "None"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "License Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Freeware"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registered To:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(2) As Long
End Type
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

'###############################################
Dim CreditLine()    As String
Dim TmpLine()       As String
Dim CreditLeft()    As Long
Dim ColorFades(100) As Long
Dim ScrollSpeed     As Integer
Dim ColText         As Long
Dim FadeIn          As Long
Dim FadeOut         As Long

Dim cDiff1          As Long
Dim cDiff2          As Double
Dim cDiff3          As Double

Dim TotalLines      As Integer
Dim LinesOffset     As Integer
Dim Yscroll         As Long
Dim CharHeight      As Integer
Dim LinesVisible    As Integer
'################################################
Private Counter As Long

Private Sub cmdok_Click()
Unload Me
End Sub


Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    Counter = 0
    tmtext.Enabled = False
End Sub

Private Sub tmtext_Timer()
Dim Ycurr       As Long
Dim TextLine    As Integer
Dim ColPrct     As Long
Dim i           As Integer
'clear pic for next draw
Piccre.Cls
Yscroll = Yscroll - ScrollSpeed
'calculate beginscroll
If Yscroll < (0 - CharHeight) Then
    Yscroll = 0
    LinesOffset = LinesOffset + 1
    If LinesOffset > TotalLines - 1 Then LinesOffset = 0
    'the offset sets the first line of the serie to be printed
    'this offset goes to the next line after each completely
    'scrolled line
    End If
'set Y for first  line
Piccre.CurrentY = Yscroll
Ycurr = Yscroll
'print only the visible lines
For i = 1 To LinesVisible
    If Ycurr > FadeIn And Ycurr < Piccre.Height Then
        'calculate fade-in forecolor
        ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        Piccre.ForeColor = ColorFades(ColPrct)
    ElseIf Ycurr < FadeOut Then
        'calculate fade-out forecolor
        ColPrct = cDiff3 * Ycurr
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        Piccre.ForeColor = ColorFades(ColPrct)
    Else
        'normal forecolor
        Piccre.ForeColor = ColText
    End If
    'get next line with offset
    TextLine = (i + LinesOffset) Mod TotalLines
    'set the X aligne value
    Piccre.CurrentX = CreditLeft(TextLine)
    'print that line
    Piccre.Print CreditLine(TextLine)
    'set Y to print next line
    Ycurr = Ycurr + CharHeight
Next i
End Sub

Private Sub tmwn_Timer()
    Randomize
    Dim RandomBits() As Integer 'The bitmap bits - Long would technically work faster as it would take half the number of loop iterations, but doesn't work well since VB's random number generator goes from .0001 to .9999 so too many digits are ignored
    Dim TheBitmapInfo As BITMAPINFO
    Dim TheWidth As Long, TheHeight As Long
    Dim i As Long 'Loop counter
    
    TheWidth = Piccre.ScaleWidth
    TheHeight = Piccre.ScaleHeight
    ReDim RandomBits(0 To TheWidth * TheHeight)
    With TheBitmapInfo.bmiHeader
        .biSize = Len(TheBitmapInfo.bmiHeader)
        .biWidth = TheWidth
        .biHeight = TheHeight
        .biPlanes = 1
        .biBitCount = 1
        .biClrUsed = 2 '2 colors
        .biClrImportant = 2 '2 colors
    End With
    TheBitmapInfo.bmiColors(0) = &HFFFFFF 'White
    TheBitmapInfo.bmiColors(1) = 0 'Black
    
    'Do
        'Output random data
        For i = 0 To UBound(RandomBits)
            RandomBits(i) = Rnd * 65535 - 32768
        Next i
        
        StretchDIBits Piccre.hdc, 0, 0, TheWidth, TheHeight, 0, 0, TheWidth, TheHeight, RandomBits(0), TheBitmapInfo, 0, vbSrcCopy 'Copy data to form
        'DoEvents 'Listen for mouse click
    Counter = Counter + 1
    If Counter >= 50 Then
        tmwn.Enabled = False
        Piccre.BackColor = &HFFFFFF
        Call IntCredits
    End If
    'Loop
End Sub

Private Sub IntCredits()
Dim Message As String
Dim FileO       As Integer
Dim FileName    As String
Dim tmp         As String
Dim i           As Integer
Dim j As Integer

Dim Rcol1       As Long
Dim Gcol1       As Long
Dim Bcol1       As Long

Dim Rcol2       As Long
Dim Gcol2       As Long
Dim Bcol2       As Long

Dim Rfade       As Long
Dim Gfade       As Long
Dim Bfade       As Long

Dim PercentFade As Integer
Dim TimeInterval As Integer
Dim AlignText  As Integer


    Piccre.AutoRedraw = True
    Piccre.ScaleMode = 3
    
    Message = "Sun Antivirus" & vbCrLf & vbCrLf & _
              "Version: Alpha" & vbCrLf & vbCrLf & _
              "Author: Elvis" & vbCrLf & vbCrLf & vbCrLf & _
              "Sun Antivirus is a small project written in VB6." & vbCrLf & _
              "Sun Antivirus has features popular to another AntiVirus:" & vbCrLf & _
              "Scan virus (MD5 checksum database), Realtime Protection, " & vbCrLf & _
              "Quarantine, Toolkits: Registry fixer, Startup manager," & vbCrLf & _
              "Process Explorer,....This is alphal version on Sun Antivirus" & vbCrLf & _
              "so it may contain bugs" & vbCrLf & _
              "Creadits:" & vbCrLf & vbCrLf & _
              "Dungcoivb for vnAntivirus Source Code" & vbCrLf & vbCrLf & _
              "aGunG adhi satya - ATV Guard Source Code" & vbCrLf & vbCrLf & _
              "D. Rijmenants for cool scrolling text" & vbCrLf & vbCrLf & _
              "Kevin Wilson (The VB Zone) for CPUInfor Module" & vbCrLf & vbCrLf & _
              "Planet Source Code for cool VB Source Codes" & vbCrLf & vbCrLf & _
              "vbforums.com - cool VB forum" & vbCrLf & vbCrLf & _
              "My friends :)" & vbCrLf & vbCrLf
              


    PercentFade = 30

    TimeInterval = 30
    ScrollSpeed = 1

    AlignText = 2 '( 1=left 2=center 3=right )
    
'################################################################

'set the number of line to be printed in the box
LinesVisible = (Piccre.Height / Piccre.TextHeight("A")) + 1

'add empty lines at beginning to start off
For i = 1 To LinesVisible
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
Next


TmpLine = Split(Message, vbCrLf)
i = TotalLines
TotalLines = TotalLines + UBound(TmpLine) + 1
ReDim Preserve CreditLine(TotalLines) As String

While i < TotalLines
CreditLine(i) = TmpLine(j)
i = i + 1
j = j + 1
Wend

'set timer interval
Me.tmtext.Interval = TimeInterval

'set the number of line to be printed in the box
LinesVisible = (Piccre.Height / Piccre.TextHeight("A")) + 1

'Next, we calculate a lot of time-eating stuff in advance.
'This is done before, to speedup timer sub ;-)

'set the fade-in and fade-out regions
CharHeight = Piccre.TextHeight("A")
If PercentFade <> 0 Then
    FadeOut = ((Piccre.Height / 100) * PercentFade) - CharHeight
    FadeIn = (Piccre.Height - FadeOut) - CharHeight - CharHeight
    Else
    FadeIn = Piccre.Height
    FadeOut = 0 - CharHeight
    End If
    
'set the percent values, ready for instant use later
ColText = Piccre.ForeColor
cDiff1 = (Piccre.Height - (CharHeight - 10)) - FadeIn
cDiff2 = 100 / cDiff1
cDiff3 = 100 / FadeOut

'calculate the left-position of each line, to center it
ReDim CreditLeft(TotalLines - 1)
For i = 0 To (TotalLines - 1)
    Select Case AlignText
    Case 1
        CreditLeft(i) = 100
    Case 2
        CreditLeft(i) = (Piccre.Width - Piccre.TextWidth(CreditLine(i))) / 2
    Case 3
        CreditLeft(i) = Piccre.Width - Piccre.TextWidth(CreditLine(i)) - 100
    End Select
Next i

'calculate 100 fade values from backcolor to forecolor
'(another time-eating thing done in advance)
Rcol1 = Piccre.ForeColor Mod 256
Gcol1 = (Piccre.ForeColor And vbGreen) / 256
Bcol1 = (Piccre.ForeColor And vbBlue) / 65536
Rcol2 = Piccre.BackColor Mod 256
Gcol2 = (Piccre.BackColor And vbGreen) / 256
Bcol2 = (Piccre.BackColor And vbBlue) / 65536
For i = 0 To 100
    Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
    Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
    Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
    ColorFades(i) = RGB(Rfade, Gfade, Bfade)
Next

'hit the throttle
Me.tmtext.Enabled = True
Exit Sub

End Sub
