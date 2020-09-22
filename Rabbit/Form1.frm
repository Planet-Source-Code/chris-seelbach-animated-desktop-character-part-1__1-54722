VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Drag Me"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   960
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   2280
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   510
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   900
      _Version        =   393216
      BackColor       =   16777215
      FullWidth       =   42
      FullHeight      =   34
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2295
         Left            =   -120
         TabIndex        =   1
         Top             =   -120
         Width           =   2775
         ExtentX         =   4895
         ExtentY         =   4048
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************
'Animated Desktop Character Part #1
'
'1) This window will be on top of other windows.
'
'2) You can drag the rabbit around the screen,
'   (provided that you can catch him)
'
'3) To end the program; click on the rabbit, then
'   press ESC. (lol)
'
'   Cheers to you.  ;)
'**************************************************
'wave sound stuff
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

'for our animated gif
Public PicPath As String
'for our window
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_NOACTIVATE    As Long = 16
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_COMBINED      As Long = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Private Const SWP_TOPMOST       As Long = -1
'use as our class
Private c As clsTransp
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Form's KeyPreview property is set to True,
'if you press the ESC button...
'Note: you must click on the rabbit to return the
'focus to the form, so this escape event will be
'recognised. You can use whichever key you want, or
'come up with your own escape code.
    If KeyCode = 27 Then
        Set Form1 = Nothing
    End
    End If
End Sub
Private Sub Form_Load()
    Dim clr&
'set our transparent color:
'1) the form's backcolor
'2) the default backcolor of the Webbrowser (when active)
'3) the backcolor of our PictureBox
    clr = vbWhite
'*********************************************************
'we will only do this once, initially upon the Form's load
   Set c = New clsTransp
   
      c.TransparentColor = clr
      c.SetTransparency hWnd, TranparentByColor
      Set c = Nothing
'*********************************************************
DoEvents
    'set our window on top
    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED
    'define the location of our html file
    PicPath = App.Path & "\rabbit.html"
    'now open this file
    WebBrowser1.Navigate PicPath
'************************************************************************
'A picturebox is used on the form to "cut-off" the Webbrowser's
'scroll bars. It's archaic, but it works. The following line will
'do the same thing, without the need for a picturebox, but I don't
'know if it works in all OS's.
'WebBrowser1.Navigate "about:<html><body scroll='no'><BODY TOPMARGIN='0' LEFTMARGIN='0' MARGINWIDTH='0' MARGINHEIGHT='0'><img src='LOCATION OF GIF'></img></body></html>" '</p>
'************************************************************************
    'start our movement timer
    Timer1.Enabled = True
    'start our sound timer
    Timer2.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    End
End Sub

Private Sub Timer1_Timer()
    'move the rabbit across the screen from right-to-left; cycle this movement
    Form1.Move Form1.Left - 100
        'when the form moves off the screen, put it
        'back somewhere on the right side
        If Form1.Left < -2000 Then
        Randomize
            Form1.Move Form1.Left + Screen.Width + Picture1.Width, Int(Rnd * Screen.Height)
        Else
        End If
End Sub

Private Sub Timer2_Timer()
'play the sound
    SoundName = "whatsup.wav"
   wFlags = SND_ASYNC Or SND_NODEFAULT
   p% = sndPlaySound(SoundName, wFlags)
'reset our timer to play the sound
'at random intervals
    Timer2.Interval = Int(Rnd * 15000) + 5000
End Sub

Private Sub WebBrowser1_GotFocus()
'when you click the WebBrowser (the rabbit), and/or hold
'the mouse button down, this control now has the focus.
Set c = New clsTransp
    'drag the rabbit
    c.GrabForm Me
    DoEvents
    'set focus to another object when you release the mouse
    'button. Without this event, the WebBrowser will not respond
    'to subsequent "clicks". You can use any object on the Form
    'that you want, I chose the Animation Control for Part #2 of
    'this submission.  ;)
    Animation1.SetFocus
    
End Sub









