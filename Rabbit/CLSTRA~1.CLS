VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'code that produces a transparent form, and allows
'movement of that form; capturing opaque regions
'*************************************************************************
'Credit is also due; Evan Toder & Ulli. Thank you from me.  ;)
'*************************************************************************
Public Enum TT
      TranparentByColor = 0
      TransparentByPercent = 1
End Enum

Private m_TransparentColor&
 
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const HTCAPTION         As Long = 2
Public Sub GrabForm(Frm As Form)
'send the message to grab the Form (you are draging it)
    ReleaseCapture
    SendMessage Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'grab form

End Sub
Public Sub SetTransparency(lHwnd As Long, TranspType As TT, Optional TranspNum& = 255)
Attribute SetTransparency.VB_Description = "If setting transparency by color, you must enter a value for RGB_Red, RGB_Green, and RGB_Blue.   If not setting transparency by color then set the transparencypercent to a number from 0-100"

    Dim Ret&

    'Set the window style to 'Layered'
    Ret = GetWindowLong(lHwnd, GWL_EXSTYLE)
    Ret = (Ret Or WS_EX_LAYERED)
    SetWindowLong lHwnd, GWL_EXSTYLE, Ret
    
    If TranspType = TransparentByPercent Then
        SetLayeredWindowAttributes lHwnd, _
               0, 128, LWA_ALPHA
    Else
        SetLayeredWindowAttributes lHwnd, _
             m_TransparentColor&, 0&, LWA_COLORKEY
    End If
    
    'now refresh
    SetWindowPos lHwnd, _
              0, 0, 0, 0, 0, _
           (&H1 Or &H2 Or &H4 Or &H20)
           
End Sub
Public Property Get TransparentColor() As Long
Attribute TransparentColor.VB_Description = "number from 0-255"
   TransparentColor = m_TransparentColor
End Property
Public Property Let TransparentColor(ByVal vNewValue As Long)
    m_TransparentColor = vNewValue
End Property
