Attribute VB_Name = "modDwmAPI"

Public Type MARGINS
 'public int Left;
m_Left As Long
  'public int Right;
m_Right As Long
  'public int Top;
m_Top As Long
 'public int Bottom;
m_Button As Long


End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Dim Inied As Boolean
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern void DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS margins);
Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, margin As MARGINS) As Long
'[DllImport("dwmapi.dll", PreserveSig=false)]
'static extern bool DwmIsCompositionEnabled();
Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetLayeredWindowAttributesByColor Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


