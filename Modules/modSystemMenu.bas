Attribute VB_Name = "modSystemMenu"
Option Explicit

Public Enum IDM
    a = 128
End Enum
Public procOld As Long
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
Private Const WM_SYSCOMMAND = &H112

Public Function WindowProc(ByVal hwnd As Long, _
ByVal iMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

   Select Case iMsg
      Case WM_SYSCOMMAND
         Select Case wParam
         Case IDM.a
            frmAbout.Show 1
         End Select
   End Select

    WindowProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)
End Function



