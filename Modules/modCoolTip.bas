Attribute VB_Name = "modCoolTip"
Option Explicit

Private defWindowTipProc As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_TIMER As Long = &H113
Private Const WMW_TIMER_EXEC As Long = &H4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub SubClassTip(hWnd As Long)
   defWindowTipProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowTipProc)
End Sub

Public Sub UnSubClassTip(hWnd As Long)
   If defWindowTipProc <> 0 Then
      SetWindowLong hWnd, GWL_WNDPROC, defWindowTipProc
      defWindowTipProc = 0
   End If
End Sub

Private Function WindowTipProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If uMsg = WM_TIMER And wParam = WMW_TIMER_EXEC Then
      WindowTipProc = 0
   Else
      WindowTipProc = CallWindowProc(defWindowTipProc, hWnd, uMsg, wParam, lParam)
   End If
End Function
