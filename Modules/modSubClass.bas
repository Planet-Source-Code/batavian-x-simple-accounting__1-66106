Attribute VB_Name = "modSubClass"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_VSCROLL As Long = &H115

Private lngPicProc As Long

Public frmVSParent As Object

Public Sub StartPicDocSubClass(lnghWnd As Long)
   lngPicProc = SetWindowLong(lnghWnd, GWL_WNDPROC, AddressOf PicProc)
End Sub

Public Sub EndPicDocSubClass(lnghWnd As Long)
   SetWindowLong lnghWnd, GWL_WNDPROC, lngPicProc
End Sub

Private Function PicProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If Msg = WM_MOUSEWHEEL Then
      Select Case wParam
         Case &H780000
            If frmVSParent.vsPreview.Visible Then
               If frmVSParent.vsPreview.Value - 45 > frmVSParent.vsPreview.Min Then
                  frmVSParent.vsPreview.Value = frmVSParent.vsPreview.Value - 45
               ElseIf frmVSParent.vsPreview.Value > frmVSParent.vsPreview.Min Then
                  frmVSParent.vsPreview.Value = frmVSParent.vsPreview.Min
               End If
            End If
         Case &HFF880000
            If frmVSParent.vsPreview.Visible Then
               If frmVSParent.vsPreview.Value + 45 < frmVSParent.vsPreview.Max Then
                  frmVSParent.vsPreview.Value = frmVSParent.vsPreview.Value + 45
               ElseIf frmVSParent.vsPreview.Value < frmVSParent.vsPreview.Max Then
                  frmVSParent.vsPreview.Value = frmVSParent.vsPreview.Max
               End If
            End If
      End Select
   End If
         
   PicProc = CallWindowProc(lngPicProc, hWnd, Msg, wParam, lParam)
End Function

