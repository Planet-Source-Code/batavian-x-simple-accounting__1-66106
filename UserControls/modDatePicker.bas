Attribute VB_Name = "modDatePicker"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Const CB_GETITEMHEIGHT = &H154
         
Public Declare Function HideCaret Lib "user32.dll" (ByVal hWnd As Long) As Long

Public myRect As RECT
Public DTValue As Date
Public sMonth(1 To 12) As String
Public m_DaysCount(1 To 12) As Integer
Public lMinYear As Long
Public lMaxYear As Long
Public CBhWnd As Long
Public isDropdowned As Boolean
Public isShowToday As Boolean

Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CBN_DROPDOWN As Long = 7
Private Const GWL_WNDPROC = (-4)

Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_PASTE As Long = &H302

Private lCBWinProc As Long
Private lTBWinProc As Long

Private mvarDayEditBox As TextBox
Private mvarMonthEditBox As TextBox
Private mvarYearEditBox As TextBox

Public Property Get DayEditBox() As TextBox
    Set DayEditBox = mvarDayEditBox
End Property

Public Property Set DayEditBox(ByVal vNewValue As TextBox)
    Set mvarDayEditBox = vNewValue
End Property

Public Sub SelectText(ctlTB As TextBox)
   ctlTB.SelStart = 0
   ctlTB.SelLength = Len(ctlTB.Text)
End Sub

Public Property Get MonthEditBox() As TextBox
    Set MonthEditBox = mvarMonthEditBox
End Property

Public Property Set MonthEditBox(ByVal vNewValue As TextBox)
    Set mvarMonthEditBox = vNewValue
End Property

Public Property Get YearEditBox() As TextBox
    Set YearEditBox = mvarYearEditBox
End Property

Public Property Set YearEditBox(ByVal vNewValue As TextBox)
    Set mvarYearEditBox = vNewValue
End Property

Public Function IsLeapYear(iYear As Integer) As Boolean
   IsLeapYear = iYear Mod 4 = 0
   If iYear Mod 100 = 0 Then IsLeapYear = iYear Mod 400 = 0
End Function

Public Sub SetCBDropDownHeigth(ParenthWnd As Long, CB As ComboBox, Items2Display As Integer) ', Optional bResizeList As Boolean = False)
   Dim pt As POINTAPI
   Dim rc As RECT
   Dim lNewHeight As Long
   Dim lItemHeight As Long
   
   lItemHeight = SendMessage(CB.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   lNewHeight = lItemHeight * (Items2Display + 2)
   
   GetWindowRect CB.hWnd, rc
   
   pt.x = rc.Left
   pt.y = rc.Top
   
   ScreenToClient ParenthWnd, pt
   GetWindowRect CB.hWnd, rc
   
   MoveWindow CB.hWnd, pt.x, pt.y, rc.Right - rc.Left, lNewHeight, True
End Sub

Public Sub ReDimDaysArray(iYear As Integer)
   Dim i As Integer
   
   For i = 1 To 12
      If i <> 4 And i <> 6 And i <> 9 And i <> 11 Then
         m_DaysCount(i) = 31
      Else
         m_DaysCount(i) = 30
      End If
   Next i
      
   If IsLeapYear(iYear) Then
      m_DaysCount(2) = 29
   Else
      m_DaysCount(2) = 28
   End If
End Sub

Public Function MaxDayOfMonth(iYear As Integer, iMonth As Integer) As Integer
   If iMonth <> 4 And iMonth <> 6 And iMonth <> 9 And iMonth <> 11 Then
      MaxDayOfMonth = 31
   Else
      MaxDayOfMonth = 30
   End If
   
   If iMonth = 2 Then
      If IsLeapYear(iYear) Then
         MaxDayOfMonth = 29
      Else
         MaxDayOfMonth = 28
      End If
   End If
End Function

Public Sub HookContext(hWnd As Long)
   lTBWinProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub UnhookContext(hWnd As Long)
   SetWindowLong hWnd, GWL_WNDPROC, lTBWinProc
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case Msg
      Case WM_CONTEXTMENU, WM_PASTE
         WndProc = 0

      Case Else
         WndProc = CallWindowProc(lTBWinProc, hWnd, Msg, wParam, lParam)

   End Select
End Function
