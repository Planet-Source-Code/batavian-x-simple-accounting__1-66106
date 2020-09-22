Attribute VB_Name = "modCommon"
Option Explicit

Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Const OF_READWRITE As Long = &H2

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOSIZE As Long = &H1
Private Const WH_CBT As Long = 5
Private Const HCBT_ACTIVATE As Long = &H5

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

Private Const GWL_WNDPROC = (-4)

Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_PASTE As Long = &H302

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const LVS_EX_FULLROWSELECT As Long = &H20
Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_INFOTIP As Long = &H400

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVM_GETITEMRECT As Long = (LVM_FIRST + 14)
Private Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)

Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)

Private Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Private Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendXMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type EDITBALLOONTIP
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Private Const ECM_FIRST As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hWndCombo  As Long
   hWndEdit  As Long
   hWndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32.dll" (ByVal hWndCombo As Long, CBInfo As COMBOBOXINFO) As Long

Public MyTBDesc As TextBox
Public MyTBRem As TextBox
Public MyActiveTB As TextBox

Public MyLV_B_DD As lvButtons_H

Public MyCoolTip As New clsCoolTip
Public intCurrRecord As Integer

Private m_lngWindowProc As Long
Private lngMBHook As Long

Public Sub SetMenuForm(frmhWnd As Long, Optional LI As Control, Optional cT As clsTips, Optional MDIFormStyle As SubClassContainers = 0)
   modMenus.HighlightDisabledMenuItems = True
   modMenus.HighlightGradient = True
   modMenus.RaisedIconOnSelect = True
   modMenus.CheckMarksXPstyle = True
   modMenus.CheckedIconBackColor = vb3DHighlight
   
   GradCol1 = RGB(191, 191, 191)
   GradCol2 = RGB(255, 255, 255)
   GradBackCol = RGB(133, 133, 133)
   
   SelFrameCol = RGB(31, 31, 31)
   
   SelGradCol = RGB(191, 191, 191)
   SelGradBackCol = RGB(255, 255, 255)
   
   DisSelGradCol = RGB(223, 223, 223)
   DisSelGradBackCol = vbWhite
   
   SepCol = RGB(0, 0, 0)

   If IsMissing(LI) Then
      If IsMissing(cT) Then
         SetMenu frmhWnd, , , MDIFormStyle
      Else
         SetMenu frmhWnd, , cT, MDIFormStyle
      End If
   Else
      If IsMissing(cT) Then
         SetMenu frmhWnd, LI, , MDIFormStyle
      Else
         SetMenu frmhWnd, LI, cT, MDIFormStyle
      End If
   End If
End Sub

Public Function AppPath() As String
   Dim strPath As String
   
   strPath = App.Path
   
   If Right$(strPath, 1) <> "\" Then
      strPath = strPath & "\"
   End If
   
   AppPath = strPath
End Function

Public Sub SetLVExtendedStyle(LV As ListView, Optional bEnableHeaderDrag As Boolean = False)
   Dim lngLVStyle As Long
   
   lngLVStyle = SendMessage(LV.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
   
   lngLVStyle = lngLVStyle Or LVS_EX_FULLROWSELECT
   lngLVStyle = lngLVStyle Or LVS_EX_GRIDLINES
   lngLVStyle = lngLVStyle Or LVS_EX_INFOTIP
   
   If bEnableHeaderDrag Then
      lngLVStyle = lngLVStyle Or LVS_EX_HEADERDRAGDROP
   End If
   
   SendMessage LV.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, lngLVStyle
End Sub

Public Function GetLVColumnOrder(LV As ListView) As String
   Dim lngColArray() As Long
   Dim lngColCount As Long
   Dim intCount As Integer
   Dim strTemp As String
   
   lngColCount = LV.ColumnHeaders.Count
   ReDim lngColArray(lngColCount - 1)
   
   SendXMessage LV.hWnd, LVM_GETCOLUMNORDERARRAY, lngColCount, lngColArray(0)
   
   For intCount = 0 To lngColCount - 1
      strTemp = strTemp & lngColArray(intCount) & ";"
   Next
   
   GetLVColumnOrder = Left(strTemp, Len(strTemp) - 1)
End Function

Public Sub SetLVColumnOrder(LV As ListView, strOrder As String)
   Dim strColArray() As String
   Dim lngColArray() As Long
   Dim intCount As Integer

   strColArray = Split(strOrder, ";")

   ReDim lngColArray(UBound(strColArray))

   For intCount = 0 To UBound(strColArray)
      lngColArray(intCount) = Val(strColArray(intCount))
   Next intCount
   
   SendXMessage LV.hWnd, LVM_SETCOLUMNORDERARRAY, LV.ColumnHeaders.Count, lngColArray(0)
End Sub

Public Sub ShowBalloonTip(hWnd As Long, sTitle As String, sTip As String, Optional eIcon As IconMode = 1)
   Dim tEBT As EDITBALLOONTIP
   Dim lhWnd As Long
   
   With tEBT
      .cbStruct = Len(tEBT)
      .pszTitle = StrConv(sTitle, vbUnicode)
      .pszText = StrConv(sTip, vbUnicode)
      .ttiIcon = CLng(eIcon)
   End With

   SendXMessage hWnd, EM_SHOWBALLOONTIP, 0, tEBT
End Sub

Public Sub SetReadonly(hWnd As Long)
   m_lngWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub SetRewrite(hWnd As Long)
   SetWindowLong hWnd, GWL_WNDPROC, m_lngWindowProc
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case Msg
      Case WM_CONTEXTMENU
         WndProc = 0
         
      Case WM_PASTE
         ShowBalloonTip hWnd, "Invalid Input Methode", "This field can only be filled by chosing value in the list!", [4-Critical Icon]
         WndProc = 0
         
      Case Else: WndProc = CallWindowProc(m_lngWindowProc, hWnd, Msg, wParam, lParam)
   End Select
End Function

Public Sub SortListByColumn(LVCtrl As ListView, LVColumn As ComctlLib.ColumnHeader)
   LVCtrl.SortKey = LVColumn.Index - 1
   
   Select Case LVColumn.Tag
    Case "Asc"
      LVCtrl.SortOrder = lvwDescending
      LVColumn.Tag = "Dsc"
    Case "Dsc"
      LVCtrl.SortOrder = lvwAscending
      LVColumn.Tag = "Asc"
    Case Else
      LVCtrl.SortOrder = lvwDescending
      LVColumn.Tag = "Dsc"
   End Select
End Sub

Public Function xMsgBox(lngParent As Long, ByVal strMessage As String, Optional ByVal MsgBoxStyle As VbMsgBoxStyle = vbOKOnly, Optional ByVal strTitle As String = "") As VbMsgBoxResult
   lngMBHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, App.hInstance, GetCurrentThreadId())
   xMsgBox = MessageBox(lngParent, strMessage, strTitle, MsgBoxStyle)
End Function

Private Function MsgBoxHookProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim ParenthWnd As Long
   Dim tRect As RECT
   Dim ParentRect As RECT

   MsgBoxHookProc = CallNextHookEx(lngMBHook, lngCode, wParam, lParam)
   
   If lngCode = HCBT_ACTIVATE Then
      SetMenuForm wParam
      ParenthWnd = GetParent(wParam)
      GetWindowRect ParenthWnd, ParentRect
      GetWindowRect wParam, tRect
      
      If ParenthWnd = 0 Then
         ParentRect.Left = 0
         ParentRect.Top = 0
         ParentRect.Right = Screen.Width / Screen.TwipsPerPixelX
         ParentRect.Bottom = Screen.Height / Screen.TwipsPerPixelY
      End If

      SetWindowPos wParam, 0, (((ParentRect.Right - ParentRect.Left) - (tRect.Right - tRect.Left)) / 2) + ParentRect.Left, _
                              (((ParentRect.Bottom - ParentRect.Top) - (tRect.Bottom - tRect.Top)) / 2) + ParentRect.Top, _
                                 tRect.Right - tRect.Left, tRect.Bottom - tRect.Top, SWP_NOSIZE Or SWP_NOZORDER
      UnhookWindowsHookEx lngMBHook
   End If
End Function

Public Function GetComboStruckHandle(CB As ComboBox, Optional EditHandle As Boolean = True) As Long
   Dim CBI As COMBOBOXINFO

   CBI.cbSize = Len(CBI)
   Call GetComboBoxInfo(CB.hWnd, CBI)
   
   If EditHandle Then
      GetComboStruckHandle = CBI.hWndEdit
   Else
      GetComboStruckHandle = CBI.hWndList
   End If
End Function

Public Sub AddReportLogo(picLogo As PictureBox, strTitle As String, qPrinter As clsPrinter, Optional qOrientation As qePrintOrientation = qLandscape, Optional intPage As Integer = 1, Optional intPaperSize As qePrinterPaperSize = eA4size, Optional sngLeft As Single = 32, Optional sngTop As Single = 20, Optional sngWidth As Single = 810, Optional sngHeight As Single = 810, Optional bFixed As Boolean = True)
   With qPrinter
      .AppName = App.Title
      .PageSize = intPaperSize
      .Orientation = qOrientation
      .ScaleMode = eTwip
      .MarginBottom = 300
      .MarginTop = 600
      .MarginLeft = 300
      .MarginRight = 300
      
      Set .PicBox = frmMain.picLogo
      
      .AddPic picLogo, intPage, sngLeft, sngTop, sngWidth, sngHeight
      
      picLogo.FontName = "Arial"
      picLogo.FontBold = True
      picLogo.FontSize = 12
      picLogo.ScaleMode = vbTwips
      
      .AddText "<LINDENT=1150>PT MAJUKO UTAMA INDONESIA<SIZE=12><LINDENT=" & IIf(qOrientation = qLandscape, 14850, 10650) - picLogo.TextWidth(strTitle) & ">" & strTitle, "Arial", 14, True
      
      If bFixed Then .TextItem(.TextItem.Count).Top = -90 '210
   End With
End Sub

Public Sub DrawPicHolderCaption(PicBox As PictureBox)
   Dim intY As Integer
   Dim intRGB As Integer

   intRGB = 127

   For intY = 0 To 14
      PicBox.Line (0, intY)-(PicBox.ScaleWidth, intY), RGB(intRGB, intRGB, intRGB)
      intRGB = intRGB + 3
   Next

   PicBox.ForeColor = vbBlack
   PicBox.CurrentX = 11
   PicBox.CurrentY = 2
   PicBox.FontBold = True
   PicBox.Print PicBox.Tag
   PicBox.ForeColor = vbWhite
   PicBox.CurrentX = 10
   PicBox.CurrentY = 1
   PicBox.Print PicBox.Tag
End Sub

Public Sub ChangeFileTime(strFileName As String)
   Dim MyFT As FILETIME
   Dim MyST As SYSTEMTIME
   Dim lngF As Long
      
   MyST.wDay = Day(Date)
   MyST.wMonth = Month(Date)
   MyST.wYear = Year(Date)
   MyST.wHour = Hour(Time)
   MyST.wMinute = Minute(Time)
   MyST.wSecond = Second(Time)
   MyST.wMilliseconds = 0
   
   SystemTimeToFileTime MyST, MyFT
   LocalFileTimeToFileTime MyFT, MyFT
   
   lngF = lOpen(strFileName, OF_READWRITE)
      SetFileTime lngF, MyFT, MyFT, MyFT
   lClose lngF
End Sub

Public Function GetOrdinal(intOrdinal As Integer) As String
   Dim strNumber As String
   
   strNumber = Right(CStr(intOrdinal), 1)
   GetOrdinal = "th"
   
   Select Case strNumber
      Case 1: GetOrdinal = "st"
      Case 2: GetOrdinal = "nd"
      Case 3: GetOrdinal = "rd"
   End Select
End Function

Public Function GetMonthName(intmonth As Integer) As String
   Dim strM(1 To 12) As String
   
   strM(1) = "January"
   strM(2) = "February"
   strM(3) = "March"
   strM(4) = "April"
   strM(5) = "May"
   strM(6) = "June"
   strM(7) = "July"
   strM(8) = "August"
   strM(9) = "September"
   strM(10) = "October"
   strM(11) = "November"
   strM(12) = "December"
   
   If intmonth < 1 Or intmonth > 12 Then
      GetMonthName = vbNullString
   Else
      GetMonthName = strM(intmonth)
   End If
End Function
