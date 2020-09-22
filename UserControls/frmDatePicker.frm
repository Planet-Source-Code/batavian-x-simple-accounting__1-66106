VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmDatePicker 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2775
      Top             =   2490
   End
   Begin VB.PictureBox pic4Day 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   180
      ScaleHeight     =   270
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   870
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   0
      Left            =   180
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   600
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picToday 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      ScaleHeight     =   255
      ScaleWidth      =   420
      TabIndex        =   12
      Top             =   2520
      Width           =   420
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmDatePicker.frx":0000
      Left            =   150
      List            =   "frmDatePicker.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   1020
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   1440
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   1860
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   5
      Left            =   2280
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin VB.PictureBox picDay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   2700
      ScaleHeight     =   300
      ScaleWidth      =   435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   585
      Width           =   435
   End
   Begin ComCtl2.UpDown UDYear 
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   165
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   503
      _Version        =   327681
      BuddyControl    =   "txtYear"
      BuddyDispid     =   196615
      OrigLeft        =   3150
      OrigTop         =   1800
      OrigRight       =   3420
      OrigBottom      =   2550
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtContainer 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   1890
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   555
      Width           =   3015
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   2085
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "2999"
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblToday 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      Height          =   195
      Left            =   630
      TabIndex        =   13
      Top             =   2580
      Width           =   450
   End
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SPI_GETWORKAREA As Long = 48

Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private iDayStart As Integer
Private iCurrDay As Integer

Private isLoading As Boolean
Private isStepApply As Boolean

Private Sub cboMonth_Click()
   If isLoading Then Exit Sub
   ApplyDays
   UpdateDate
End Sub

Private Sub cboMonth_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   Dim lRet As Long
   Dim rct As RECT
   Dim hRgn As Long
   Dim lWS As Long
   Dim i As Integer
   Dim sDay(6, 1) As String
   Dim lLeft As Long
   Dim lTop As Long
   Dim hBrush As Long
   
   hRgn = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 3, 3)
   hBrush = CreateSolidBrush(GetSysColor(vbActiveTitleBar And &HFF&))
   FrameRgn Me.hDC, hRgn, hBrush, 1, 1
   
   DeleteObject hRgn
   DeleteObject hBrush
   
   sDay(0, 0) = "Sun"
   sDay(1, 0) = "Mon"
   sDay(2, 0) = "Tue"
   sDay(3, 0) = "Wed"
   sDay(4, 0) = "Thu"
   sDay(5, 0) = "Fri"
   sDay(6, 0) = "Sat"
   
   sDay(0, 1) = "Sunday"
   sDay(1, 1) = "Monday"
   sDay(2, 1) = "Tuesday"
   sDay(3, 1) = "Wednesday"
   sDay(4, 1) = "Thursday"
   sDay(5, 1) = "Friday"
   sDay(6, 1) = "Saturday"
   
   isLoading = True
   
   SystemParametersInfo SPI_GETWORKAREA, 0&, rct, 0&
   
   rct.Left = (rct.Left) * Screen.TwipsPerPixelX
   rct.Right = (rct.Right) * Screen.TwipsPerPixelX
   rct.Top = (rct.Top) * Screen.TwipsPerPixelY
   rct.Bottom = (rct.Bottom) * Screen.TwipsPerPixelY
   
   With rct
      lLeft = (myRect.Left - 1) * Screen.TwipsPerPixelX
      lTop = (myRect.Bottom + 1) * Screen.TwipsPerPixelY
      
      If lLeft + Me.Width > .Right Then lLeft = .Right - Me.Width
      If lLeft < .Left Then lLeft = .Left
      If lTop + Me.Height > rct.Bottom Then lTop = (myRect.Top * Screen.TwipsPerPixelY) - Me.Height
      If lTop < .Top Then lTop = .Top
   End With
   
   Me.Left = lLeft
   Me.Top = lTop
   
   For i = 2 To 42
      Load pic4Day(i)
      pic4Day(i).Visible = True
      pic4Day(i).TabIndex = i + 1
      
      If i > 7 Then
         pic4Day(i).Left = pic4Day(i - 7).Left
         pic4Day(i).Top = pic4Day(i - 7).Top + pic4Day(i - 7).Height - 15
      Else
         pic4Day(i).Left = picDay(i - 1).Left
         pic4Day(i).Top = pic4Day(i - 1).Top
      End If
      
      pic4Day(i).ZOrder
   Next i
   
   For i = 1 To 12
      cboMonth.AddItem sMonth(i)
   Next i
   
   For i = 0 To 6
      picDay(i).CurrentX = (picDay(i).Width - picDay(i).TextWidth(sDay(i, 0))) / 2
      picDay(i).CurrentY = (picDay(i).Height - picDay(i).TextHeight(sDay(i, 0))) / 2
      picDay(i).Print sDay(i, 0)
      picDay(i).Line (0, 0)-(picDay(i).ScaleWidth - 0, 0), vbActiveTitleBar
      DrawBorder picDay(i)
   Next i
   
   UDYear.Max = lMaxYear
   UDYear.Min = lMinYear
   
   txtYear = YearEditBox.Text
   cboMonth.Text = MonthEditBox.Text
   iCurrDay = DayEditBox.Text
   
   lblToday = lblToday & ": " & sDay(Weekday(Date, vbSunday) - 1, 1) & ", " & sMonth(Month(Date)) & " " & Day(Date) & " " & Year(Date)
   ReDimDaysArray txtYear
   ApplyDays
   
   isLoading = False
   
   hRgn = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 3, 3)
   SetWindowRgn Me.hWnd, hRgn, True
   DeleteObject hRgn
   
   DrawToday picToday, False
   HookContext txtYear.hWnd
   SetCBDropDownHeigth Me.hWnd, cboMonth, 13
   UDYear.Left = txtYear.Left + txtYear.Width - UDYear.Width - 15
   
   If Not isShowToday Then
      Me.Height = 2610
      picToday.Visible = False
      lblToday.Visible = False
   End If
End Sub

Private Sub Form_LostFocus()
   Unload Me
End Sub

Private Sub Form_Paint()
'   Dim hRgn As Long
'   Dim hBrush As Long
'
'   hRgn = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 3, 3)
'   hBrush = CreateSolidBrush(GetSysColor(vbActiveTitleBar And &HFF&))
'   FrameRgn Me.hDC, hRgn, hBrush, 1, 1
'
'   DeleteObject hRgn
'   DeleteObject hBrush
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MyLV_B_DD.Value = False
   MyLV_B_DD.Enabled = True

   Set DayEditBox = Nothing
   Set MonthEditBox = Nothing
   Set YearEditBox = Nothing
   
   Set MyLV_B_DD = Nothing
   
   UnhookContext txtYear.hWnd
   isDropdowned = False
End Sub

Private Sub lblToday_Click()
   picToday_MouseUp 1, 0, 0, 0
End Sub

Private Sub lblToday_DblClick()
   Unload Me
End Sub

Private Sub pic4Day_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub pic4Day_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      Dim iPDay As Integer
      
      iPDay = m_DaysCount(IIf(cboMonth.ListIndex = 0, 12, cboMonth.ListIndex))
      iPDay = iPDay - iDayStart
      
      If Index < iDayStart + 1 Then
         iCurrDay = Index + iPDay
         
         If cboMonth.ListIndex = 0 Then
            cboMonth.ListIndex = 11
            txtYear = IIf(txtYear <= lMinYear, lMaxYear, txtYear - 1)
         Else
            cboMonth.ListIndex = cboMonth.ListIndex - 1
         End If
      ElseIf Index >= iDayStart + 1 And Index <= m_DaysCount(cboMonth.ListIndex + 1) + iDayStart Then
         iCurrDay = Index - iDayStart
      ElseIf Index > m_DaysCount(cboMonth.ListIndex + 1) + iDayStart Then
         iCurrDay = Index - m_DaysCount(cboMonth.ListIndex + 1) - iDayStart
         
         If cboMonth.ListIndex = 11 Then
            cboMonth.ListIndex = 0
            txtYear = IIf(txtYear >= lMaxYear, lMinYear, txtYear + 1)
         Else
            cboMonth.ListIndex = cboMonth.ListIndex + 1
         End If
      End If
      
      ApplyDays
      UpdateDate
   End If
End Sub

Private Sub pic4Day_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
   Static ctlMov As Boolean

   If (x < 0) Or (x > pic4Day(Index).ScaleWidth) Or (y < 0) Or (y > pic4Day(Index).ScaleHeight) Then
      If Button <> vbLeftButton Then
         ReleaseCapture
         ctlMov = False
         
         If pic4Day(Index).FontBold = True Then
            pic4Day(Index).FontBold = False
            pic4Day(Index).FontSize = 8
            DrawDays pic4Day(Index), Index
         End If
         
         If cboMonth.ListIndex + 1 = Month(Date) And Year(Date) = txtYear And isShowToday Then
            DrawToday pic4Day(Day(Date) + iDayStart)
         End If
      End If
   Else
      If Button <> vbLeftButton Then
         SetCapture pic4Day(Index).hWnd
         
         If ctlMov = False Then ctlMov = True
         pic4Day(Index).FontBold = True
         pic4Day(Index).FontSize = 10
         DrawDays pic4Day(Index), Index
         
         If cboMonth.ListIndex + 1 = Month(Date) And Year(Date) = txtYear And isShowToday Then
            DrawToday pic4Day(Day(Date) + iDayStart)
         End If
      End If
   End If
End Sub

Private Sub pic4Day_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim pAPI As POINTAPI
   
   If Button = vbLeftButton Then
      If Index = iCurrDay + iDayStart Then
         GetCursorPos pAPI
         
         If WindowFromPoint(pAPI.x, pAPI.y) = pic4Day(Index).hWnd Then
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub ApplyDays()
   Dim sDay As String
   Dim i As Integer
   
   i = Weekday("1/" & cboMonth.ListIndex + 1 & "/" & txtYear, vbSunday)
   sDay = WeekdayName(i, , vbSunday)
   i = 0
   
   Do
      i = i + 1
      
      If sDay = WeekdayName(i, , vbSunday) Then
         iDayStart = i
         Exit Do
      End If
   Loop
   
   For i = 1 To 42
      pic4Day(i).BackColor = vbWindowBackground
      pic4Day(i).ForeColor = vbWindowText
      pic4Day(i).FontBold = False
      pic4Day(i).FontSize = 8
   Next i
   
   For i = 1 To 42 Step 7
      pic4Day(i).ForeColor = &HC0&
   Next i
   
   For i = 6 To 42 Step 7
      pic4Day(i).ForeColor = &H8000&
   Next i
   
   iDayStart = iDayStart - 1
   
   If iCurrDay > m_DaysCount(cboMonth.ListIndex + 1) Then
      iCurrDay = m_DaysCount(cboMonth.ListIndex + 1)
   End If
   
   pic4Day(iCurrDay + iDayStart).BackColor = vbHighlight
   pic4Day(iCurrDay + iDayStart).ForeColor = vbHighlightText
   
   For i = 1 To 42
      DrawDays pic4Day(i), i
   Next i
   
   If cboMonth.ListIndex + 1 = Month(Date) And Year(Date) = txtYear And isShowToday Then
      DrawToday pic4Day(Day(Date) + iDayStart)
   End If
End Sub

Private Sub DrawDays(pD2Draw As PictureBox, ByVal Index As Integer)
   Dim iPDay As Integer
   
   iPDay = m_DaysCount(IIf(cboMonth.ListIndex = 0, 12, cboMonth.ListIndex))
   iPDay = iPDay - iDayStart
   
   pD2Draw.Cls
   If Index > 35 Then pD2Draw.Line (pD2Draw.ScaleWidth, pD2Draw.ScaleHeight - 15)-(0, pD2Draw.ScaleHeight - 15), vbActiveTitleBar
      
   If Index < iDayStart + 1 Then
      Index = Index + iPDay
      pD2Draw.ForeColor = RGB(191, 191, 191)
   ElseIf Index >= iDayStart + 1 And Index <= m_DaysCount(cboMonth.ListIndex + 1) + iDayStart Then
      Index = Index - iDayStart
   ElseIf Index > m_DaysCount(cboMonth.ListIndex + 1) + iDayStart Then
      Index = Index - m_DaysCount(cboMonth.ListIndex + 1) - iDayStart
      pD2Draw.ForeColor = RGB(191, 191, 191)
   End If
   
   pD2Draw.CurrentX = (pD2Draw.Width - pD2Draw.TextWidth(Index)) - 120  '/ 2
   pD2Draw.CurrentY = (pD2Draw.Height - pD2Draw.TextHeight(Index)) / 2
   pD2Draw.Print Index
   
   DrawBorder pD2Draw
End Sub

Private Sub DrawBorder(pBorder As PictureBox)
   pBorder.Line (0, 0)-(0, pBorder.ScaleHeight - 0), vbActiveTitleBar
   pBorder.Line (pBorder.ScaleWidth - 15, pBorder.ScaleHeight)-(pBorder.ScaleWidth - 15, -15), vbActiveTitleBar
   
'=== should be optional ===
'   pBorder.Line (0, 0)-(pBorder.ScaleWidth - 0, 0), vbActiveTitleBar
'   pBorder.Line (pBorder.ScaleWidth, pBorder.ScaleHeight - 15)-(0, pBorder.ScaleHeight - 15), vbActiveTitleBar
End Sub

Private Sub DrawToday(pToday As PictureBox, Optional DrawNumber As Boolean = True)
'   pToday.DrawWidth = 2
'   pToday.Line (-15, 45)-(465, 240), vbRed, B '(-60, -60)-(pToday.Width + 45, pToday.Height + 45), , B
'   pToday.DrawWidth = 1
   pToday.Line (30, 0)-(pToday.ScaleWidth - 45, 60), vbRed
   pToday.Line (pToday.ScaleWidth - 45, 60)-(pToday.ScaleWidth - 45, pToday.ScaleHeight - 30), vbRed
   pToday.Line (30, pToday.ScaleHeight - 30)-(pToday.ScaleWidth - 45, pToday.ScaleHeight - 30), vbRed
   pToday.Line (30, 60)-(30, pToday.ScaleHeight - 30), vbRed
   pToday.Line (30, 60)-(pToday.ScaleWidth - 45, 0), vbRed
   
   If DrawNumber Then
      pToday.CurrentX = (pToday.Width - pToday.TextWidth(Day(Date))) - 120
      pToday.CurrentY = (pToday.Height - pToday.TextHeight(Day(Date))) / 2
      pToday.Print Day(Date)
      DrawBorder pToday
   End If
End Sub

Private Sub picToday_DblClick()
   Unload Me
End Sub

Private Sub picToday_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
   Static ctlMov As Boolean

   If (x < 0) Or (x > picToday.ScaleWidth) Or (y < 0) Or (y > picToday.ScaleHeight) Then
      ReleaseCapture
      ctlMov = False
      picToday.BackColor = &HC0C0C0
   Else
      SetCapture picToday.hWnd
      If ctlMov = False Then ctlMov = True
      picToday.BackColor = vbActiveTitleBar
   End If
   
   DrawToday picToday, False
End Sub

Private Sub picToday_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   cboMonth.ListIndex = Month(Date) - 1
   txtYear = Year(Date)
   iCurrDay = Day(Date)
   ApplyDays
   UpdateDate
End Sub

Private Sub Timer1_Timer()
   If GetForegroundWindow <> Me.hWnd Then
      Timer1.Enabled = False
      Unload Me
   End If
End Sub

Private Sub UpdateDate()
   modDatePicker.DayEditBox.Text = iCurrDay
   modDatePicker.MonthEditBox.Text = cboMonth.List(cboMonth.ListIndex)
   modDatePicker.YearEditBox.Text = txtYear
   
   DTValue = DateSerial(txtYear, cboMonth.ListIndex + 1, iCurrDay)
End Sub

Private Sub txtYear_Change()
   If isLoading Then Exit Sub
   If txtYear > lMaxYear Then txtYear = lMaxYear
   If txtYear < lMinYear Then txtYear = lMinYear
   
   ReDimDaysArray txtYear
   ApplyDays
   UpdateDate
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub

Private Sub txtYear_Click()
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub

Private Sub txtYear_DblClick()
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub

Private Sub txtYear_GotFocus()
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
   
   Select Case KeyCode
      Case vbKeyUp, vbKeyAdd, vbKeyRight
         If txtYear = lMaxYear Then
            txtYear = lMinYear
         Else
            txtYear = Val(txtYear) + 1
         End If
      Case vbKeyDown, vbKeySubtract, vbKeyLeft
         If txtYear = lMinYear Then
            txtYear = lMaxYear
         Else
            txtYear = txtYear - 1
         End If
      Case vbKeyPageUp
         If txtYear >= lMaxYear Then
            If Not txtYear = lMaxYear Then txtYear = lMaxYear
         Else
            txtYear = Val(txtYear) + 10
         End If
      Case vbKeyPageDown
         If txtYear < lMinYear Then
            If Not txtYear = lMinYear Then txtYear = lMinYear
         Else
            txtYear = Val(txtYear) - 10
         End If
   End Select
   
   SelectText txtYear
   KeyCode = 0
End Sub

Private Sub txtYear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub
