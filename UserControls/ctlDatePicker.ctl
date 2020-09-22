VERSION 5.00
Begin VB.UserControl ctlDatePicker 
   BackColor       =   &H000000C0&
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   PropertyPages   =   "ctlDatePicker.ctx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   2715
   ToolboxBitmap   =   "ctlDatePicker.ctx":000F
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1710
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   30
      Width           =   255
      Begin Project1.lvButtons_H lvB_DropDown 
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         Image           =   "ctlDatePicker.ctx":0109
         ImgSize         =   40
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1245
      Locked          =   -1  'True
      MaxLength       =   4
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Text            =   "1900"
      Top             =   45
      Width           =   420
   End
   Begin VB.TextBox txtMonth 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   390
      Locked          =   -1  'True
      MaxLength       =   9
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Text            =   "September"
      Top             =   45
      Width           =   810
   End
   Begin VB.TextBox txtDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   75
      Locked          =   -1  'True
      MaxLength       =   2
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "01"
      Top             =   45
      Width           =   240
   End
   Begin VB.TextBox txtContainer 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   1
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "ctlDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

'Default Property Values:
Const m_def_DTPDay = 1
Const m_def_DTPMonth = 1
Const m_def_DTPYear = 1900
Const m_def_ShowToday = True
Const m_def_MinYear = 1900
Const m_def_MaxYear = 2999
Private Const m_def_Locked As Boolean = False
Private Const m_def_DTPValue As Date = "1/1/2000"

'Property Variables:
Dim m_DTPDay As Integer
Dim m_DTPMonth As Integer
Dim m_DTPYear As Integer
Dim m_ShowToday As Boolean
Dim m_MinYear As Long
Dim m_MaxYear As Long
Dim m_Locked As Boolean
Dim m_DTPValue As Date

'Event Declarations:
Event Change() 'MappingInfo=txtDay,txtDay,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."

Private Sub lvB_DropDown_Click()
   If Locked Then Exit Sub
   If isDropdowned Then Exit Sub
   
   Set DayEditBox = txtDay
   Set MonthEditBox = txtMonth
   Set YearEditBox = txtYear
   
   Set MyLV_B_DD = lvB_DropDown
   
   lvB_DropDown.Enabled = False
   
   isShowToday = m_ShowToday
   GetWindowRect hWnd, myRect
   frmDatePicker.Show vbModeless, Me
   isDropdowned = True
End Sub

Private Sub picHolder_Click()
   If lvB_DropDown.Value Then lvB_DropDown.Value = False
End Sub

Private Sub txtContainer_DblClick()
   lvB_DropDown_Click
End Sub

Private Sub txtContainer_GotFocus()
   If lvB_DropDown.Value Then lvB_DropDown.Value = False
End Sub

Private Sub txtDay_DblClick()
   lvB_DropDown_Click
End Sub

Private Sub txtMonth_DblClick()
   lvB_DropDown_Click
End Sub

Private Sub txtYear_DblClick()
   lvB_DropDown_Click
End Sub

Private Sub txtDay_Change()
   Dim iM As Integer
   
   HideCaret txtDay.hWnd
   If Len(txtDay) < 2 Then txtDay = Format(txtDay, "0#")
   
   Do
      iM = iM + 1
      
      If UCase(sMonth(iM)) = UCase(txtMonth) Then
         Exit Do
      End If
   Loop
   
   Let DTPValue = DateSerial(txtYear, iM, txtDay)
   RaiseEvent Change
End Sub

Private Sub txtDay_GotFocus()
   SelectText txtDay
   HideCaret txtDay.hWnd
End Sub

Private Sub txtDay_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   
   If Locked Then Exit Sub
   i = txtDay
   
   Select Case KeyCode
      Case vbKeyUp, vbKeyAdd, vbKeyPageUp
         If i = m_DaysCount(Month(DTPValue)) Then i = 0
         txtDay = Format(i + 1, "0#")
         
      Case vbKeyDown, vbKeySubtract, vbKeyPageDown
         If Shift = vbAltMask Then
            lvB_DropDown_Click
         Else
            If i = 1 Then i = m_DaysCount(Month(DTPValue)) + 1
            txtDay = Format(i - 1, "0#")
         End If
         
      Case vbKeyRight
         txtMonth.SetFocus
         
      Case vbKeyLeft
         txtYear.SetFocus
   End Select
   
   SelectText txtDay
   KeyCode = 0
End Sub

Private Sub txtDay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Ambient.UserMode Then SelectText txtDay
End Sub

Private Sub txtMonth_Change()
   Dim iM As Integer
   
   HideCaret txtMonth.hWnd
   
   If IsNumeric(txtMonth) Then
      txtMonth = sMonth(txtMonth)
   Else
      Do
         iM = iM + 1
         
         If UCase(sMonth(iM)) = UCase(txtMonth) Then
            Exit Do
         End If
      Loop
   End If
   
   txtYear.Left = txtMonth.Left + UserControl.TextWidth(txtMonth) + 75
   Let DTPValue = DateSerial(txtYear, iM, txtDay)
   RaiseEvent Change
End Sub

Private Sub txtMonth_GotFocus()
   SelectText txtMonth
   HideCaret txtMonth.hWnd
End Sub

Private Sub txtMonth_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   
   If Locked Then Exit Sub
   
   For i = 1 To 12
      If UCase(txtMonth) = UCase(sMonth(i)) Then Exit For
   Next i
   
   Select Case KeyCode
      Case vbKeyUp, vbKeyAdd, vbKeyPageUp
         If i = 12 Then i = 0
         i = i + 1
         
      Case vbKeyDown, vbKeySubtract, vbKeyPageDown
         If Shift = vbAltMask Then
            lvB_DropDown_Click
         Else
            If i = 1 Then i = 13
            i = i - 1
         End If
         
      Case vbKeyRight
         txtYear.SetFocus
         
      Case vbKeyLeft
         txtDay.SetFocus
   End Select
   
   txtMonth = sMonth(i)
   If txtDay > m_DaysCount(i) Then txtDay = m_DaysCount(i)
   SelectText txtMonth
   KeyCode = 0
End Sub

Private Sub txtMonth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Ambient.UserMode Then SelectText txtMonth
End Sub

Private Sub txtYear_Change()
   Dim iM As Integer
   
   HideCaret txtYear.hWnd
   
   If txtYear > m_MaxYear Then txtYear = m_MaxYear
   If txtYear < m_MinYear Then txtYear = m_MinYear
   
   ReDimDaysArray Val(txtYear.Text)
   
   Do
      iM = iM + 1
      
      If UCase(sMonth(iM)) = UCase(txtMonth) Then
         Exit Do
      End If
   Loop
   
   Let DTPValue = DateSerial(txtYear, iM, txtDay)
   RaiseEvent Change
End Sub

Private Sub txtYear_GotFocus()
   SelectText txtYear
   HideCaret txtYear.hWnd
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
   If Locked Then Exit Sub
   
   Select Case KeyCode
      Case vbKeyUp, vbKeyAdd
         If txtYear = MaxYear Then txtYear = MinYear Else txtYear = Val(txtYear) + 1
      Case vbKeyDown, vbKeySubtract
         If Shift = vbAltMask Then
            lvB_DropDown_Click
         Else
            If txtYear = MinYear Then txtYear = MaxYear Else txtYear = txtYear - 1
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
      
      Case vbKeyRight
         txtDay.SetFocus
         
      Case vbKeyLeft
         txtMonth.SetFocus
   End Select
   
   SelectText txtYear
   KeyCode = 0
End Sub

Private Sub txtYear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Ambient.UserMode Then SelectText txtYear
End Sub

Private Sub UserControl_DblClick()
   lvB_DropDown_Click
End Sub

Private Sub UserControl_Initialize()
   sMonth(1) = "January"
   sMonth(2) = "February"
   sMonth(3) = "March"
   sMonth(4) = "April"
   sMonth(5) = "May"
   sMonth(6) = "June"
   sMonth(7) = "July"
   sMonth(8) = "August"
   sMonth(9) = "September"
   sMonth(10) = "October"
   sMonth(11) = "November"
   sMonth(12) = "December"
End Sub

Private Sub UserControl_Resize()
   If UserControl.Height <> 315 Then UserControl.Height = 315
   If UserControl.Width < 1980 Then UserControl.Width = 1980
   
   txtContainer.Width = UserControl.Width
   picHolder.Left = UserControl.Width - 285
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   
   txtYear.Enabled = New_Enabled
   txtMonth.Enabled = New_Enabled
   txtDay.Enabled = New_Enabled
   lvB_DropDown.Enabled = New_Enabled

   PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
   Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   m_Locked = New_Locked
   PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,0,01/01/1900
Public Property Get DTPValue() As Date
Attribute DTPValue.VB_ProcData.VB_Invoke_Property = "Kalendar"
   DTPValue = m_DTPValue
End Property

Public Property Let DTPValue(ByVal New_DTPValue As Date)
   If Not IsDate(New_DTPValue) Then
      MsgBox "Format salah!"
      Exit Property
   End If
   
   m_DTPValue = New_DTPValue
   PropertyChanged "DTPValue"
   DTValue = New_DTPValue

   If txtDay <> Format(Day(New_DTPValue), "0#") Then txtDay = Format(Day(New_DTPValue), "0#")
   If txtMonth <> sMonth(Month(New_DTPValue)) Then txtMonth = sMonth(Month(New_DTPValue))
   If txtYear <> Year(New_DTPValue) Then txtYear = Year(New_DTPValue)
   
   m_DTPDay = Day(New_DTPValue)
   m_DTPMonth = Month(New_DTPValue)
   m_DTPYear = Year(New_DTPValue)
   
   RaiseEvent Change
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Locked = m_def_Locked
   m_DTPValue = m_def_DTPValue
   DTValue = DTPValue
   m_MinYear = m_def_MinYear
   m_MaxYear = m_def_MaxYear
   lMinYear = m_MinYear
   lMaxYear = m_MaxYear
   m_ShowToday = m_def_ShowToday
   m_DTPDay = m_def_DTPDay
   m_DTPMonth = m_def_DTPMonth
   m_DTPYear = m_def_DTPYear
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
   m_DTPValue = PropBag.ReadProperty("DTPValue", m_def_DTPValue)
   m_MinYear = PropBag.ReadProperty("MinYear", m_def_MinYear)
   m_MaxYear = PropBag.ReadProperty("MaxYear", m_def_MaxYear)
   
   lMinYear = m_MinYear
   lMaxYear = m_MaxYear
   
   txtYear.Enabled = Enabled
   txtMonth.Enabled = Enabled
   txtDay.Enabled = Enabled
   lvB_DropDown.Enabled = Enabled
      
   If Ambient.UserMode Then
      HookContext txtDay.hWnd
      HookContext txtMonth.hWnd
      HookContext txtYear.hWnd
   End If
   
   m_ShowToday = PropBag.ReadProperty("ShowToday", m_def_ShowToday)
   m_DTPDay = PropBag.ReadProperty("DTPDay", m_def_DTPDay)
   m_DTPMonth = PropBag.ReadProperty("DTPMonth", m_def_DTPMonth)
   m_DTPYear = PropBag.ReadProperty("DTPYear", m_def_DTPYear)
End Sub

Private Sub UserControl_Terminate()
   UnhookContext txtDay.hWnd
   UnhookContext txtMonth.hWnd
   UnhookContext txtYear.hWnd
   
   If isDropdowned Then Unload frmDatePicker
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
   Call PropBag.WriteProperty("DTPValue", m_DTPValue, m_def_DTPValue)
   Call PropBag.WriteProperty("MinYear", m_MinYear, m_def_MinYear)
   Call PropBag.WriteProperty("MaxYear", m_MaxYear, m_def_MaxYear)
   
   lMinYear = m_MinYear
   lMaxYear = m_MaxYear
   
   Call PropBag.WriteProperty("ShowToday", m_ShowToday, m_def_ShowToday)
   Call PropBag.WriteProperty("DTPDay", m_DTPDay, m_def_DTPDay)
   Call PropBag.WriteProperty("DTPMonth", m_DTPMonth, m_def_DTPMonth)
   Call PropBag.WriteProperty("DTPYear", m_DTPYear, m_def_DTPYear)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "Sends a command string to the source application in a DDE conversation."
   txtDay.LinkExecute Command
   txtMonth.LinkExecute Command
   txtYear.LinkExecute Command
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkMode
Public Property Get LinkMode() As Integer
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
   LinkMode = txtDay.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As Integer)
   txtDay.LinkMode() = New_LinkMode
   PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
   LinkItem = txtDay.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
   txtDay.LinkItem() = New_LinkItem
   PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "Transfers contents of Label, PictureBox, or TextBox to source application in DDE conversation."
   txtDay.LinkPoke
   txtMonth.LinkPoke
   txtYear.LinkPoke
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "Asks the source DDE application to update the contents of a Label, PictureBox, or Textbox control."
   txtDay.LinkRequest
   txtMonth.LinkRequest
   txtYear.LinkRequest
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "Transfers contents of PictureBox to destination application in DDE conversation."
   txtDay.LinkSend
   txtMonth.LinkSend
   txtYear.LinkSend
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
   LinkTimeout = txtDay.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
   txtDay.LinkTimeout() = New_LinkTimeout
   PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDay,txtDay,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
   LinkTopic = txtDay.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
   txtDay.LinkTopic() = New_LinkTopic
   PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MinYear() As Variant
'Attribute MinYear.VB_ProcData.VB_Invoke_Property = "Kalendar"
   MinYear = m_MinYear
End Property

Public Property Let MinYear(ByVal New_MinYear As Variant)
   m_MinYear = New_MinYear
   lMinYear = New_MinYear
   
   If txtYear < New_MinYear Then
      txtYear = New_MinYear
      txtYear_Change
   End If
   
   PropertyChanged "MinYear"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MaxYear() As Variant
'Attribute MaxYear.VB_ProcData.VB_Invoke_Property = "Kalendar"
   MaxYear = m_MaxYear
End Property

Public Property Let MaxYear(ByVal New_MaxYear As Variant)
   m_MaxYear = New_MaxYear
   lMaxYear = New_MaxYear
   
   If txtYear > New_MaxYear Then
      txtYear = New_MaxYear
      txtYear_Change
   End If
   
   PropertyChanged "MaxYear"
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
   MsgBox "Date Picker" & vbCrLf & "=======" & vbCrLf & vbCrLf & "CODER" & vbTab & "~»  Batavian" & vbCrLf & "REGION" & vbTab & "~»  Indonesia - Jakarta" & vbCrLf & "E-MAIL" & vbTab & "~»  batavian_codes@yahoo.com  " & vbCrLf & vbTab & "~»  batavian@xasamail.com ", vbOKOnly Or vbInformation, "About ctlDatePicker"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowToday() As Boolean
   ShowToday = m_ShowToday
End Property

Public Property Let ShowToday(ByVal New_ShowToday As Boolean)
   m_ShowToday = New_ShowToday
   PropertyChanged "ShowToday"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ShowDropDown()
   lvB_DropDown_Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get DTPDay() As Integer
   DTPDay = m_DTPDay
End Property

Public Property Let DTPDay(ByVal New_DTPDay As Integer)
   If New_DTPDay < 1 Or New_DTPDay > MaxDayOfMonth(DTPYear, DTPMonth) Then
      MsgBox "Format hari salah!"
      Exit Property
   End If
   
   m_DTPDay = New_DTPDay
   PropertyChanged "DTPDay"
   
   DTPValue = DateSerial(DTPYear, DTPMonth, New_DTPDay)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get DTPMonth() As Integer
   DTPMonth = m_DTPMonth
End Property

Public Property Let DTPMonth(ByVal New_DTPMonth As Integer)
   If New_DTPMonth < 1 Or New_DTPMonth > 12 Then
      MsgBox "Format bulan salah!"
      Exit Property
   End If
   
   m_DTPMonth = New_DTPMonth
   PropertyChanged "DTPMonth"
   
   DTPValue = DateSerial(DTPYear, New_DTPMonth, DTPDay)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1900
Public Property Get DTPYear() As Integer
   DTPYear = m_DTPYear
End Property

Public Property Let DTPYear(ByVal New_DTPYear As Integer)
   If New_DTPYear < MinYear Or New_DTPYear > MaxYear Then
      MsgBox "Format tahun salah!"
      Exit Property
   End If
   
   m_DTPYear = New_DTPYear
   PropertyChanged "DTPYear"
   
   DTPValue = DateSerial(New_DTPYear, DTPMonth, DTPDay)
End Property

