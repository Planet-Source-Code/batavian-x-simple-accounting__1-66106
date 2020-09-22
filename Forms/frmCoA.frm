VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCoA 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Chart of Accounts"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   Begin Project1.lvButtons_H lvButtons_H6 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Tag             =   "Print CoA List."
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmCoA.frx":0000
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   285
      Left            =   7050
      TabIndex        =   6
      Tag             =   "Close this window."
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   2
      cBack           =   255
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   330
      Index           =   1
      Left            =   1515
      TabIndex        =   1
      Top             =   540
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "Liabilities"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "Assets"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   330
      Index           =   2
      Left            =   2910
      TabIndex        =   2
      Top             =   540
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "Equities"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   330
      Index           =   3
      Left            =   4305
      TabIndex        =   3
      Top             =   540
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "Incomes"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin ComctlLib.ListView LV1 
      Height          =   3345
      Left            =   120
      TabIndex        =   4
      Top             =   930
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   5900
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code"
         Object.Width           =   926
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remark"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "D/C Pos"
         Object.Width           =   661
      EndProperty
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   330
      Index           =   4
      Left            =   5700
      TabIndex        =   5
      Top             =   540
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "Expenses"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H3 
      Height          =   285
      Left            =   6645
      TabIndex        =   7
      Tag             =   "Add new item."
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   2
      Image           =   "frmCoA.frx":0352
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H4 
      Height          =   285
      Left            =   6315
      TabIndex        =   8
      Tag             =   "Edit selected item."
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   2
      Image           =   "frmCoA.frx":06A4
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H5 
      Height          =   285
      Left            =   5985
      TabIndex        =   9
      Tag             =   "Delete selected item."
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   2
      Image           =   "frmCoA.frx":09F6
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmCoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SPI_GETWORKAREA As Long = 48

Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private bItemClicked As Boolean
Private blvBEnabled As Boolean
Private intCategoryIndex As Integer
Private bChildLoaded As Boolean

Public intCoAMode As Integer
Public frmParent As Form

Private Sub RefreshForm()
   Dim sL As Single
   Dim iC As Integer
   Dim hRgn As Long
   Dim hBrush As Long

   Me.Cls

   iC = 47

   For sL = 0 To 15
      Me.Line (0, sL)-(Me.Width / Screen.TwipsPerPixelX, sL), RGB(iC, iC, iC)
      iC = iC + 6
   Next sL

   For sL = 15 To 30
      Me.Line (0, sL)-(Me.Width / Screen.TwipsPerPixelX, sL), RGB(iC, iC, iC)
      iC = iC - 5
   Next sL

   Me.FontBold = True
   Me.CurrentX = ((Me.Width / Screen.TwipsPerPixelX) - Me.TextWidth(Me.Caption)) / 2
   Me.CurrentY = 9
   Me.Print Me.Caption
   Me.Line (0, 30)-(Me.Width / Screen.TwipsPerPixelX, 30), vbWhite

   hRgn = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, (Me.Height / Screen.TwipsPerPixelY), 6, 6)
   hBrush = CreateSolidBrush(RGB(31, 31, 31))
   FrameRgn Me.hDC, hRgn, hBrush, 1, 1

   DeleteObject hRgn
   DeleteObject hBrush

   hRgn = CreateRoundRectRgn(1, 1, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, 6, 6)
   hBrush = CreateSolidBrush(RGB(255, 255, 255))
   FrameRgn Me.hDC, hRgn, hBrush, 1, 1

   DeleteObject hRgn
   DeleteObject hBrush
End Sub

Public Sub FillLV()
   Call lvButtons_H1_Click(0)
End Sub

Private Sub Form_Activate()
   bChildLoaded = False
End Sub

Private Sub Form_Deactivate()
   If Not bChildLoaded Then Unload Me
End Sub

Private Sub Form_Load()
   Dim hRgn As Long
   Dim rct As RECT
   Dim lLeft As Long
   Dim lTop As Long
   
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
   
   hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 6, 6)
   SetWindowRgn Me.hWnd, hRgn, True
   
   SetLVExtendedStyle LV1
   
   Call lvButtons_H1_Click(0)
   
   RefreshForm
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton And y < 30 Then
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
   End If
End Sub

Private Sub Form_Paint()
'   RefreshForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set MyTBDesc = Nothing
   Set MyTBRem = Nothing
   
   MyLV_B_DD.Value = False
   MyLV_B_DD.Enabled = True
   
   Set MyLV_B_DD = Nothing
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
   SortListByColumn LV1, ColumnHeader
End Sub

Private Sub LV1_DblClick()
   If bItemClicked Then
      Unload Me
   End If
   
   bItemClicked = False
End Sub

Private Sub LV1_ItemClick(ByVal Item As ComctlLib.ListItem)
   bItemClicked = True
   
   MyTBDesc = LV1.SelectedItem.SubItems(1)
   MyTBRem = LV1.SelectedItem.SubItems(2)
   MyTBRem.Tag = lvButtons_H1(intCategoryIndex).Caption
   
   blvBEnabled = True
End Sub

Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      If bItemClicked Then
         frmMain.mnuSub3(2).Enabled = True
         frmMain.mnuSub3(3).Enabled = True
      Else
         frmMain.mnuSub3(2).Enabled = False
         frmMain.mnuSub3(3).Enabled = False
      End If
      
      PopupMenu frmMain.mnuMain3
      bItemClicked = False
   End If
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
   Dim LV As ListItem
   
   intCategoryIndex = Index
   
   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblCoA WHERE Category = '" & lvButtons_H1(Index).Caption & "';", ADOCnn
   
   LV1.ListItems.Clear
   
   Do Until ADORS.EOF
      Set LV = LV1.ListItems.Add(, , "")
      
      LV.Tag = ADORS!ID
      LV.Text = ADORS!CoA
      LV.SubItems(1) = ADORS!Description
      LV.SubItems(2) = ADORS!Remark
      LV.SubItems(3) = IIf(ADORS!IsDebt, "Debit", "Credit")
      
      ADORS.MoveNext
   Loop
   
   blvBEnabled = False
End Sub

Private Sub lvButtons_H1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H1(Index).hWnd And _
      MyCoolTip.TipText = lvButtons_H1(Index).Caption Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H1(Index).hWnd
   MyCoolTip.TipText = lvButtons_H1(Index).Caption
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H1_MouseOnButton(Index As Integer, OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub lvButtons_H2_Click()
   Unload Me
End Sub

Private Sub lvButtons_H2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H2.hWnd And _
      MyCoolTip.TipText = lvButtons_H2.Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H2.hWnd
   MyCoolTip.TipText = lvButtons_H2.Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H2_MouseOnButton(OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Public Sub lvButtons_H3_Click()
   frmNewCoA.bNewItem = True
   frmNewCoA.Show vbModeless, Me
   
   bChildLoaded = True
End Sub

Private Sub lvButtons_H3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H3.hWnd And _
      MyCoolTip.TipText = lvButtons_H3.Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H3.hWnd
   MyCoolTip.TipText = lvButtons_H3.Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H3_MouseOnButton(OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Public Sub lvButtons_H4_Click()
   If blvBEnabled Then
      If bItemClicked Then
         frmNewCoA.bNewItem = False
         frmNewCoA.strCategory = lvButtons_H1(intCategoryIndex).Caption
         frmNewCoA.strRemark = LV1.SelectedItem.SubItems(2)
         frmNewCoA.strCode = LV1.SelectedItem.Text
         frmNewCoA.strDescription = LV1.SelectedItem.SubItems(1)
         frmNewCoA.bDebit = LV1.SelectedItem.SubItems(3) = "Debit"
         frmNewCoA.lngID = LV1.SelectedItem.Tag
         frmNewCoA.Show vbModeless, Me
         
         bChildLoaded = True
      End If
   Else
      xMsgBox Me.hWnd, "There are no currently selected item!", vbInformation, "Invalid Action"
   End If
End Sub

Private Sub lvButtons_H4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H4.hWnd And _
      MyCoolTip.TipText = lvButtons_H4.Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H4.hWnd
   MyCoolTip.TipText = lvButtons_H4.Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H4_MouseOnButton(OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Public Sub lvButtons_H5_Click()
   If blvBEnabled Then
      If xMsgBox(Me.hWnd, "Are you sure to delete selected item? " & vbCrLf & vbCrLf & _
         "Category: " & lvButtons_H1(intCategoryIndex).Caption & " " & vbCrLf & _
         "Code: " & LV1.SelectedItem.Text & " " & vbCrLf & _
         "Description: " & LV1.SelectedItem.SubItems(1) & " " & vbCrLf & _
         "Remark: " & LV1.SelectedItem.SubItems(2) & " " & vbCrLf & _
         "Normal Pos: " & IIf(LV1.SelectedItem.SubItems(2) = "Debit", "Debit", "Credit"), _
         vbQuestion Or vbYesNo, "Delete Item") = vbYes Then
         
         CloseConn , Me.hWnd
         ADORS.Open "SELECT * FROM tblJournal WHERE (DebitCoA = '" & LV1.SelectedItem.SubItems(1) & "' " & _
                    "OR CreditCoA = '" & LV1.SelectedItem.SubItems(1) & "') " & _
                    "AND Remark = '" & LV1.SelectedItem.SubItems(2) & "' " & _
                    "AND Category = '" & lvButtons_H1(intCategoryIndex).Caption & "';", ADOCnn
                    
         If ADORS.RecordCount > 0 Then
            xMsgBox Me.hWnd, "You cannot delete items which is currently in use. ", vbInformation Or vbOKOnly, "Invalid Action"
            Exit Sub
         Else
            CloseConn , Me.hWnd
            ADORS.Open "SELECT * FROM tblCoA WHERE ID = " & CLng(LV1.SelectedItem.Tag) & ";", ADOCnn, adOpenDynamic, adLockOptimistic
            ADORS.Delete
            ADORS.Update
            ADORS.UpdateBatch
            
            LV1.ListItems.Remove LV1.SelectedItem.Index
            
            MyTBDesc = vbNullString
            MyTBRem = vbNullString
         End If
      End If
   Else
      xMsgBox Me.hWnd, "There are nothing selected! ", vbInformation, "Invalid Action"
   End If
End Sub

Private Sub lvButtons_H5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H5.hWnd And _
      MyCoolTip.TipText = lvButtons_H5.Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H5.hWnd
   MyCoolTip.TipText = lvButtons_H5.Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H5_MouseOnButton(OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub lvButtons_H6_Click()
   Me.Enabled = False

   Const strProgress As String = "Printing, please wait ...."

   Dim MyPrinter As New clsPrinter
   Dim intCount As Integer
   
   frmPrintProgress.Show vbModeless, frmParent
   frmPrintProgress.Label1 = strProgress & vbCrLf & "0% complete."
   DoEvents

   AddReportLogo frmMain.picLogo, "Chart of Accounts", MyPrinter, qPortrait
   
   MyPrinter.MarginBottom = 600
   MyPrinter.AddText " ", "Arial", 24
   MyPrinter.AddText "<LINDENT=210>Category<LINDENT=1450>Code<LINDENT=2400>Description<LINDENT=7500>Remark", "Arial", 12, True
   MyPrinter.AddText " ", "Arial", 8
   MyPrinter.AddLine 180, 1150, 10650, 1150, 1
   MyPrinter.AddLine 180, 1105, 10650, 1105, 1, 2
   MyPrinter.Footer(EvenPage_hf).Text = "Page #pagenumber# of #pagetotal#"
   MyPrinter.Footer(OddPage_hf).Text = "Page #pagenumber# of #pagetotal#"
   MyPrinter.Footer(OddPage_hf).Alignment = eCentre
   MyPrinter.Footer(EvenPage_hf).Alignment = eCentre
   MyPrinter.Footer(OddPage_hf).FontSize = 7
   MyPrinter.Footer(EvenPage_hf).FontSize = 7
   MyPrinter.SetFooter(OddPage_hf) = True
   MyPrinter.SetFooter(EvenPage_hf) = True

   CloseConn , Me.hWnd
   ADORS.Open "SELECT Category, CoA, Description, Remark From tblCoA ORDER BY CoA;", ADOCnn

   Do Until ADORS.EOF
      frmPrintProgress.Label1 = strProgress & vbCrLf & Round((intCount / ADORS.RecordCount) * 100) & "% complete."
      MyPrinter.AddText "<LINDENT=210>" & ADORS!Category & "<LINDENT=1450>" & ADORS!CoA & "<LINDENT=2400>" & ADORS!Description & "<LINDENT=7500>" & ADORS!Remark, "Arial", 10
      ADORS.MoveNext
      intCount = intCount + 1
      DoEvents
   Loop

   MyPrinter.Preview frmParent, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
   
   Unload frmPrintProgress
   
   Set MyPrinter = Nothing

   Me.Enabled = True
End Sub

Private Sub lvButtons_H6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H6.hWnd And _
      MyCoolTip.TipText = lvButtons_H6.Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H6.hWnd
   MyCoolTip.TipText = lvButtons_H6.Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H6_MouseOnButton(OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub
