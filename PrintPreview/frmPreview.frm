VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picScroll 
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   600
      Width           =   7575
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
      Begin VB.HScrollBar hsPreview 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   4560
         Width           =   1725
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   240
         ScaleHeight     =   3615
         ScaleWidth      =   3540
         TabIndex        =   2
         Top             =   480
         Width           =   3540
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   480
            ScaleHeight     =   1185
            ScaleWidth      =   1665
            TabIndex        =   3
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.PictureBox picHold 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   2175
            TabIndex        =   4
            Top             =   120
            Width           =   2175
            Begin VB.PictureBox picDoc 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1215
               Left            =   240
               ScaleHeight     =   1185
               ScaleWidth      =   1665
               TabIndex        =   5
               Top             =   225
               Visible         =   0   'False
               Width           =   1695
            End
         End
      End
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      Begin VB.ComboBox cboPage 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3825
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   90
         Width           =   1620
      End
      Begin Project1.lvButtons_H lvB_Next 
         Height          =   405
         Left            =   3045
         TabIndex        =   14
         Top             =   45
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPreview.frx":27A2
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvB_Previous 
         Height          =   405
         Left            =   2700
         TabIndex        =   15
         Top             =   45
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   714
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
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPreview.frx":2AF4
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cboZoom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPreview.frx":2E46
         Left            =   780
         List            =   "frmPreview.frx":2E48
         TabIndex        =   12
         Text            =   "cboZoom"
         Top             =   90
         Width           =   1815
      End
      Begin VB.PictureBox picPages 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   3825
         ScaleHeight     =   255
         ScaleWidth      =   1155
         TabIndex        =   10
         Top             =   90
         Width           =   1215
         Begin VB.Label lblStatus 
            Caption         =   "Page :"
            Height          =   255
            Left            =   30
            TabIndex        =   11
            Top             =   15
            Width           =   1095
         End
      End
      Begin Project1.lvButtons_H lvB_Cancel 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   6435
         TabIndex        =   9
         Top             =   45
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   714
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPreview.frx":2E4A
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvB_Print 
         Default         =   -1  'True
         Height          =   405
         Left            =   5520
         TabIndex        =   13
         Top             =   45
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         Caption         =   "&Print"
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
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPreview.frx":319C
         cBack           =   -2147483633
      End
      Begin VB.Label lblView 
         AutoSize        =   -1  'True
         Caption         =   "Zoom :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   135
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HALFTONE As Long = 4

Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private mDocument As clsPrinter
Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer
Private bDisplayPage As Boolean

Public Property Set Document(ByVal vNewValue As clsPrinter)
   Set mDocument = vNewValue
End Property

Private Sub cboPage_Click()
   lPage = cboPage.ListIndex + 1
   Preview_Display lPage
End Sub

Private Sub cboZoom_Click()
   Dim iEvents As Integer

   If Not bScrollCode Then
      If cboZoom.ListIndex >= 0 Then
         iEvents = DoEvents
       
         If cboZoom.ItemData(cboZoom.ListIndex) <> sZoom Then
            sZoom = cboZoom.ItemData(cboZoom.ListIndex)
            Zoom_Check
         End If
      End If
   End If
End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)
   Dim sNewZoom As Single
   
   If KeyAscii = 13 Then
      sNewZoom = Val(cboZoom.Text)
      
      If sNewZoom > 0 And sNewZoom <= 200 Then
         cboZoom.Text = sNewZoom & " %"
         If sNewZoom = sZoom Then Exit Sub
         sZoom = sNewZoom
         Zoom_Check
      Else
         If cboZoom.ListIndex >= 0 Then cboZoom.Text = cboZoom.List(cboZoom.ListIndex) Else cboZoom.Text = sZoom & " %"
      End If
      
      Exit Sub
   End If

   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      ShowBalloonTip GetComboStruckHandle(cboZoom), "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
End Sub

Private Sub lvB_Cancel_Click()
   Set mDocument = Nothing
   Me.Hide
End Sub

Private Sub lvB_Next_Click()
   lPage = lPage + 1
   cboPage.ListIndex = lPage - 1
   Preview_Display lPage
End Sub

Private Sub lvB_Previous_Click()
   lPage = lPage - 1
   cboPage.ListIndex = lPage - 1
   Preview_Display lPage
End Sub

Private Sub lvB_Print_Click()
   Dim lPrintStart As Integer
   Dim lPrintEnd As Integer
   Dim iCopies As Integer
   Dim bCollate As Boolean
   Dim iPrinter As Integer

   frmPrint.Flags = mDocument.PrintOptions
   frmPrint.PageCurrent = lPage
   frmPrint.PageMax = lPageMax
   
   frmPrint.Show vbModal
   
   lPrintStart = frmPrint.PageStart
   lPrintEnd = frmPrint.PageEnd
   iCopies = frmPrint.Copies
   bCollate = frmPrint.Collate
   iPrinter = frmPrint.PrinterNumber

   If frmPrint.PrintDoc Then
      lblStatus.Caption = "Printing..."
      lblStatus.Refresh
      
      mDocument.PrintDoc lPrintStart, lPrintEnd, iCopies, bCollate, iPrinter
      
      lblStatus.Caption = "Page: " & lPage & " / " & lPageMax
      
      Unload frmPrint
      Unload Me
   Else
      Unload frmPrint
   End If
End Sub

Private Sub Form_Activate()
   Me.Refresh
   bDisplayPage = True
   Preview_Display lPage
End Sub

Private Sub Form_Load()
   Dim prtPrinter As Printer
   Dim iPrinter As Integer
   Dim sY As Single
   Dim intCount As Integer
   Dim sngAddedW As Single

'   HookCBContext GetComboStruckHandle(cboZoom)
   sZoom = 100
   
   With cboZoom
      .AddItem "100 %"
      .ItemData(.ListCount - 1) = 100
      .AddItem "75 %"
      .ItemData(.ListCount - 1) = 75
      .AddItem "50 %"
      .ItemData(.ListCount - 1) = 50
      .AddItem "Full page"
      .ItemData(.ListCount - 1) = 0
      .AddItem "Page width"
      .ItemData(.ListCount - 1) = -1
      bScrollCode = True
      .ListIndex = 4
      bScrollCode = False
   End With
   
   sZoom = 100
   lPage = 1
   lPageMax = mDocument.Pages
   
   For intCount = lPage To lPageMax
      cboPage.AddItem "Page-" & intCount
   Next intCount
   
   If bCustomPage Then
      For intCount = 0 To cboPage.ListCount - 1
         cboPage.List(intCount) = "Page " & intCount + 1 & " - " & strPage(intCount)
         
         If sngAddedW < picControl.TextWidth(cboPage.List(intCount)) Then
            sngAddedW = picControl.TextWidth(cboPage.List(intCount))
         End If
      Next intCount
      
      If sngAddedW > 9000 Then sngAddedW = 9000
      sngAddedW = sngAddedW - cboPage.Width + 390
      
      cboPage.Width = cboPage.Width + sngAddedW
      lvB_Print.Left = lvB_Print.Left + sngAddedW
      lvB_Cancel.Left = lvB_Cancel.Left + sngAddedW
   End If
   
   SetCBDropDownHeigth Me.hWnd, cboZoom, 5
   SetCBDropDownHeigth Me.hWnd, cboPage, IIf(lPageMax > 20, 20, lPageMax)
   
   cboPage.ListIndex = 0
   
   Call cboZoom_Click
'   SetCBListStyle cboZoom
   SetMenuForm Me.hWnd
   SetMinMaxInfo Me.hWnd, -1, -1, -1, -1, -1, -1, 640, 480
   
   Set frmVSParent = Me
   StartPicDocSubClass picDoc.hWnd
End Sub

Public Sub Preview_Display(ByVal iPage As Integer)
   Dim iMin As Integer
   Dim iMax As Integer
   
   Screen.MousePointer = vbHourglass
   picNormal.Cls
   picDoc.Visible = False
   mDocument.PreviewPage picNormal, iPage
   Preview_Status
   Zoom_Check
   Screen.MousePointer = vbDefault
End Sub

Private Sub Zoom_Check()
   Dim sSizeX As Single
   Dim sSizeY As Single
   Dim sRatio As Single
   Dim spImage As StdPicture
   Dim sWidth As Single
   Dim sHeight As Single
   Dim bScroll As Byte
   Dim bOldScroll As Byte

   Screen.MousePointer = vbHourglass
   sWidth = picScroll.ScaleWidth
   sHeight = picScroll.ScaleHeight

   Do
      bOldScroll = bScroll
      
      If sZoom = 0 Then
         sRatio = (sHeight - 480) / picNormal.Height
      ElseIf sZoom = -1 Then
         sRatio = (sWidth - 480) / picNormal.Width
      Else
         sRatio = sZoom / 100
      End If
   
      sSizeX = picNormal.Width * sRatio
      sSizeY = picNormal.Height * sRatio
      
      If sSizeX > sWidth And (bScroll And 1) <> 1 Then
         sHeight = sHeight - hsPreview.Height
         bScroll = bScroll + 1
      End If
      
      If sSizeY > sHeight And (bScroll And 2) <> 2 Then
         sWidth = sWidth - vsPreview.Width
         bScroll = bScroll + 2
      End If
   Loop While bOldScroll <> bScroll

   If sHeight > 15 Then
      vsPreview.Height = sHeight
      hsPreview.Width = sWidth

      picShow.Move 0, 0, sWidth, sHeight
      picDoc.Move 240, 240, sSizeX, sSizeY
      picDoc.Cls
      
      If GetStretchBltMode(picNormal.hDC) <> HALFTONE Then
         SetStretchBltMode picNormal.hDC, HALFTONE
      End If
      
      If GetStretchBltMode(picDoc.hDC) <> HALFTONE Then
         SetStretchBltMode picDoc.hDC, HALFTONE
      End If

      StretchBlt picDoc.hDC, 0, 0, picDoc.Width / Screen.TwipsPerPixelX, picDoc.Height / Screen.TwipsPerPixelY, picNormal.hDC, 0, 0, picNormal.Width / Screen.TwipsPerPixelX, picNormal.Height / Screen.TwipsPerPixelY, vbSrcCopy

      bScrollCode = True
      picHold.Move 0, 0, sSizeX + 480, sSizeY + 480
   End If
   
   If (bScroll And 2) = 2 Then
      vsPreview.Visible = True
      vsPreview.Max = (picHold.ScaleHeight - picShow.ScaleHeight) / 14.4 + 1
      vsPreview.Min = 0
      vsPreview.SmallChange = 14
      vsPreview.LargeChange = picShow.ScaleHeight / 14.4
      vsPreview.Value = vsPreview.Min
   Else
      vsPreview.Visible = False
   End If

   If (bScroll And 1) = 1 Then
      hsPreview.Visible = True
      hsPreview.Max = (picHold.ScaleWidth - picShow.ScaleWidth) / 14.4 + 1
      hsPreview.Min = 0
      hsPreview.SmallChange = 14
      hsPreview.LargeChange = picShow.ScaleWidth / 14.4
      hsPreview.Value = hsPreview.Min
   Else
      hsPreview.Visible = False
   End If
   
   bScrollCode = False
   Screen.MousePointer = vbDefault
   
   If bDisplayPage Then picDoc.Visible = True
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   If Me.ScaleHeight > 600 Then picScroll.Move 0, 480, Me.ScaleWidth, Me.ScaleHeight - 450
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode <> vbFormCode Then
      Me.Hide
      Cancel = 1
   End If
   
   EndPicDocSubClass picDoc.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bDisplayPage = False
End Sub

Private Sub hsPreview_Change()
   If Not bScrollCode Then picHold.Left = -hsPreview.Value * 14.4
End Sub

Private Sub picScroll_Resize()
   vsPreview.Move picScroll.ScaleWidth - vsPreview.Width, 0, 255, picScroll.ScaleHeight
   hsPreview.Move 0, picScroll.ScaleHeight - hsPreview.Height, picScroll.ScaleWidth
   Zoom_Check
End Sub

Private Sub vsPreview_Change()
   If Not bScrollCode Then picHold.Top = -vsPreview.Value * 14.4
End Sub

Public Sub Preview_Status()
   lvB_Previous.Enabled = CBool(lPage > 1)
   lvB_Next.Enabled = CBool(lPage < lPageMax)
   picPages.Cls
   lblStatus.Caption = "Hal.: " & lPage & " / " & lPageMax
   lblStatus.Visible = True
End Sub

Private Sub vsPreview_GotFocus()
   picScroll.SetFocus
End Sub

Private Sub hsPreview_GotFocus()
   picScroll.SetFocus
End Sub
