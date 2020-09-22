VERSION 5.00
Begin VB.Form frmOpeningBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Balance"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   Icon            =   "frmOpeningBalance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   1
      Left            =   184
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   618
      TabIndex        =   12
      Tag             =   "Opening Balance - Chart of Accounts"
      Top             =   958
      Width           =   9300
      Begin VB.TextBox txtTotalC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7275
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   990
         Width           =   1695
      End
      Begin VB.TextBox txtTotalD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5565
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   990
         Width           =   1695
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   0
         Left            =   4875
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "Add a field."
         Top             =   285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         CapAlign        =   2
         BackStyle       =   4
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Image           =   "frmOpeningBalance.frx":000C
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   1
         Left            =   4635
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Remove the last field."
         Top             =   285
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   476
         CapAlign        =   2
         BackStyle       =   4
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Image           =   "frmOpeningBalance.frx":035E
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.PictureBox picScroller 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -90
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   651
         TabIndex        =   18
         Top             =   600
         Width           =   9765
         Begin VB.VScrollBar VS1 
            Enabled         =   0   'False
            Height          =   330
            LargeChange     =   2
            Left            =   9075
            Max             =   0
            TabIndex        =   19
            Top             =   0
            Width           =   285
         End
         Begin VB.PictureBox picGrabber 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            ScaleHeight     =   22
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   608
            TabIndex        =   23
            Top             =   0
            Width           =   9120
            Begin VB.TextBox txtValueCredit 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   7245
               TabIndex        =   1
               Top             =   15
               Width           =   1695
            End
            Begin VB.TextBox txtAccount 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   375
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   15
               Width           =   2610
            End
            Begin VB.TextBox txtValueDebit 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   0
               Left            =   5535
               TabIndex        =   0
               Top             =   15
               Width           =   1695
            End
            Begin Project1.lvButtons_H lvB_CoA 
               Height          =   240
               Index           =   0
               Left            =   5235
               TabIndex        =   25
               TabStop         =   0   'False
               Tag             =   "Show Chart of Accounts List."
               Top             =   45
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   423
               CapAlign        =   2
               BackStyle       =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Focus           =   0   'False
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               Image           =   "frmOpeningBalance.frx":06B0
               ImgSize         =   40
               cBack           =   -2147483633
            End
            Begin VB.TextBox txtRemark 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   3000
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   15
               Width           =   2520
            End
            Begin VB.Label lblNumber 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
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
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   60
               Width           =   315
            End
         End
      End
      Begin VB.Label lblCoAFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Values (Rp.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   3
         Left            =   7515
         TabIndex        =   29
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   345
         Width           =   120
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
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
         Left            =   4995
         TabIndex        =   21
         Tag             =   "Scrollabled"
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts"
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
         Index           =   0
         Left            =   510
         TabIndex        =   17
         Top             =   345
         Width           =   660
      End
      Begin VB.Label lblCoAFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Values (Rp.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   1
         Left            =   5835
         TabIndex        =   16
         Top             =   345
         Width           =   1305
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Index           =   2
         Left            =   3105
         TabIndex        =   15
         Top             =   345
         Width           =   615
      End
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Index           =   2
      Left            =   6102
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   9
      Tag             =   "Date"
      Top             =   186
      Width           =   3360
      Begin Project1.ctlDatePicker ctlDatePicker1 
         Height          =   315
         Left            =   795
         TabIndex        =   4
         Top             =   285
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   5
         Left            =   690
         TabIndex        =   11
         Top             =   330
         Width           =   60
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   7
         Left            =   105
         TabIndex        =   10
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Index           =   0
      Left            =   177
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   5
      Tag             =   "Administrative"
      Top             =   186
      Width           =   3360
      Begin VB.TextBox txtVoucher 
         BackColor       =   &H00E0E0E0&
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
         Left            =   795
         Locked          =   -1  'True
         MaxLength       =   12
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   2460
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slip No."
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
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   0
         Left            =   690
         TabIndex        =   7
         Top             =   330
         Width           =   60
      End
   End
   Begin Project1.lvButtons_H lvB_OK 
      Height          =   420
      Left            =   5415
      TabIndex        =   2
      Top             =   2484
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   741
      Caption         =   "&OK"
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
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmOpeningBalance.frx":0A02
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   7260
      TabIndex        =   3
      Top             =   2484
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   741
      Caption         =   "&Cancel"
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
      Image           =   "frmOpeningBalance.frx":1114
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private intCtlCount As Integer

Private bFlagDb() As Boolean
Private bFlagCr() As Boolean

Private bFlagTotalDb As Boolean
Private bFlagTotalCr As Boolean

Private Sub ctlDatePicker1_Change()
   txtVoucher = GenerateFullSlip(Right(ctlDatePicker1.DTPYear, 2), Format(ctlDatePicker1.DTPMonth, "00"), Format(ctlDatePicker1.DTPDay, "00")) & "." & GenerateSlipNo(ctlDatePicker1.DTPYear, ctlDatePicker1.DTPMonth, ctlDatePicker1.DTPDay)
End Sub

Private Sub Form_Load()
   Dim intCount As Integer
   
   For intCount = 0 To 2
      DrawPicHolderCaption picHolder(intCount)
   Next intCount
   
   ctlDatePicker1.DTPValue = DateSerial(Year(Date), 1, 1)
'   intCurrRecord = GenerateSlipNo(Year(Date), Month(Date), Day(Date), Me.hWnd)
'   txtVoucher = GenerateFullSlip(Right(Year(Date), 2), Format(Month(Date), "00"), Format(Day(Date), "00"))
   txtVoucher = GenerateFullSlip(Right(ctlDatePicker1.DTPYear, 2), Format(ctlDatePicker1.DTPMonth, "00"), Format(ctlDatePicker1.DTPDay, "00")) & "." & GenerateSlipNo(ctlDatePicker1.DTPYear, ctlDatePicker1.DTPMonth, ctlDatePicker1.DTPDay)
   intCtlCount = 0
   
   ReDim bFlagDb(0)
   ReDim bFlagCr(0)
   
   bFlagDb(0) = True
   bFlagCr(0) = True
   
   lvButtons_H2_Click 0
   SetMenuForm Me.hWnd
   
   SetReadonly txtAccount(0).hWnd
   SetReadonly txtRemark(0).hWnd
   SetReadonly txtTotalD.hWnd
   SetReadonly txtTotalC.hWnd
   SetReadonly txtVoucher.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set MyActiveTB = Nothing
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_CoA_Click(Index As Integer)
   Set MyLV_B_DD = lvB_CoA(Index)
   lvB_CoA(Index).Enabled = False
   LoadCoA txtAccount(Index), txtRemark(Index)
End Sub

Private Sub lvB_CoA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvB_CoA(Index).hWnd And _
      MyCoolTip.TipText = lvB_CoA(Index).Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvB_CoA(Index).hWnd
   MyCoolTip.TipText = lvB_CoA(Index).Tag
   MyCoolTip.Create
End Sub

Private Sub lvB_CoA_MouseOnButton(Index As Integer, OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub lvB_OK_Click()
   Dim intCount As Integer
   Dim intCount2 As Integer
'   Dim dblDebit As Double
'   Dim dblCredit As Double
   Dim strTemp As String
   
   For intCount = 0 To intCtlCount
      If Trim(txtAccount(intCount)) = vbNullString Then
         EnsureAccountVisible intCount
         txtRemark(intCount) = String(50, " ")
         txtRemark(intCount).SelStart = Len(txtRemark(intCount))
         ShowBalloonTip txtRemark(intCount).hWnd, "Invalid Field", "Accounts & Remarks field must be filled!" & vbCrLf & "Click here to select from the list!", [2-Information Icon]
         txtRemark(intCount) = vbNullString
         Exit Sub
      End If
      
      If Trim(txtValueDebit(intCount)) = vbNullString And Trim(txtValueCredit(intCount)) = vbNullString Then
         EnsureAccountVisible intCount
         ShowBalloonTip txtValueDebit(intCount).hWnd, "Invalid Debit/Credit Amount", "Please fill debit or credit field!", [2-Information Icon]
         Exit Sub
      End If
      
      If Trim(txtValueDebit(intCount)) <> vbNullString And Trim(txtValueCredit(intCount)) <> vbNullString Then
         EnsureAccountVisible intCount
         ShowBalloonTip txtValueDebit(intCount).hWnd, "Invalid Debit/Credit Amount", "Please fill one field only!", [2-Information Icon]
         Exit Sub
      End If
      
      If Trim(txtValueDebit(intCount)) = "0" Then
         EnsureAccountVisible intCount
         ShowBalloonTip txtValueDebit(intCount).hWnd, "Invalid Debit Amount", "Debit amount cannot be zero!", [2-Information Icon]
         Exit Sub
      End If
      
      If Trim(txtValueCredit(intCount)) = "0" Then
         EnsureAccountVisible intCount
         ShowBalloonTip txtValueCredit(intCount).hWnd, "Invalid Credit Amount", "Credit amount cannot be zero!", [2-Information Icon]
         Exit Sub
      End If
      
'      If Trim(txtValueDebit(intCount)) = vbNullString Then
'         strTemp = Replace(Replace(Trim(txtValueCredit(intCount)), ".", vbNullString), ",", vbNullString)
'         dblCredit = dblCredit + CDbl(IIf(strTemp = vbNullString, 0, strTemp))
'      Else
'         strTemp = Replace(Replace(Trim(txtValueDebit(intCount)), ".", vbNullString), ",", vbNullString)
'         dblDebit = dblDebit + CDbl(IIf(strTemp = vbNullString, 0, strTemp))
'      End If
   Next intCount
   
   For intCount = 0 To intCtlCount
      For intCount2 = 0 To intCtlCount
         If intCount2 <> intCount Then
            If (txtAccount(intCount) = txtAccount(intCount2)) And (txtRemark(intCount) = txtRemark(intCount2)) Then
               EnsureAccountVisible intCount2
               ShowBalloonTip txtAccount(intCount2).hWnd, "Same Account Name", "This account name with current remark has already been used!", [2-Information Icon]
               Exit Sub
            End If
         End If
      Next intCount2
   Next intCount
   
   If (txtTotalD <> txtTotalC) Or (bFlagTotalDb <> bFlagTotalCr) Then
      ShowBalloonTip txtTotalD.hWnd, "Invalid Debit/Credit Amount", "Debit and credit amount must be identical!", [2-Information Icon]
      Exit Sub
   End If

   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblJournal WHERE InternalID = 'OB';", ADOCnn, adOpenDynamic, adLockOptimistic

   If ADORS.RecordCount > 0 Then
      Do Until ADORS.EOF
         ADORS.Delete
         ADORS.MoveNext
      Loop
   End If

   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblJournal;", ADOCnn, adOpenDynamic, adLockOptimistic

   For intCount = 0 To intCtlCount
      ADORS.AddNew
      ADORS!Voucher = txtVoucher
      ADORS!Date = ctlDatePicker1.DTPValue
      
      If Trim(txtValueDebit(intCount)) = vbNullString Then
         strTemp = Replace(Replace(Trim(txtValueCredit(intCount)), ".", vbNullString), ",", vbNullString)
         
         ADORS!CreditCoa = txtAccount(intCount)
         ADORS!Credit = CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagCr(intCount), 1, -1)
      Else
         strTemp = Replace(Replace(Trim(txtValueDebit(intCount)), ".", vbNullString), ",", vbNullString)
         
         ADORS!DebitCoA = txtAccount(intCount)
         ADORS!Debit = CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagDb(intCount), 1, -1)
      End If
      
      ADORS!Description = "Opening Balance"
      ADORS!Remark = txtRemark(intCount)
      ADORS!Category = txtRemark(intCount).Tag
      ADORS!InternalID = "OB"
   Next intCount
   
   ADORS.Update
   ADORS.UpdateBatch
   
   frmMain.FillMainLV
   Unload Me
End Sub

'Private Sub LoadCalc(MyTB As TextBox)
'   Dim MyPA As POINTAPI
'   Dim strTemp As String
'   Dim dblTemp As Double
'
'   Set MyTBDesc = MyTB
'   ClientToScreen MyTB.hWnd, MyPA
'
'   With frmCalc
'      strTemp = Replace(MyTB, ".", vbNullString)
'      strTemp = Replace(strTemp, ".", vbNullString)
'
'      If Trim(strTemp) = vbNullString Then strTemp = 0
'
'      dblTemp = CDbl(GetSetting("Batavian's Accounting Program", "Setting", "Rate", "0"))
'      dblTemp = Int(CDbl(strTemp) / dblTemp)
'
'      .txtCalc(1) = Format(dblTemp, "#,###")
'
'      .Left = (MyPA.x + MyTB.Width + 2) * Screen.TwipsPerPixelX
'      .Top = (MyPA.y) * Screen.TwipsPerPixelY
'
'      If .Top + .Height > Screen.Height Then
'         .Top = Screen.Height - .Height
'      End If
'
'      If .Top < 0 Then .Top = 0
'
'      If .Left + .Width > Screen.Width Then
'         .Top = Screen.Width - .Width
'      End If
'
'      If .Left < 0 Then .Left = 0
'
'      .Show vbModeless, Me
'   End With
'End Sub

Private Sub LoadCoA(MyTBx As TextBox, MyTBy As TextBox)
   Set MyTBDesc = MyTBx
   Set MyTBRem = MyTBy
   
   GetWindowRect MyTBx.hWnd, myRect

   Set frmCoA.frmParent = Me
   frmCoA.Show vbModeless, Me
End Sub

Private Sub lvButtons_H2_Click(Index As Integer)
   If Index = 0 Then
      intCtlCount = intCtlCount + 1
      
      ReDim Preserve bFlagDb(intCtlCount)
      ReDim Preserve bFlagCr(intCtlCount)
      
      bFlagDb(intCtlCount) = True
      bFlagCr(intCtlCount) = True
      
      Load txtAccount(intCtlCount)
      txtAccount(intCtlCount) = vbNullString
      txtAccount(intCtlCount).Top = txtAccount(intCtlCount - 1).Top + 21
      txtAccount(intCtlCount).Visible = True
      
      Load txtRemark(intCtlCount)
      txtRemark(intCtlCount) = vbNullString
      txtRemark(intCtlCount).Top = txtRemark(intCtlCount - 1).Top + 21
      txtRemark(intCtlCount).Visible = True
      
      Load txtValueDebit(intCtlCount)
      txtValueDebit(intCtlCount).ForeColor = vbBlack
      txtValueDebit(intCtlCount) = vbNullString
      txtValueDebit(intCtlCount).TabIndex = txtValueCredit(intCtlCount - 1).TabIndex + 1
      txtValueDebit(intCtlCount).Top = txtValueDebit(intCtlCount - 1).Top + 21
      txtValueDebit(intCtlCount).Visible = True
      
      Load txtValueCredit(intCtlCount)
      txtValueCredit(intCtlCount).ForeColor = vbBlack
      txtValueCredit(intCtlCount) = vbNullString
      txtValueCredit(intCtlCount).TabIndex = txtValueDebit(intCtlCount).TabIndex + 1
      txtValueCredit(intCtlCount).Top = txtValueCredit(intCtlCount - 1).Top + 21
      txtValueCredit(intCtlCount).Visible = True
      
      Load lvB_CoA(intCtlCount)
      lvB_CoA(intCtlCount).Visible = True
      lvB_CoA(intCtlCount).ZOrder
      lvB_CoA(intCtlCount).Top = lvB_CoA(intCtlCount - 1).Top + 21
      
      Load lblNumber(intCtlCount)
      lblNumber(intCtlCount) = intCtlCount + 1
      lblNumber(intCtlCount).Top = lblNumber(intCtlCount - 1).Top + 21
      lblNumber(intCtlCount).Visible = True
     
      SetReadonly txtAccount(intCtlCount).hWnd
      SetReadonly txtRemark(intCtlCount).hWnd
      
      picGrabber.Height = picGrabber.Height + 21
      
      If intCtlCount < 13 Then
         picHolder(1).Height = picHolder(1).Height + 21
         
         txtTotalD.Top = txtTotalD.Top + 21
         txtTotalC.Top = txtTotalC.Top + 21
         
         lblTotal.Top = lblTotal.Top + 21
         
         picScroller.Height = picScroller.Height + 21
         VS1.Height = VS1.Height + 21
         Me.Height = Me.Height + 21 * Screen.TwipsPerPixelY
         
         lvB_OK.Top = lvB_OK.Top + 21
         lvB_Cancel.Top = lvB_Cancel.Top + 21
      End If
   Else
      If intCtlCount = 1 Then Exit Sub
      
      SetRewrite txtAccount(intCtlCount).hWnd
      SetRewrite txtRemark(intCtlCount).hWnd
      
      Unload txtAccount(intCtlCount)
      Unload txtRemark(intCtlCount)
      Unload txtValueDebit(intCtlCount)
      Unload txtValueCredit(intCtlCount)
      Unload lvB_CoA(intCtlCount)
      Unload lblNumber(intCtlCount)
      
      picGrabber.Height = picGrabber.Height - 21
      
      If intCtlCount <= 13 Then VS1.Enabled = False
      
      If intCtlCount < 13 Then
         VS1.Enabled = False
         picHolder(1).Height = picHolder(1).Height - 21
         
         txtTotalD.Top = txtTotalD.Top - 21
         txtTotalC.Top = txtTotalC.Top - 21
            
         lblTotal.Top = lblTotal.Top - 21
      
         lvB_OK.Top = lvB_OK.Top - 21
         lvB_Cancel.Top = lvB_Cancel.Top - 21
         
         Me.Height = Me.Height - 21 * Screen.TwipsPerPixelY
         picScroller.Height = picScroller.Height - 21
         VS1.Height = VS1.Height - 21
      End If

      intCtlCount = intCtlCount - 1
      
      ReDim Preserve bFlagDb(intCtlCount)
      ReDim Preserve bFlagCr(intCtlCount)
   End If
      
   If intCtlCount >= 13 Then
      VS1.Enabled = True
      VS1.Min = 0
      VS1.Max = picGrabber.Height - VS1.Height
      VS1.LargeChange = IIf((VS1.Max / 21) > 1, 42, 21)
      VS1.SmallChange = 21
   Else
      picGrabber.Top = 0
   End If
End Sub

Private Sub lvButtons_H2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H2(Index).hWnd And _
      MyCoolTip.TipText = lvButtons_H2(Index).Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H2(Index).hWnd
   MyCoolTip.TipText = lvButtons_H2(Index).Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H2_MouseOnButton(Index As Integer, OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub txtAccount_GotFocus(Index As Integer)
   HideCaret txtAccount(Index).hWnd
End Sub

Private Sub txtRemark_GotFocus(Index As Integer)
   HideCaret txtRemark(Index).hWnd
End Sub

Private Sub txtTotalC_GotFocus()
   HideCaret txtTotalC.hWnd
End Sub

Private Sub txtTotalD_GotFocus()
   HideCaret txtTotalD.hWnd
End Sub

Private Sub txtValueCredit_GotFocus(Index As Integer)
   txtValueCredit(Index) = Replace(Replace(txtValueCredit(Index), ",", vbNullString), ".", vbNullString)
   txtValueCredit(Index).SelStart = 0
   txtValueCredit(Index).SelLength = Len(txtValueCredit(Index))
   
   Set MyActiveTB = txtValueCredit(Index)
End Sub

Private Sub txtValueCredit_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 45 Then
      KeyAscii = 0
      
      If txtValueCredit(Index).ForeColor = vbBlack Then
         txtValueCredit(Index).ForeColor = vbRed
      Else
         txtValueCredit(Index).ForeColor = vbBlack
      End If
      
      GoTo ExitSub
   End If
   
   If KeyAscii = 43 Then
      KeyAscii = 0
      txtValueCredit(Index).ForeColor = vbBlack
      GoTo ExitSub
   End If
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
      ShowBalloonTip txtValueCredit(Index).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
   
   Exit Sub
   
ExitSub:
   bFlagCr(Index) = txtValueCredit(Index).ForeColor = vbBlack
End Sub

Private Sub txtValueCredit_LostFocus(Index As Integer)
   Dim intCount As Integer
   Dim strTemp As String
   Dim dblCredit As Double
   
   For intCount = 0 To intCtlCount
      strTemp = Replace(Replace(Trim(txtValueCredit(intCount)), ",", vbNullString), ".", vbNullString)
      
      If Not IsNumeric(strTemp) And strTemp <> vbNullString Then
         txtValueCredit(intCount).SelStart = 0
         txtValueCredit(intCount).SelLength = Len(txtValueCredit(intCount))
         ShowBalloonTip txtValueCredit(intCount).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
         Exit Sub
      End If
      
      dblCredit = dblCredit + (CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagCr(intCount), 1, -1))
   Next
   
   bFlagTotalCr = dblCredit > 0
   
   If dblCredit < 0 Then
      dblCredit = dblCredit * -1
      txtTotalC.ForeColor = vbRed
   Else
      txtTotalC.ForeColor = vbBlack
   End If
   
   txtTotalC = CStr(dblCredit)
   
   If Len(txtTotalC) > 3 Then
      txtTotalC = Format(txtTotalC, "#,###")
   End If
   
   If Len(txtValueCredit(Index)) > 3 Then
      txtValueCredit(Index) = Format(txtValueCredit(Index), "#,###")
   End If
End Sub

Private Sub txtValueDebit_GotFocus(Index As Integer)
   txtValueDebit(Index) = Replace(Replace(txtValueDebit(Index), ",", vbNullString), ".", vbNullString)
   txtValueDebit(Index).SelStart = 0
   txtValueDebit(Index).SelLength = Len(txtValueDebit(Index))
   
   Set MyActiveTB = txtValueDebit(Index)
End Sub

Private Sub txtValueDebit_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 45 Then
      KeyAscii = 0
      
      If txtValueDebit(Index).ForeColor = vbBlack Then
         txtValueDebit(Index).ForeColor = vbRed
      Else
         txtValueDebit(Index).ForeColor = vbBlack
      End If
      
      GoTo ExitSub
   End If
   
   If KeyAscii = 43 Then
      KeyAscii = 0
      txtValueDebit(Index).ForeColor = vbBlack
      
      GoTo ExitSub
   End If
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
      ShowBalloonTip txtValueDebit(Index).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
   
   Exit Sub
   
ExitSub:
   bFlagDb(Index) = txtValueDebit(Index).ForeColor = vbBlack
End Sub

Private Sub txtValueDebit_LostFocus(Index As Integer)
   Dim intCount As Integer
   Dim strTemp As String
   Dim dblDebit As Double
   
   On Error GoTo ErrHandler
   
   For intCount = 0 To intCtlCount
      strTemp = Replace(Replace(Trim(txtValueDebit(intCount)), ",", vbNullString), ".", vbNullString)
      
      If Not IsNumeric(strTemp) And strTemp <> vbNullString Then
         txtValueDebit(intCount).SelStart = 0
         txtValueDebit(intCount).SelLength = Len(txtValueDebit(intCount))
         ShowBalloonTip txtValueDebit(intCount).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
         Exit Sub
      End If
      
      dblDebit = dblDebit + (CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagDb(intCount), 1, -1))
   Next
   
   bFlagTotalDb = dblDebit > 0
   
   If dblDebit < 0 Then
      dblDebit = dblDebit * -1
      txtTotalD.ForeColor = vbRed
   Else
      txtTotalD.ForeColor = vbBlack
   End If
   
   txtTotalD = CStr(dblDebit)
   
   If Len(txtTotalD) > 3 Then
      txtTotalD = Format(txtTotalD, "#,###")
   End If
   
   If Len(txtValueDebit(Index)) > 3 Then
      txtValueDebit(Index) = Format(txtValueDebit(Index), "#,###")
   End If
   
   Exit Sub
   
ErrHandler:
   MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub txtVoucher_GotFocus()
   HideCaret txtVoucher.hWnd
End Sub

Private Sub VS1_Change()
   VS1.Value = (VS1.Value \ 21) * 21
   picGrabber.Top = -VS1.Value
End Sub

Private Sub VS1_Scroll()
   VS1.Value = (VS1.Value \ 21) * 21
   picGrabber.Top = -VS1.Value
End Sub

Private Sub EnsureAccountVisible(intItem As Integer)
   Dim int1stVisible As Integer
   
   If VS1.Enabled Then
      int1stVisible = (VS1.Value \ 21)
      
      If int1stVisible + 12 < intItem Then
         VS1.Value = (intItem - 12) * 21
      End If
      
      If int1stVisible <= intItem And (int1stVisible + 12) >= intItem Then
         Exit Sub
      End If
      
      If int1stVisible > intItem Then
         VS1.Value = intItem * 21
         Exit Sub
      End If
   End If
End Sub
