VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Dialog"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPrinter 
      Caption         =   " Printer "
      Height          =   735
      Left            =   188
      TabIndex        =   11
      Top             =   150
      Width           =   5085
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   4815
      End
   End
   Begin VB.Frame fraCopies 
      Caption         =   " Copies "
      Height          =   1215
      Left            =   188
      TabIndex        =   9
      Top             =   2550
      Width           =   5085
      Begin VB.CheckBox chkCollate 
         Appearance      =   0  'Flat
         Caption         =   "Colate"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3975
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.PictureBox picCopies 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2445
         TabIndex        =   10
         Top             =   240
         Width           =   2445
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3990
         TabIndex        =   13
         Top             =   330
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         Alignment       =   0
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196614
         OrigRight       =   255
         OrigBottom      =   750
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3975
         TabIndex        =   6
         Tag             =   "NumberOnly"
         Text            =   "1"
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   " Page "
      Height          =   1455
      Left            =   188
      TabIndex        =   8
      Top             =   990
      Width           =   5085
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   105
         ScaleHeight     =   1155
         ScaleWidth      =   4905
         TabIndex        =   12
         Top             =   240
         Width           =   4905
         Begin VB.TextBox txtStart 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            TabIndex        =   4
            Tag             =   "NumberOnly"
            Top             =   405
            Width           =   1110
         End
         Begin VB.TextBox txtEnd 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3720
            TabIndex        =   5
            Tag             =   "NumberOnly"
            Top             =   780
            Width           =   1110
         End
         Begin VB.OptionButton optPrint 
            Appearance      =   0  'Flat
            Caption         =   "Current page"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   30
            TabIndex        =   1
            Top             =   15
            Width           =   1395
         End
         Begin VB.OptionButton optPrint 
            Appearance      =   0  'Flat
            Caption         =   "All pages"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   30
            TabIndex        =   2
            Top             =   375
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton optPrint 
            Appearance      =   0  'Flat
            Caption         =   "Page ..."
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   30
            TabIndex        =   3
            Top             =   735
            Width           =   945
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            X1              =   1035
            X2              =   2040
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Label lblPages 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3735
            TabIndex        =   20
            Top             =   105
            Width           =   345
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            Caption         =   "Page"
            Height          =   195
            Left            =   2220
            TabIndex        =   19
            Top             =   90
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   195
            Left            =   3615
            TabIndex        =   16
            Top             =   825
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   195
            Left            =   3615
            TabIndex        =   15
            Top             =   450
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   195
            Left            =   3615
            TabIndex        =   14
            Top             =   90
            Width           =   45
         End
         Begin VB.Label lblStart 
            Caption         =   "Page from"
            Height          =   255
            Left            =   2220
            TabIndex        =   18
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lblEnd 
            Caption         =   "Page to"
            Height          =   255
            Left            =   2220
            TabIndex        =   17
            Top             =   825
            Width           =   1350
         End
      End
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   3488
      TabIndex        =   21
      Top             =   3945
      Width           =   1785
      _ExtentX        =   3149
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
      Image           =   "frmPrint.frx":27A2
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Print 
      Default         =   -1  'True
      Height          =   420
      Left            =   2115
      TabIndex        =   22
      Top             =   3945
      Width           =   1770
      _ExtentX        =   3122
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
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmPrint.frx":2AF4
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarCurrent As Integer
Private mvarMax As Integer
Private mvarStart As Integer
Private mvarEnd As Integer
Private mvarPrint As Boolean
Private mvarCollate As Boolean
Private mvarFlags As qePrintOptionFlags
Private mvarPrinter As Integer
Private mvarCopies As Integer
Private bInternal As Boolean

Friend Property Let Flags(ByVal eFlag As qePrintOptionFlags)
   mvarFlags = eFlag
End Property

Private Sub cboPrinter_Click()
   If Not bInternal And cboPrinter.ListIndex > -1 Then
      mvarPrinter = cboPrinter.ItemData(cboPrinter.ListIndex)
   End If
End Sub

Private Sub chkCollate_Click()
   mvarCollate = CBool(chkCollate.Value = vbChecked)
   Copies_ShowImage
End Sub

Private Sub Copies_ShowImage()
   Dim sX As Single
   Dim sXX As Single
   Dim i As Integer
   Dim iPage As Integer

   picCopies.Cls
   picCopies.FontSize = 7
   iPage = 1
   sXX = 90
   
   If chkCollate.Value Then
      For i = 1 To 2
         For sX = sXX To sXX + 750 Step 330
            picCopies.Line (sX, 210)-(sX + 300, 635), vbWhite, BF
            picCopies.Line (sX, 210)-(sX + 300, 635), vbBlack, B
            picCopies.CurrentX = sX + 240 - picCopies.TextWidth(iPage)
            picCopies.CurrentY = 460
            picCopies.Print iPage
            iPage = iPage + 1
         Next sX
         
         sXX = 1170
         iPage = 1
      Next i
   Else
      For i = 1 To 3
         For sX = sXX To sXX + 420 Step 330
            picCopies.Line (sX, 210)-(sX + 300, 635), vbWhite, BF
            picCopies.Line (sX, 210)-(sX + 300, 635), vbBlack, B
            picCopies.CurrentX = sX + 240 - picCopies.TextWidth(i)
            picCopies.CurrentY = 460
            picCopies.Print i
         Next sX
         
         sXX = sXX + 750
      Next i
   End If

End Sub

Private Sub lvB_Cancel_Click()
   mvarPrint = False
   Me.Hide
End Sub

Private Sub lvB_Print_Click()
   Me.Enabled = False
   
   Dim lStart As Integer
   Dim lEnd As Integer
   Dim bEnable As Boolean

   bEnable = True
   lStart = Val(txtStart.Text)
   lEnd = Val(txtEnd.Text)
   
   If lStart = 0 Or lEnd = 0 Then bEnable = False
   If lStart > lEnd Then bEnable = False
   If lStart <> CInt(lStart) Then bEnable = False
   If lEnd <> CInt(lEnd) Then bEnable = False

   If optPrint(0).Value Then
      bEnable = True
      mvarStart = mvarCurrent
      mvarEnd = mvarCurrent
   ElseIf optPrint(1).Value Then
      bEnable = True
      mvarStart = 1
      mvarEnd = mvarMax
   ElseIf optPrint(2).Value Then
      mvarStart = lStart
      mvarEnd = lEnd
   End If

   mvarCopies = Val(txtCopies.Text)

   If Not bEnable Then
      xMsgBox Me.hWnd, "Please type in a correct value in page from field. ", vbCritical, "Invalid Value"
      mvarPrint = False
   Else
      mvarPrint = True
      Me.Hide
   End If
End Sub

Public Property Get PrintDoc() As Boolean
   PrintDoc = mvarPrint
End Property

Public Property Let PageCurrent(ByVal vNewValue As Integer)
   mvarCurrent = vNewValue
End Property

Public Property Get PageStart() As Integer
   PageStart = mvarStart
End Property

Public Property Get PageEnd() As Integer
   PageEnd = mvarEnd
End Property

Public Property Get Collate() As Boolean
   Collate = mvarCollate
End Property

Public Property Get Copies() As Integer
   Copies = mvarCopies
End Property

Public Property Get PrinterNumber() As Integer
   PrinterNumber = mvarPrinter
End Property

Public Property Let PageMax(ByVal vNewValue As Integer)
   mvarMax = vNewValue
End Property

Private Sub Form_Load()
   Dim ctlAll As Control
   Dim prtPrinter As Printer
   Dim iPrinter As Integer

   For Each ctlAll In Me
      If TypeOf ctlAll Is TextBox Then
         HookContext ctlAll.hWnd
      End If
   Next ctlAll
   
   bInternal = True
   
   If CBool(mvarFlags And ShowPrinter_po) Then
      fraPrinter.Visible = True
      cboPrinter.Clear
      iPrinter = 0

      For Each prtPrinter In Printers
         cboPrinter.AddItem prtPrinter.DeviceName 'sPrinter
         cboPrinter.ItemData(cboPrinter.NewIndex) = iPrinter
         
         If prtPrinter.DeviceName = Printer.DeviceName And Printer.Port = prtPrinter.Port Then
            cboPrinter.ListIndex = cboPrinter.NewIndex
            mvarPrinter = iPrinter
         End If
         
         iPrinter = iPrinter + 1
      Next
   Else
      fraPrinter.Enabled = False
   End If

   optPrint(0).Value = True
   optPrint(1).Value = CBool(mvarMax > 1)
   optPrint(1).Enabled = CBool(mvarMax > 1)
   optPrint(2).Enabled = CBool(mvarMax > 1)
   lblPages.Caption = mvarCurrent & " / " & mvarMax

'   txtStart.Enabled = CBool(mvarMax > 1)
'   txtEnd.Enabled = CBool(mvarMax > 1)
   
   txtStart.Text = "1"
   txtEnd.Text = mvarMax

   If CBool(mvarFlags And ShowCopies_po) Then
      fraCopies.Visible = True
      mvarCollate = True
      Copies_ShowImage
   Else
      fraCopies.Visible = False
   End If

   bInternal = False
'   SetCBListStyle cboPrinter
   SetMenuForm Me.hWnd
End Sub

Private Sub optPrint_Click(Index As Integer)
   txtStart.Enabled = CBool(Index = 2)
   txtEnd.Enabled = CBool(Index = 2)

   If Index = 0 Then
      txtStart.Text = mvarCurrent
      txtEnd.Text = mvarCurrent
   Else
      txtStart.Text = 1
      txtEnd.Text = mvarMax
   End If
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      ShowBalloonTip txtCopies.hWnd, "Kesalahan", "Kotak ini hanya dapat diisi dengan angka!", [4-Critical Icon]
   End If
End Sub

Private Sub txtEnd_GotFocus()
   optPrint(2).Value = True
'   Set ActiveTB = txtEnd
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      ShowBalloonTip txtEnd.hWnd, "Kesalahan", "Kotak ini hanya dapat diisi dengan angka!", [4-Critical Icon]
   End If
End Sub

Private Sub txtStart_GotFocus()
   optPrint(2).Value = True
'   Set ActiveTB = txtStart
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      ShowBalloonTip txtStart.hWnd, "Kesalahan", "Kotak ini hanya dapat diisi dengan angka!", [4-Critical Icon]
   End If
End Sub
