VERSION 5.00
Begin VB.Form frmJournal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transaction"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Journal.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   10620
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
      Height          =   705
      Index           =   3
      Left            =   7125
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   27
      Tag             =   "Date"
      Top             =   105
      Width           =   3360
      Begin Project1.ctlDatePicker ctlDatePicker1 
         Height          =   315
         Left            =   795
         TabIndex        =   9
         Top             =   285
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   29
         Top             =   330
         Width           =   345
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   5
         Left            =   690
         TabIndex        =   28
         Top             =   330
         Width           =   60
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
      Index           =   4
      Left            =   5370
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   339
      TabIndex        =   23
      Tag             =   "Memo"
      Top             =   885
      Width           =   5115
      Begin VB.TextBox txtNote 
         Height          =   315
         Left            =   75
         MaxLength       =   254
         TabIndex        =   5
         Top             =   285
         Width           =   4935
      End
   End
   Begin Project1.lvButtons_H lvB_OK 
      Height          =   420
      Left            =   4575
      TabIndex        =   6
      Top             =   3480
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
      Image           =   "Journal.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   8265
      TabIndex        =   7
      Top             =   3480
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
      Image           =   "Journal.frx":071E
      cBack           =   -2147483633
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
      Left            =   135
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   342
      TabIndex        =   15
      Tag             =   "Description"
      Top             =   885
      Width           =   5160
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   75
         MaxLength       =   254
         TabIndex        =   4
         Top             =   285
         Width           =   4980
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
      Height          =   1740
      Index           =   1
      Left            =   135
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   688
      TabIndex        =   12
      Tag             =   "Chart of Accounts"
      Top             =   1665
      Width           =   10350
      Begin VB.PictureBox picScroller2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -90
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   727
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1320
         Width           =   10905
         Begin VB.VScrollBar VS2 
            Enabled         =   0   'False
            Height          =   330
            LargeChange     =   2
            Left            =   10140
            Max             =   0
            TabIndex        =   37
            Top             =   0
            Width           =   285
         End
         Begin VB.PictureBox picGrabber2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            ScaleHeight     =   22
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   677
            TabIndex        =   38
            Top             =   0
            Width           =   10155
            Begin VB.TextBox txtCoATo 
               BackColor       =   &H00E0E0E0&
               Height          =   300
               Index           =   0
               Left            =   375
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   15
               Width           =   2865
            End
            Begin VB.TextBox txtValueTo 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   0
               Left            =   5865
               TabIndex        =   1
               Tag             =   "NUMBERONLY"
               Top             =   15
               Width           =   1650
            End
            Begin VB.TextBox txtCostumerTo 
               Height          =   300
               Index           =   0
               Left            =   7530
               MaxLength       =   254
               TabIndex        =   3
               Top             =   15
               Width           =   2475
            End
            Begin Project1.lvButtons_H lvB_CoATo 
               Height          =   240
               Index           =   0
               Left            =   5565
               TabIndex        =   40
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
               Mode            =   1
               Value           =   0   'False
               Image           =   "Journal.frx":0A70
               ImgSize         =   40
               cBack           =   -2147483633
            End
            Begin VB.TextBox txtRemTo 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   3255
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   15
               Width           =   2595
            End
            Begin VB.Label lblNumber2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   60
               Width           =   315
            End
         End
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
         ScaleWidth      =   727
         TabIndex        =   32
         Top             =   585
         Width           =   10905
         Begin VB.VScrollBar VS1 
            Enabled         =   0   'False
            Height          =   330
            LargeChange     =   2
            Left            =   10140
            Max             =   0
            TabIndex        =   33
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
            ScaleWidth      =   677
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   10155
            Begin VB.TextBox txtCoAFrom 
               BackColor       =   &H00E0E0E0&
               Height          =   300
               Index           =   0
               Left            =   375
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   15
               Width           =   2865
            End
            Begin VB.TextBox txtValueFrom 
               Alignment       =   1  'Right Justify
               Height          =   300
               Index           =   0
               Left            =   5865
               TabIndex        =   0
               Tag             =   "NUMBERONLY"
               Top             =   15
               Width           =   1650
            End
            Begin VB.TextBox txtCostumerFrom 
               Height          =   300
               Index           =   0
               Left            =   7530
               MaxLength       =   254
               TabIndex        =   2
               Top             =   15
               Width           =   2475
            End
            Begin Project1.lvButtons_H lvB_CoAFrom 
               Height          =   240
               Index           =   0
               Left            =   5565
               TabIndex        =   43
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
               Mode            =   1
               Value           =   0   'False
               Image           =   "Journal.frx":0DC2
               ImgSize         =   40
               cBack           =   -2147483633
            End
            Begin VB.TextBox txtRemFrom 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   3255
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   15
               Width           =   2595
            End
            Begin VB.Label lblNumber 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   60
               Width           =   315
            End
         End
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   0
         Left            =   5205
         TabIndex        =   16
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
         Image           =   "Journal.frx":1114
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   1
         Left            =   5205
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "Add a field."
         Top             =   1020
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
         Image           =   "Journal.frx":1466
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   2
         Left            =   4980
         TabIndex        =   18
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
         Image           =   "Journal.frx":17B8
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   270
         Index           =   3
         Left            =   4980
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "Remove the last field."
         Top             =   1020
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
         Image           =   "Journal.frx":1B0A
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customers"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   3
         Left            =   7605
         TabIndex        =   31
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblCoATo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customers"
         ForeColor       =   &H00404080&
         Height          =   195
         Index           =   3
         Left            =   7605
         TabIndex        =   30
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   2
         Left            =   3345
         TabIndex        =   25
         Top             =   345
         Width           =   615
      End
      Begin VB.Label lblCoATo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00404080&
         Height          =   195
         Index           =   2
         Left            =   3345
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCoATo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Values (Rp.)"
         ForeColor       =   &H00404080&
         Height          =   195
         Index           =   1
         Left            =   6600
         TabIndex        =   21
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblCoAFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Values (Rp.)"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   1
         Left            =   6600
         TabIndex        =   20
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lblCoAFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#   Debit Accounts"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label lblCoATo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#   Credit Accounts"
         ForeColor       =   &H00404080&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   1080
         Width           =   1395
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
      Left            =   135
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   11
      Tag             =   "Administrative"
      Top             =   105
      Width           =   3360
      Begin VB.TextBox txtVoucher 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   795
         Locked          =   -1  'True
         MaxLength       =   12
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   285
         Width           =   2460
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   26
         Top             =   330
         Width           =   60
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slip No."
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   22
         Top             =   330
         Width           =   540
      End
   End
   Begin Project1.lvButtons_H lvB_Print 
      Height          =   420
      Left            =   6435
      TabIndex        =   8
      Top             =   3480
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   741
      Caption         =   "&Print"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
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
      Image           =   "Journal.frx":1E5C
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "= Negative value"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   405
      TabIndex        =   46
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   195
      Top             =   3645
      Width           =   150
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bNewTransaction As Boolean
Public strSlipNo As String

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long

Private MyCtlCoADb As Integer
Private MyCtlCoACr As Integer

Private bFlagDb() As Boolean
Private bFlagCr() As Boolean

Private bFlagTotalDb As Boolean
Private bFlagTotalCr As Boolean

Private bOpeningBalance As Boolean
Private bOBYear As Boolean

Private Sub ctlDatePicker1_Change()
   If bNewTransaction Then
      intCurrRecord = GenerateSlipNo(Year(ctlDatePicker1.DTPValue), Month(ctlDatePicker1.DTPValue), Day(ctlDatePicker1.DTPValue), Me.hWnd)
      txtVoucher = GenerateFullSlip(Right(Year(ctlDatePicker1.DTPValue), 2), Format(Month(ctlDatePicker1.DTPValue), "00"), Format(Day(ctlDatePicker1.DTPValue), "00"))
   End If
End Sub

Private Sub Form_Load()
   Dim b1stD As Boolean
   Dim b1stC As Boolean
   Dim dtDefault As Date
   Dim intCount As Integer
   
   ReDim bFlagDb(0)
   ReDim bFlagCr(0)
   
   bFlagDb(0) = True
   bFlagCr(0) = True
   
   For intCount = 0 To 4
      DrawPicHolderCaption picHolder(intCount)
   Next intCount
   
   ctlDatePicker1.DTPValue = Date
   
   MyCtlCoADb = 0
   MyCtlCoACr = 0
   
   b1stD = True
   b1stC = True
   
   If Not bNewTransaction Then
      CloseConn , Me.hWnd
      ADORS.Open "SELECT * FROM tblJournal WHERE Voucher = '" & strSlipNo & "';", ADOCnn
      
      txtVoucher = ADORS!Voucher
      dtDefault = ADORS!Date
      txtDesc = ADORS!Description
      txtNote = IIf(IsNull(ADORS!Notes), vbNullString, ADORS!Notes)
      
      If Not IsNull(ADORS!InternalID) Then
         bOpeningBalance = ADORS!InternalID = "OB"
         bOBYear = ADORS!InternalID = "OBY"
      End If
      
      If bOpeningBalance Or bOBYear Then
         txtDesc.Locked = True
         txtDesc.BackColor = &HE0E0E0
         txtDesc.MousePointer = 1
         
         SetReadonly txtDesc.hWnd
      End If
      
      Do Until ADORS.EOF
         If ADORS!Debit <> 0 Then
            If b1stD Then
               b1stD = False
            Else
               Call lvButtons_H2_Click(0)
            End If
            
            txtCoAFrom(MyCtlCoADb) = ADORS!DebitCoA
            txtRemFrom(MyCtlCoADb) = ADORS!Remark
            txtRemFrom(MyCtlCoADb).Tag = ADORS!Category
            
            bFlagDb(MyCtlCoADb) = ADORS!Debit > 0
            
            txtValueFrom(MyCtlCoADb).ForeColor = IIf(bFlagDb(MyCtlCoADb), vbBlack, vbRed)
            txtValueFrom(MyCtlCoADb) = Format(IIf(bFlagDb(MyCtlCoADb), ADORS!Debit, ADORS!Debit * -1), "#,###")
            txtCostumerFrom(MyCtlCoADb) = IIf(IsNull(ADORS!Customer), vbNullString, ADORS!Customer)
         Else
            If b1stC Then
               b1stC = False
            Else
               Call lvButtons_H2_Click(1)
            End If
            
            txtCoATo(MyCtlCoACr) = ADORS!CreditCoa
            txtRemTo(MyCtlCoACr) = ADORS!Remark
            txtRemTo(MyCtlCoACr).Tag = ADORS!Category
            
            bFlagCr(MyCtlCoACr) = ADORS!Credit > 0
            
            txtValueTo(MyCtlCoACr).ForeColor = IIf(bFlagCr(MyCtlCoACr), vbBlack, vbRed)
            txtValueTo(MyCtlCoACr) = Format(IIf(bFlagCr(MyCtlCoACr), ADORS!Credit, ADORS!Credit * -1), "#,###")
            txtCostumerTo(MyCtlCoACr) = IIf(IsNull(ADORS!Customer), vbNullString, ADORS!Customer)
         End If
         
         ADORS.MoveNext
      Loop
      
      txtVoucher.Enabled = False
      
      ctlDatePicker1.Enabled = False
      ctlDatePicker1.DTPValue = dtDefault
   Else
      txtVoucher = GenerateFullSlip(Right(Year(ctlDatePicker1.DTPValue), 2), Format(Month(ctlDatePicker1.DTPValue), "00"), Format(Day(ctlDatePicker1.DTPValue), "00"))
   End If
   
   SetMenuForm Me.hWnd
   SetReadonly txtCoAFrom(0).hWnd
   SetReadonly txtCoATo(0).hWnd
   SetReadonly txtRemFrom(0).hWnd
   SetReadonly txtRemTo(0).hWnd
   SetReadonly txtVoucher.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set MyActiveTB = Nothing
End Sub

Private Sub lvB_Cancel_Click()
   frmMain.FillMainLV
   Unload Me
End Sub

Private Sub lvB_CoAFrom_Click(Index As Integer)
   Set MyLV_B_DD = lvB_CoAFrom(Index)
   
   lvB_CoAFrom(Index).Enabled = False
   LoadCoA txtCoAFrom(Index), txtRemFrom(Index)
End Sub

Private Sub lvB_CoAFrom_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvB_CoAFrom(Index).hWnd And _
      MyCoolTip.TipText = lvB_CoAFrom(Index).Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvB_CoAFrom(Index).hWnd
   MyCoolTip.TipText = lvB_CoAFrom(Index).Tag
   MyCoolTip.Create
End Sub

Private Sub lvB_CoAFrom_MouseOnButton(Index As Integer, OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub lvB_CoATo_Click(Index As Integer)
   Set MyLV_B_DD = lvB_CoATo(Index)

   lvB_CoATo(Index).Enabled = False
   LoadCoA txtCoATo(Index), txtRemTo(Index)
End Sub

Private Sub lvB_CoATo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvB_CoATo(Index).hWnd And _
      MyCoolTip.TipText = lvB_CoATo(Index).Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvB_CoATo(Index).hWnd
   MyCoolTip.TipText = lvB_CoATo(Index).Tag
   MyCoolTip.Create
End Sub

Private Sub lvB_OK_Click()
   Dim intLoop As Integer
   Dim dblDebit As Double
   Dim dblCredit As Double
   Dim strTemp As String
   
   On Error GoTo ErrHandler
   
   If Trim(txtVoucher) = vbNullString Then
      ShowBalloonTip txtVoucher.hWnd, "Invalid Entry Value", "Please specify correct entry value!", [2-Information Icon]
      Exit Sub
   End If
   
   For intLoop = 0 To MyCtlCoADb
      If txtCoAFrom(intLoop) = vbNullString Then
         txtRemFrom(intLoop) = String(52, " ")
         txtRemFrom(intLoop).SelStart = Len(txtRemFrom(intLoop))
         EnsureAccountVisible VS1, intLoop
         ShowBalloonTip txtRemFrom(intLoop).hWnd, "Invalid Field", "Accounts & Remarks field must be filled!" & vbCrLf & "Click here to select from the list!", [2-Information Icon]
         txtRemFrom(intLoop) = vbNullString
         Exit Sub
      End If
      
      If txtValueFrom(intLoop) = vbNullString Or txtValueFrom(intLoop) = "0" Then
         EnsureAccountVisible VS1, intLoop
         ShowBalloonTip txtValueFrom(intLoop).hWnd, "Invalid Debit Amount", "Debit value must be filled!", [2-Information Icon]
         Exit Sub
      End If
   Next intLoop

   For intLoop = 0 To MyCtlCoACr
      If txtCoATo(intLoop) = vbNullString Then
         txtRemTo(intLoop) = String(52, " ")
         txtRemTo(intLoop).SelStart = Len(txtRemTo(intLoop))
         EnsureAccountVisible VS2, intLoop
         ShowBalloonTip txtRemTo(intLoop).hWnd, "Invalid Field", "Accounts & Remarks field must be filled!" & vbCrLf & "Click here to select from the list!", [2-Information Icon]
         txtRemTo(intLoop) = vbNullString
         Exit Sub
      End If
   
      If txtValueTo(intLoop) = vbNullString Or txtValueTo(intLoop) = "0" Then
         EnsureAccountVisible VS2, intLoop
         ShowBalloonTip txtValueTo(intLoop).hWnd, "Invalid Credit Amount", "Credit value must be filled!", [2-Information Icon]
         Exit Sub
      End If
   Next intLoop
   
   For intLoop = 0 To MyCtlCoADb
      strTemp = Replace(txtValueFrom(intLoop), ",", vbNullString)
      strTemp = Replace(strTemp, ".", vbNullString)
      
      dblDebit = dblDebit + CDbl(IIf(Trim(strTemp) = vbNullString, 0, strTemp)) * IIf(bFlagDb(intLoop), 1, -1)
   Next
   
   For intLoop = 0 To MyCtlCoACr
      strTemp = Replace(txtValueTo(intLoop), ",", vbNullString)
      strTemp = Replace(strTemp, ".", vbNullString)
      
      dblCredit = dblCredit + CDbl(IIf(Trim(strTemp) = vbNullString, 0, strTemp)) * IIf(bFlagCr(intLoop), 1, -1)
   Next
   
   If dblCredit <> dblDebit Then
      EnsureAccountVisible VS1, 0
      ShowBalloonTip txtValueFrom(0).hWnd, "Invalid Debit/Credit Amount", "Debit and credit amount must be identical!", [2-Information Icon]
      Exit Sub
   End If
   
   If Trim(txtDesc) = vbNullString Then
      ShowBalloonTip txtDesc.hWnd, "Invalid Field Value", "This field must be filled!", [2-Information Icon]
      Exit Sub
   End If
   
   With ADORS
      If bNewTransaction Then
         CloseConn , Me.hWnd
         .Open "SELECT Voucher FROM tblJournal;", ADOCnn
         
         Do Until .EOF
            If UCase(Trim(txtVoucher)) = UCase(!Voucher) Then
               ShowBalloonTip txtVoucher.hWnd, "Invalid Field Value", "Slip No value has already been used, try another!", [2-Information Icon]
               Exit Sub
            End If
            
            ADORS.MoveNext
         Loop
      Else
         CloseConn , Me.hWnd
         .Open "SELECT * FROM tblJournal WHERE Voucher = '" & strSlipNo & "';", ADOCnn, adOpenDynamic, adLockOptimistic
         
         Do Until .EOF
            .Delete
            .MoveNext
         Loop
      End If

      CloseConn , Me.hWnd
      .Open "SELECT * FROM tblJournal;", ADOCnn, adOpenDynamic, adLockOptimistic
      
      For intLoop = 0 To MyCtlCoADb
         .AddNew
         !Voucher = txtVoucher
         !Date = ctlDatePicker1.DTPValue
         !DebitCoA = txtCoAFrom(intLoop)

         strTemp = Replace(txtValueFrom(intLoop), ".", vbNullString)
         strTemp = Replace(strTemp, ",", vbNullString)

         !Debit = CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagDb(intLoop), 1, -1)
         !Description = txtDesc
         !Remark = txtRemFrom(intLoop)
         !Category = txtRemFrom(intLoop).Tag
         !Customer = txtCostumerFrom(intLoop)
         !Notes = txtNote
         
         If bOpeningBalance Then !InternalID = "OB"
         If bOBYear Then !InternalID = "OBY"
      Next intLoop
      
      For intLoop = 0 To MyCtlCoACr
         .AddNew
         !Voucher = txtVoucher
         !Date = ctlDatePicker1.DTPValue
         !CreditCoa = txtCoATo(intLoop)
         
         strTemp = Replace(txtValueTo(intLoop), ".", vbNullString)
         strTemp = Replace(strTemp, ",", vbNullString)
         
         !Credit = CDbl(IIf(strTemp = vbNullString, "0", strTemp)) * IIf(bFlagCr(intLoop), 1, -1)
         !Description = txtDesc
         !Remark = txtRemTo(intLoop)
         !Category = txtRemTo(intLoop).Tag
         !Customer = txtCostumerTo(intLoop)
         !Notes = txtNote
         
         If bOpeningBalance Then !InternalID = "OB"
         If bOBYear Then !InternalID = "OBY"
      Next intLoop
      
      .Update
      .UpdateBatch
   End With
   
   intCurrRecord = intCurrRecord + 1
   Call frmMain.FillMainLV
   
   Unload Me
   Exit Sub
   
ErrHandler:
   MsgBox Err.Number & " - " & Err.Description
   MsgBox intLoop & " - " & txtCoAFrom.UBound
End Sub

Private Sub lvB_Print_Click()
   Me.Enabled = False
   
   Dim MyPrinter As New clsPrinter
   Dim strTemp As String
   Dim intCount As Integer
   Dim sngWidth As Single
   Dim strTemp2 As String
   Dim intPage As Integer
   Dim bNegative As Boolean
   
   AddReportLogo frmMain.picLogo, "Transaction Voucher", MyPrinter, qPortrait, , , , 80, , , False
   
   With MyPrinter
      .TextItem(.TextItem.Count).Top = 825
      
      frmMain.picLogo.FontSize = 10
      frmMain.picLogo.FontName = "Arial"
      
      strTemp = "<LINDENT=180>Slip Number: " & txtVoucher & "<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(GetMonthName(ctlDatePicker1.DTPMonth) & " " & ctlDatePicker1.DTPDay & GetOrdinal(ctlDatePicker1.DTPDay) & ", " & ctlDatePicker1.DTPYear) & ">" & GetMonthName(ctlDatePicker1.DTPMonth) & " " & ctlDatePicker1.DTPDay & GetOrdinal(ctlDatePicker1.DTPDay) & ", " & ctlDatePicker1.DTPYear
      
      .AddText strTemp, "Arial", 10, True
      .TextItem(.TextItem.Count).Top = 600
      
      .AddLine 180, 2060, 10650, 2060, 1
      .AddLine 180, 2090, 10650, 2090, 1, 3
      
      strTemp = "<LINDENT=180>Account<LINDENT=" & 5550 - frmMain.picLogo.TextWidth("Debit") & ">Debit<LINDENT=" & 7350 - frmMain.picLogo.TextWidth("Credit") & ">Credit<LINDENT=7650>Customer"
      .AddText strTemp, "Arial", 10, True
      .TextItem(.TextItem.Count).Top = 195
      
      frmMain.picLogo.FontBold = False
      frmMain.picLogo.FontSize = 9
      
      .AddText " ", "Arial", 10, False
      .TextItem(.TextItem.Count).Top = -60
      
      .AddLine 180, 2490, 10650, 2490, 1, 1
      
      For intCount = 0 To MyCtlCoADb
         If txtValueFrom(intCount) = vbNullString Then
            strTemp2 = " "
         Else
            strTemp2 = FormatNumber(txtValueFrom(intCount) * IIf(bFlagDb(intCount), 1, -1), 0, vbTrue, vbTrue, vbTrue)
            bNegative = (txtValueFrom(intCount) * IIf(bFlagDb(intCount), 1, -1)) < 0
            
            If bNegative Then
               strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
            End If
         End If
         
         strTemp = "<LINDENT=180>" & txtCoAFrom(intCount) & "<LINDENT=" & 5550 - frmMain.picLogo.TextWidth(strTemp2) & ">" & strTemp2 & IIf(bNegative, ")", "") & "<LINDENT=7650>" & IIf(txtCostumerFrom(intCount) = vbNullString, " ", txtCostumerFrom(intCount))
         .AddText strTemp, "Arial", 9, False
         .TextItem(.TextItem.Count).Top = 60
      Next

      For intCount = 0 To MyCtlCoACr
         If txtValueTo(intCount) = vbNullString Then
            strTemp2 = " "
         Else
            strTemp2 = FormatNumber(txtValueTo(intCount) * IIf(bFlagCr(intCount), 1, -1), 0, vbTrue, vbTrue, vbTrue)
            bNegative = (txtValueTo(intCount) * IIf(bFlagCr(intCount), 1, -1)) < 0
            
            If bNegative Then
               strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
            End If
         End If
         
         strTemp = "<LINDENT=180>    " & txtCoATo(intCount) & "<LINDENT=" & 7350 - frmMain.picLogo.TextWidth(strTemp2) & ">" & strTemp2 & IIf(bNegative, ")", "") & "<LINDENT=7650>" & IIf(txtCostumerTo(intCount) = vbNullString, " ", txtCostumerTo(intCount))
         .AddText strTemp, "Arial", 9, False
         .TextItem(.TextItem.Count).Top = 60
      Next
      
      strTemp = "<LINDENT=180><B>Description<LINDENT=1450>:</B><LINDENT=1635>" & txtDesc
      .AddText strTemp, "Arial", 10, False
      .TextItem(.TextItem.Count).Top = 450
      
      If Trim(txtNote) <> vbNullString Then
         strTemp = "<LINDENT=180><B>Memo<LINDENT=1450>:</B><LINDENT=1635>" & txtNote
         .AddText strTemp, "Arial", 10, False
         .TextItem(.TextItem.Count).Top = 60
      End If
      
      sngWidth = (MyCtlCoADb + MyCtlCoACr + 2) * 270 + 2760 '1845
      
      intPage = .Pages
      
      .AddText " ", "Arial", 32, True
      .AddText "<LINDENT=7890>Accounting<LINDENT=9570>Chief", "Arial", 9, True
      
      Dim sngTop As Single
      
      sngTop = 13200
      
      .TextItem(.TextItem.Count).Top = sngTop
      .TextItem(.TextItem.Count).AbsPage = 1
      .TextItem(.TextItem.Count).Absolute = True
      
      .AddLine 7680, sngTop - 60, 10500, sngTop - 60, 1, 3
      .AddLine 7680, sngTop - 90, 10500, sngTop - 90, 1, 1
      .AddLine 7680, sngTop + 285, 10500, sngTop + 285, 1, 3
      
      .AddLine 9090, sngTop - 60, 9090, sngTop + 1200, 1, 1
      .AddLine 7680, sngTop + 1200, 10500, sngTop + 1200, 1
      .AddLine 7680, sngTop + 1170, 10500, sngTop + 1170, 1, 3
      
      .AddLine 3750, 2060, 3750, sngWidth, 1, 1
      .AddLine 5700, 2060, 5700, sngWidth, 1, 1
      .AddLine 7500, 2060, 7500, sngWidth, 1, 1
      .AddLine 180, sngWidth, 10650, sngWidth, 1, 1
      .AddLine 180, sngWidth - 30, 10650, sngWidth - 30, 1, 3
      
      .Preview Me, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
   End With
   
   Set MyPrinter = Nothing
   
   Me.Enabled = True
End Sub

Private Sub txtCoAFrom_GotFocus(Index As Integer)
   HideCaret txtCoAFrom(Index).hWnd
End Sub

Private Sub txtCoATo_GotFocus(Index As Integer)
   HideCaret txtCoATo(Index).hWnd
End Sub

Private Sub txtRemFrom_GotFocus(Index As Integer)
   HideCaret txtRemFrom(Index).hWnd
End Sub

Private Sub txtRemTo_GotFocus(Index As Integer)
   HideCaret txtRemFrom(Index).hWnd
End Sub

Private Sub VS1_Change()
   VS1.Value = (VS1.Value \ 21) * 21
   picGrabber.Top = -VS1.Value
End Sub

Private Sub VS1_Scroll()
   VS1.Value = (VS1.Value \ 21) * 21
   picGrabber.Top = -VS1.Value
End Sub

Private Sub VS2_Change()
   VS2.Value = (VS2.Value \ 21) * 21
   picGrabber2.Top = -VS2.Value
End Sub

Private Sub VS2_Scroll()
   VS2.Value = (VS2.Value \ 21) * 21
   picGrabber2.Top = -VS2.Value
End Sub

Private Sub lvButtons_H2_Click(Index As Integer)
   Dim intIndex As Integer

   If Index = 0 Then
      MyCtlCoADb = MyCtlCoADb + 1
            
      ReDim Preserve bFlagDb(MyCtlCoADb)
      
      bFlagDb(MyCtlCoADb) = True
      
      Load txtCoAFrom(MyCtlCoADb)
      txtCoAFrom(MyCtlCoADb).Visible = True
      txtCoAFrom(MyCtlCoADb) = vbNullString
      txtCoAFrom(MyCtlCoADb).Top = txtCoAFrom(MyCtlCoADb - 1).Top + 21
      
      Load lvB_CoAFrom(MyCtlCoADb)
      lvB_CoAFrom(MyCtlCoADb).Visible = True
      lvB_CoAFrom(MyCtlCoADb).ZOrder
      lvB_CoAFrom(MyCtlCoADb).Top = lvB_CoAFrom(MyCtlCoADb - 1).Top + 21
      
      Load txtValueFrom(MyCtlCoADb)
      txtValueFrom(MyCtlCoADb).ForeColor = vbBlack
      txtValueFrom(MyCtlCoADb).Visible = True
      txtValueFrom(MyCtlCoADb) = vbNullString
      txtValueFrom(MyCtlCoADb).TabIndex = txtValueFrom(MyCtlCoADb - 1).TabIndex + 1
      txtValueFrom(MyCtlCoADb).Top = txtValueFrom(MyCtlCoADb - 1).Top + 21
      
      Load txtRemFrom(MyCtlCoADb)
      txtRemFrom(MyCtlCoADb).Visible = True
      txtRemFrom(MyCtlCoADb) = vbNullString
      txtRemFrom(MyCtlCoADb).Top = txtRemFrom(MyCtlCoADb - 1).Top + 21
      
      Load txtCostumerFrom(MyCtlCoADb)
      txtCostumerFrom(MyCtlCoADb).Visible = True
      txtCostumerFrom(MyCtlCoADb) = vbNullString
      txtCostumerFrom(MyCtlCoADb).TabIndex = txtCostumerFrom(MyCtlCoADb - 1).TabIndex + 1
      txtCostumerFrom(MyCtlCoADb).Top = txtCostumerFrom(MyCtlCoADb - 1).Top + 21
      
      Load lblNumber(MyCtlCoADb)
      lblNumber(MyCtlCoADb).Visible = True
      lblNumber(MyCtlCoADb).Top = lblNumber(MyCtlCoADb - 1).Top + 21
      lblNumber(MyCtlCoADb) = MyCtlCoADb + 1
      
      SetReadonly txtCoAFrom(MyCtlCoADb).hWnd
      SetReadonly txtRemFrom(MyCtlCoADb).hWnd
      
      picGrabber.Height = picGrabber.Height + 21
      
      If MyCtlCoADb < 8 Then
         Label1.Top = Label1.Top + 21 * Screen.TwipsPerPixelY
         Shape1.Top = Shape1.Top + 21 * Screen.TwipsPerPixelY
         
         picHolder(1).Height = picHolder(1).Height + 21 * Screen.TwipsPerPixelY
         picScroller.Height = picScroller.Height + 21
         
         For intIndex = 0 To 3
            lblCoATo(intIndex).Top = lblCoATo(intIndex).Top + 21
         Next intIndex
   
         lvButtons_H2(1).Top = lvButtons_H2(1).Top + 21
         lvButtons_H2(3).Top = lvButtons_H2(3).Top + 21
         picScroller2.Top = picScroller2.Top + 21
      
         Me.Height = Me.Height + 21 * Screen.TwipsPerPixelY
         
         lvB_OK.Top = lvB_OK.Top + 21 * Screen.TwipsPerPixelY
         lvB_Print.Top = lvB_Print.Top + 21 * Screen.TwipsPerPixelY
         lvB_Cancel.Top = lvB_Cancel.Top + 21 * Screen.TwipsPerPixelY
         
         VS1.Height = VS1.Height + 21
      End If
   ElseIf Index = 1 Then
      MyCtlCoACr = MyCtlCoACr + 1
      
      ReDim Preserve bFlagCr(MyCtlCoACr)
      
      bFlagCr(MyCtlCoACr) = True
      
      Load txtCoATo(MyCtlCoACr)
      txtCoATo(MyCtlCoACr).Visible = True
      txtCoATo(MyCtlCoACr) = vbNullString
      txtCoATo(MyCtlCoACr).Top = txtCoATo(MyCtlCoACr - 1).Top + 21
      
      Load txtValueTo(MyCtlCoACr)
      txtValueTo(MyCtlCoACr).ForeColor = vbBlack
      txtValueTo(MyCtlCoACr).Visible = True
      txtValueTo(MyCtlCoACr) = vbNullString
      txtValueTo(MyCtlCoACr).TabIndex = txtValueTo(MyCtlCoACr - 1).TabIndex + 1
      txtValueTo(MyCtlCoACr).Top = txtValueTo(MyCtlCoACr - 1).Top + 21
      
      Load txtRemTo(MyCtlCoACr)
      txtRemTo(MyCtlCoACr).Visible = True
      txtRemTo(MyCtlCoACr) = vbNullString
      txtRemTo(MyCtlCoACr).Top = txtRemTo(MyCtlCoACr - 1).Top + 21
      
      Load txtCostumerTo(MyCtlCoACr)
      txtCostumerTo(MyCtlCoACr).Visible = True
      txtCostumerTo(MyCtlCoACr) = vbNullString
      txtCostumerTo(MyCtlCoACr).TabIndex = txtCostumerTo(MyCtlCoACr - 1).TabIndex + 1
      txtCostumerTo(MyCtlCoACr).Top = txtCostumerTo(MyCtlCoACr - 1).Top + 21
      
      Load lvB_CoATo(MyCtlCoACr)
      lvB_CoATo(MyCtlCoACr).Visible = True
      lvB_CoATo(MyCtlCoACr).ZOrder
      lvB_CoATo(MyCtlCoACr).Top = lvB_CoATo(MyCtlCoACr - 1).Top + 21
      
      Load lblNumber2(MyCtlCoACr)
      lblNumber2(MyCtlCoACr).Visible = True
      lblNumber2(MyCtlCoACr).Top = lblNumber2(MyCtlCoACr - 1).Top + 21
      lblNumber2(MyCtlCoACr) = MyCtlCoACr + 1
      
      SetReadonly txtCoATo(MyCtlCoACr).hWnd
      SetReadonly txtRemTo(MyCtlCoACr).hWnd
      
      picGrabber2.Height = picGrabber2.Height + 21
      
      If MyCtlCoACr < 8 Then
         Label1.Top = Label1.Top + 21 * Screen.TwipsPerPixelY
         Shape1.Top = Shape1.Top + 21 * Screen.TwipsPerPixelY
         
         Me.Height = Me.Height + 21 * Screen.TwipsPerPixelY
      
         picScroller2.Height = picScroller2.Height + 21
         picHolder(1).Height = picHolder(1).Height + 21 * Screen.TwipsPerPixelY
         
         lvB_OK.Top = lvB_OK.Top + 21 * Screen.TwipsPerPixelY
         lvB_Print.Top = lvB_Print.Top + 21 * Screen.TwipsPerPixelY
         lvB_Cancel.Top = lvB_Cancel.Top + 21 * Screen.TwipsPerPixelY
         
         VS2.Height = VS2.Height + 21
      End If
   ElseIf Index = 2 Then
      If MyCtlCoADb = 0 Then Exit Sub

      SetRewrite txtCoAFrom(MyCtlCoADb).hWnd
      SetRewrite txtRemFrom(MyCtlCoADb).hWnd
      
      Unload txtCoAFrom(MyCtlCoADb)
      Unload txtValueFrom(MyCtlCoADb)
      Unload txtRemFrom(MyCtlCoADb)
      Unload txtCostumerFrom(MyCtlCoADb)
      Unload lvB_CoAFrom(MyCtlCoADb)
      Unload lblNumber(MyCtlCoADb)

      picGrabber.Height = picGrabber.Height - 21
      
      If MyCtlCoADb <= 8 Then VS1.Enabled = False
      
      If MyCtlCoADb < 8 Then
         Label1.Top = Label1.Top - 21 * Screen.TwipsPerPixelY
         Shape1.Top = Shape1.Top - 21 * Screen.TwipsPerPixelY
         
         picHolder(1).Height = picHolder(1).Height - 21 * Screen.TwipsPerPixelY
         picScroller.Height = picScroller.Height - 21

         Me.Height = Me.Height - 21 * Screen.TwipsPerPixelY
         
         lvButtons_H2(1).Top = lvButtons_H2(1).Top - 21
         lvButtons_H2(3).Top = lvButtons_H2(3).Top - 21
   
         lvB_OK.Top = lvB_OK.Top - 21 * Screen.TwipsPerPixelY
         lvB_Print.Top = lvB_Print.Top - 21 * Screen.TwipsPerPixelY
         lvB_Cancel.Top = lvB_Cancel.Top - 21 * Screen.TwipsPerPixelY
   
         For intIndex = 0 To 3
            lblCoATo(intIndex).Top = lblCoATo(intIndex).Top - 21
         Next intIndex
         
         picScroller2.Top = picScroller2.Top - 21
         VS1.Height = VS1.Height - 21
      End If
      
      MyCtlCoADb = MyCtlCoADb - 1
   ElseIf Index = 3 Then
      If MyCtlCoACr = 0 Then Exit Sub

      SetRewrite txtCoATo(MyCtlCoACr).hWnd
      SetRewrite txtRemTo(MyCtlCoACr).hWnd
   
      Unload txtCoATo(MyCtlCoACr)
      Unload txtValueTo(MyCtlCoACr)
      Unload txtRemTo(MyCtlCoACr)
      Unload txtCostumerTo(MyCtlCoACr)
      Unload lvB_CoATo(MyCtlCoACr)
      Unload lblNumber2(MyCtlCoACr)

      picGrabber2.Height = picGrabber2.Height - 21

      If MyCtlCoACr <= 8 Then VS2.Enabled = False
      
      If MyCtlCoACr < 8 Then
         Label1.Top = Label1.Top - 21 * Screen.TwipsPerPixelY
         Shape1.Top = Shape1.Top - 21 * Screen.TwipsPerPixelY
         
         Me.Height = Me.Height - 21 * Screen.TwipsPerPixelY
         picHolder(1).Height = picHolder(1).Height - 21 * Screen.TwipsPerPixelY
         
         lvB_OK.Top = lvB_OK.Top - 21 * Screen.TwipsPerPixelY
         lvB_Print.Top = lvB_Print.Top - 21 * Screen.TwipsPerPixelY
         lvB_Cancel.Top = lvB_Cancel.Top - 21 * Screen.TwipsPerPixelY
         
         picScroller2.Height = picScroller2.Height - 21
         VS2.Height = VS2.Height - 21
      End If
      
      MyCtlCoACr = MyCtlCoACr - 1
   End If
   
   If MyCtlCoADb >= 8 Then
      VS1.Enabled = True
      VS1.Min = 0
      VS1.Max = picGrabber.Height - VS1.Height
      VS1.LargeChange = IIf((VS1.Max / 21) > 1, 42, 21)
      VS1.SmallChange = 21
   Else
      picGrabber.Top = 0
   End If
   
   If MyCtlCoACr >= 8 Then
      VS2.Enabled = True
      VS2.Min = 0
      VS2.Max = picGrabber2.Height - VS2.Height
      VS2.LargeChange = IIf((VS2.Max / 21) > 1, 42, 21)
      VS2.SmallChange = 21
   Else
      picGrabber2.Top = 0
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

Private Sub txtCoAFrom_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call lvB_CoAFrom_Click(Index)
   Else
      ShowBalloonTip txtCoAFrom(Index).hWnd, "Invalid Input Methode", "To avoid misspelling this field can only be" & vbCrLf & "filled by chose an item from CoA list!", [2-Information Icon]
   End If
   
   KeyAscii = 0
End Sub

Private Sub txtCoATo_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call lvB_CoATo_Click(Index)
   Else
      ShowBalloonTip txtCoATo(Index).hWnd, "Invalid Input Methode", "To avoid misspelling this field can only be" & vbCrLf & "filled by chose an item from CoA list!", [2-Information Icon]
   End If
   
   KeyAscii = 0
End Sub

Private Sub txtDesc_GotFocus()
   If bOpeningBalance Or bOBYear Then HideCaret txtDesc.hWnd
End Sub

Private Sub txtRemFrom_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call lvB_CoAFrom_Click(Index)
   Else
      ShowBalloonTip txtRemFrom(Index).hWnd, "Invalid Input Methode", "To avoid misspelling this field can only be" & vbCrLf & "filled by chose an item from CoA list!", [2-Information Icon]
   End If
   
   KeyAscii = 0
End Sub

Private Sub txtRemTo_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call lvB_CoATo_Click(Index)
   Else
      ShowBalloonTip txtRemTo(Index).hWnd, "Invalid Input Methode", "To avoid misspelling this field can only be" & vbCrLf & "filled by chose an item from CoA list!", [2-Information Icon]
   End If
   
   KeyAscii = 0
End Sub

Private Sub txtValueFrom_GotFocus(Index As Integer)
   txtValueFrom(Index) = Replace(txtValueFrom(Index), ".", vbNullString)
   txtValueFrom(Index) = Replace(txtValueFrom(Index), ",", vbNullString)

   txtValueFrom(Index).SelStart = 0
   txtValueFrom(Index).SelLength = Len(txtValueFrom(Index))
   
   Set MyActiveTB = txtValueFrom(Index)
End Sub

Private Sub txtValueFrom_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 45 Then
      KeyAscii = 0
      
      If txtValueFrom(Index).ForeColor = vbBlack Then
         txtValueFrom(Index).ForeColor = vbRed
      Else
         txtValueFrom(Index).ForeColor = vbBlack
      End If
      
      GoTo ExitSub
   End If
   
   If KeyAscii = 43 Then
      KeyAscii = 0
      txtValueFrom(Index).ForeColor = vbBlack
      GoTo ExitSub
   End If
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
      ShowBalloonTip txtValueFrom(Index).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
   
   Exit Sub
   
ExitSub:
   bFlagDb(Index) = txtValueFrom(Index).ForeColor = vbBlack
End Sub

Private Sub txtValueFrom_LostFocus(Index As Integer)
   If Len(txtValueFrom(Index)) > 3 Then
      txtValueFrom(Index) = Format(txtValueFrom(Index), "#,###")
   End If
End Sub

Private Sub txtValueTo_GotFocus(Index As Integer)
   txtValueTo(Index) = Replace(txtValueTo(Index), ".", vbNullString)
   txtValueTo(Index) = Replace(txtValueTo(Index), ",", vbNullString)

   txtValueTo(Index).SelStart = 0
   txtValueTo(Index).SelLength = Len(txtValueTo(Index))
   
   Set MyActiveTB = txtValueTo(Index)
End Sub

Private Sub txtValueTo_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 45 Then
      KeyAscii = 0
      
      If txtValueTo(Index).ForeColor = vbBlack Then
         txtValueTo(Index).ForeColor = vbRed
      Else
         txtValueTo(Index).ForeColor = vbBlack
      End If
      
      GoTo ExitSub
   End If
   
   If KeyAscii = 43 Then
      KeyAscii = 0
      txtValueTo(Index).ForeColor = vbBlack
      GoTo ExitSub
   End If
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
      ShowBalloonTip txtValueTo(Index).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
   
   Exit Sub
   
ExitSub:
   bFlagCr(Index) = txtValueTo(Index).ForeColor = vbBlack
End Sub

Private Sub txtValueTo_LostFocus(Index As Integer)
   If Len(txtValueTo(Index)) > 3 Then
      txtValueTo(Index) = Format(txtValueTo(Index), "#,###")
   End If
End Sub

Private Sub LoadCoA(MyTBx As TextBox, MyTBy As TextBox)
   Set MyTBDesc = MyTBx
   Set MyTBRem = MyTBy
   
   GetWindowRect MyTBx.hWnd, myRect
   
   Set frmCoA.frmParent = Me
   frmCoA.Show vbModeless, Me
End Sub

Private Sub EnsureAccountVisible(VSBar As VScrollBar, intItem As Integer)
   Dim int1stVisible As Integer
   
   If VSBar.Enabled Then
      int1stVisible = Round(VSBar.Value / 21)
      
      If int1stVisible <= intItem And (int1stVisible + 7) >= intItem Then
         Exit Sub
      End If
      
      If int1stVisible + 7 < intItem Then
         VSBar.Value = (intItem - 7) * 21
      End If
      
      If int1stVisible > intItem Then
         VSBar.Value = intItem * 21
         Exit Sub
      End If
   End If
End Sub
