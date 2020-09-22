VERSION 5.00
Begin VB.Form frmByDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View By Date"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4860
   Icon            =   "frmByDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.ctlDatePicker DateTo 
      Height          =   315
      Left            =   2505
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
   End
   Begin Project1.ctlDatePicker DateFrom 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   3375
      TabIndex        =   4
      Top             =   990
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
      Image           =   "frmByDate.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   2220
      TabIndex        =   5
      Top             =   990
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   714
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
      Image           =   "frmByDate.frx":035E
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date From :"
      Height          =   195
      Index           =   1
      Left            =   255
      TabIndex        =   3
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date To :"
      Height          =   195
      Index           =   2
      Left            =   2580
      TabIndex        =   2
      Top             =   195
      Width           =   675
   End
End
Attribute VB_Name = "frmByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strDate As String

Private Sub DateFrom_Change()
   strDate = "BETWEEN #" & DateFrom.DTPMonth & "/" & DateFrom.DTPDay & "/" & DateFrom.DTPYear & "# AND #" & DateTo.DTPMonth & "/" & DateTo.DTPDay & "/" & DateTo.DTPYear & "#"
   
   If DateFrom.DTPValue > DateTo.DTPValue Then
      DateTo.DTPValue = DateFrom.DTPValue
   End If
End Sub

Private Sub DateTo_Change()
   strDate = "BETWEEN #" & DateFrom.DTPMonth & "/" & DateFrom.DTPDay & "/" & DateFrom.DTPYear & "# AND #" & DateTo.DTPMonth & "/" & DateTo.DTPDay & "/" & DateTo.DTPYear & "#"
   
   If DateFrom.DTPValue > DateTo.DTPValue Then
      DateFrom.DTPValue = DateTo.DTPValue
   End If
End Sub

Private Sub Form_Load()
   DateTo.DTPValue = DateSerial(Year(Date), Month(Date), Day(Date))
   DateFrom.DTPValue = DateSerial(Year(Date), Month(Date), 1)
   SetMenuForm Me.hWnd
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_OK_Click()
   frmMain.strDate = strDate
   frmMain.FillMainLV
   Unload Me
End Sub
