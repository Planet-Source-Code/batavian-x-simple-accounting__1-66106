VERSION 5.00
Begin VB.Form frmNewRemark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Remark"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   Icon            =   "frmNewRemark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRemark 
      Height          =   345
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5040
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   3405
      TabIndex        =   1
      Top             =   645
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
      Image           =   "frmNewRemark.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Default         =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   2025
      TabIndex        =   2
      Top             =   645
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
      Image           =   "frmNewRemark.frx":035E
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmNewRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strCategory As String

Private Sub Form_Load()
   Me.Caption = " New Remark - Category = " & strCategory
   SetMenuForm Me.hWnd
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
   If Index = 0 Then
      Unload Me
      Exit Sub
   End If
   
   frmNewCoA.cboRemark.AddItem Trim(StrConv(txtRemark, vbProperCase))
   frmNewCoA.cboRemark = Trim(StrConv(txtRemark, vbProperCase))
   
   Unload Me
   
'   CloseConn , Me.hWnd
'   ADORS.Open "SELECT DISTINCT Remark FROM tblCoA;", ADOCnn, adOpenDynamic, adLockOptimistic
'
'   Do Until ADORS.EOF
'      If ADORS!Remark = Trim(StrConv(txtRemark, vbProperCase)) Then
'         ShowBalloonTip txtRemark.hWnd, "Invalid Value", "This field pointed to an existed Remark item, try another!", [2-Information Icon]
'         Exit Sub
'      End If
'
'      ADORS.MoveNext
'   Loop
'
'   ADORS.AddNew
'   ADORS!Remark = Trim(StrConv(txtRemark, vbProperCase))
'   ADORS.Update
'   ADORS.UpdateBatch
End Sub
