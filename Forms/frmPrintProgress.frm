VERSION 5.00
Begin VB.Form frmPrintProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   Moveable        =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   1860
      Picture         =   "frmPrintProgress.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   390
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Printing, please wait ...."
      Height          =   900
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404000&
      BorderWidth     =   4
      Height          =   2535
      Left            =   15
      Top             =   15
      Width           =   4665
   End
End
Attribute VB_Name = "frmPrintProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
