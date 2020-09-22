VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Calculator"
   ClientHeight    =   2325
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCalc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   555
      Width           =   3165
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   1
      Left            =   2175
      TabIndex        =   14
      Top             =   915
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      Caption         =   "/"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvBCalc 
      Height          =   315
      Left            =   615
      TabIndex        =   13
      Top             =   1950
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "Calc."
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   2
      Mode            =   1
      Value           =   -1  'True
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   1950
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "0"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtCalc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   210
      Width           =   1875
   End
   Begin VB.TextBox txtCalc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Text            =   "0"
      Top             =   210
      Width           =   1290
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   1605
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "1"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1605
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "2"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   3
      Left            =   1290
      TabIndex        =   5
      Top             =   1605
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "3"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "4"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   5
      Left            =   600
      TabIndex        =   7
      Top             =   1260
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "5"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   6
      Left            =   1290
      TabIndex        =   8
      Top             =   1260
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "6"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   7
      Left            =   60
      TabIndex        =   9
      Top             =   915
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "7"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   8
      Left            =   600
      TabIndex        =   10
      Top             =   915
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "8"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Number 
      Height          =   315
      Index           =   9
      Left            =   1290
      TabIndex        =   11
      Top             =   915
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      Caption         =   "9"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   12582912
      cFHover         =   12582912
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Fill 
      Height          =   315
      Left            =   1410
      TabIndex        =   12
      Top             =   1950
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      Caption         =   "Fill Active Field"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   0
      Left            =   2550
      TabIndex        =   15
      Top             =   915
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "*"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   2
      Left            =   2175
      TabIndex        =   16
      Top             =   1260
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      Caption         =   "-"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   3
      Left            =   2175
      TabIndex        =   17
      Top             =   1605
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      Caption         =   "+"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   4
      Left            =   2535
      TabIndex        =   18
      Top             =   1605
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "="
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Operator 
      Height          =   315
      Index           =   5
      Left            =   2550
      TabIndex        =   19
      Top             =   1260
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "C"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2

Private Enum eOperator
   [0-Multiply] = 0
   [1-Divide] = 1
   [2-Subtract] = 2
   [3-Adds] = 3
   [4-Previous] = 4
   [5-Empty] = 5
End Enum

Private MyPrevOp As eOperator
Private MyOperator As eOperator
Private dblTemp As Double
Private dblRef As Double
Private sngTop As Single
Private sngLeft As Single

Private Sub Form_Load()
   Dim bCalc As Boolean
   
   bCalc = GetSetting("Batavian's Accounting Program", "Setting", "CalcMode", "0") = "1"
   txtCalc(0) = GetSetting("Batavian's Accounting Program", "Setting", "Rate", "8500")
   
   Me.Left = CSng(GetSetting("Batavian's Accounting Program", "Setting", "CalcLeft", "150"))
   Me.Top = CSng(GetSetting("Batavian's Accounting Program", "Setting", "CalcTop", "150"))
   
   If txtCalc(0) <> "0" Then txtCalc(1).TabIndex = 0
   If bCalc Then lvBCalc.Value = False
   
   Call lvBCalc_Click
   MyOperator = [5-Empty]
   MyPrevOp = [5-Empty]
   SetMenuForm Me.hWnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
   End If
End Sub

Private Sub Form_Paint()
   Me.CurrentX = 6: Me.CurrentY = 1
   Me.Print "Rate:"
   Me.CurrentX = 92: Me.CurrentY = 1
   Me.Print "Foreign:"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Batavian's Accounting Program", "Setting", "CalcMode", IIf(lvBCalc.Value, "0", "1")
   SaveSetting "Batavian's Accounting Program", "Setting", "Rate", IIf(txtCalc(0) = vbNullString, "0", txtCalc(0))
   SaveSetting "Batavian's Accounting Program", "Setting", "CalcLeft", CStr(Me.Left)
   SaveSetting "Batavian's Accounting Program", "Setting", "CalcTop", CStr(Me.Top)
End Sub

Private Sub lvB_Fill_Click()
   If MyActiveTB Is Nothing Or Trim(txtCalc(2)) = vbNullString Then Exit Sub
   MyActiveTB = txtCalc(2)
   MyActiveTB.SetFocus
End Sub

Private Function GetNumber(Optional intIndex As Integer = -1) As String
   Dim strTemp As String
   Dim intCalc As Integer
   
   If lvBCalc.Value Then intCalc = 2 Else intCalc = 1
   If intIndex <> -1 Then intCalc = intIndex
   
   strTemp = Replace(txtCalc(intCalc), ",", vbNullString)
   strTemp = Replace(strTemp, ".", vbNullString)
   GetNumber = IIf(strTemp = vbNullString, "0", strTemp)
End Function

Private Sub lvB_Number_Click(Index As Integer)
   Dim strTemp As String
   
   If lvBCalc.Value Then
      If MyOperator < [5-Empty] Then
         MyOperator = [5-Empty]
         dblTemp = CDbl(GetNumber)
         txtCalc(2) = 0
      End If
         
      strTemp = GetNumber & CStr(Index)
      
      If strTemp = "0" Then
         txtCalc(2) = "0"
      Else
         txtCalc(2) = Format(strTemp, "#,###")
      End If
   Else
      strTemp = GetNumber & CStr(Index)
      
      If strTemp = "0" Then
         txtCalc(1) = "0"
      Else
         txtCalc(1) = Format(strTemp, "#,###")
      End If
      
      txtCalc(2) = Format(CDbl(GetNumber(0)) * CDbl(strTemp), "#,###")
   End If
End Sub

Private Sub lvB_Operator_Click(Index As Integer)
   Static bIsDivOrSub As Boolean
   
   If Index = 4 Then
      MyOperator = [4-Previous]
      If bIsDivOrSub Then dblRef = CDbl(GetNumber)
      bIsDivOrSub = False
      
      Select Case MyPrevOp
         Case 0: txtCalc(2) = Format(dblTemp * CDbl(txtCalc(2)), "#,###")
         Case 1:
            txtCalc(2) = IIf(Int(dblTemp / CDbl(dblRef)) = 0, "0", Format(dblTemp / CDbl(dblRef), "#,###"))
            dblTemp = CDbl(GetNumber)
         Case 2
            txtCalc(2) = IIf(Int(dblTemp - CDbl(dblRef)) = 0, "0", Format(dblTemp - CDbl(dblRef), "#,###"))
            dblTemp = CDbl(GetNumber)
         Case 3: txtCalc(2) = Format(dblTemp + CDbl(txtCalc(2)), "#,###")
         Case 5: txtCalc(2) = "0"
      End Select
   ElseIf Index = 5 Then
      MyPrevOp = Index
      MyOperator = Index
      txtCalc(2) = "0"
      If Not lvBCalc.Value Then txtCalc(1) = "0"
      bIsDivOrSub = False
   Else
      MyPrevOp = Index
      MyOperator = Index
      bIsDivOrSub = False
      
      If Index = 1 Or Index = 2 Then bIsDivOrSub = True
   End If
End Sub

Private Sub lvBCalc_Click()
   Dim intCount As Integer
   
   For intCount = 0 To 4
      lvB_Operator(intCount).Enabled = lvBCalc.Value
   Next intCount
   
   txtCalc(2).Enabled = lvBCalc.Value
   txtCalc(1).Enabled = Not lvBCalc.Value
End Sub

Private Sub txtCalc_Change(Index As Integer)
   Dim dblTmp As Double
   Dim strTmp1 As String
   Dim strTmp2 As String
   
   If Index <> 2 Then
      strTmp1 = Replace(txtCalc(0), ".", vbNullString)
      strTmp1 = Replace(strTmp1, ",", vbNullString)
      
      strTmp2 = Replace(txtCalc(1), ".", vbNullString)
      strTmp2 = Replace(strTmp2, ",", vbNullString)
      
      dblTmp = CDbl(IIf(strTmp1 = vbNullString, 0, strTmp1)) * CDbl(IIf(strTmp2 = vbNullString, 0, strTmp2))
      
      If dblTmp < 1 Then
         txtCalc(2) = "0"
      Else
         txtCalc(2) = Format(dblTmp, "#,###")
      End If
   End If
End Sub

Private Sub txtCalc_GotFocus(Index As Integer)
   Dim intLoop As Integer
   
   If Index = 2 Then
      txtCalc(0).SetFocus
   End If
End Sub

Private Sub txtCalc_KeyPress(Index As Integer, KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
      ShowBalloonTip txtCalc(Index).hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
End Sub
