VERSION 5.00
Begin VB.Form frmNewCoA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Item"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewCoA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optDC 
      Caption         =   "Credit"
      Height          =   495
      Index           =   1
      Left            =   4755
      TabIndex        =   11
      Top             =   1185
      Width           =   780
   End
   Begin VB.OptionButton optDC 
      Caption         =   "Debit"
      Height          =   495
      Index           =   0
      Left            =   3900
      TabIndex        =   10
      Top             =   1185
      Value           =   -1  'True
      Width           =   780
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   3810
      TabIndex        =   8
      Top             =   2205
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
      Image           =   "frmNewCoA.frx":000C
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1729
      Width           =   4395
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1294
      Width           =   1185
   End
   Begin VB.ComboBox cboRemark 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   672
      Width           =   4080
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   237
      Width           =   4395
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Default         =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   2415
      TabIndex        =   9
      Top             =   2205
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
      Image           =   "frmNewCoA.frx":035E
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_CoAFrom 
      Height          =   270
      Left            =   5340
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "Add new Remark of current Category"
      Top             =   705
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   476
      CapAlign        =   2
      BackStyle       =   3
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
      Image           =   "frmNewCoA.frx":0A70
      ImgSize         =   40
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Index           =   8
      Left            =   1095
      TabIndex        =   16
      Top             =   285
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Index           =   7
      Left            =   1095
      TabIndex        =   15
      Top             =   720
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Index           =   6
      Left            =   1095
      TabIndex        =   14
      Top             =   1335
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Index           =   5
      Left            =   1095
      TabIndex        =   13
      Top             =   1770
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Normal Pos :"
      Height          =   195
      Index           =   4
      Left            =   2820
      TabIndex        =   12
      Top             =   1335
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1770
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code No"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1335
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remark"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   285
      Width           =   675
   End
End
Attribute VB_Name = "frmNewCoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bNewItem As Boolean
Public strCategory As String
Public strRemark As String
Public strCode As String
Public strDescription As String
Public bDebit As Boolean
Public lngID As Long

Private strPreCode() As String

Private Sub cboCategory_Click()
   Dim intLoop As Integer
   Dim bExist As Boolean
   
   ReDim strPreCode(0) As String
   
   cboRemark.Clear
   
   CloseConn , Me.hWnd
   ADORS.Open "SELECT Remark,CoA FROM tblCoA WHERE Category = '" & cboCategory & "';", ADOCnn
   
   Do Until ADORS.EOF
      bExist = False
      
      If cboRemark.ListCount > 0 Then
         For intLoop = 0 To cboRemark.ListCount - 1
            If cboRemark.List(intLoop) = ADORS!Remark Then
               bExist = True
            End If
         Next intLoop
      End If
      
      If Not bExist Then
         cboRemark.AddItem ADORS!Remark
         
         ReDim Preserve strPreCode(cboRemark.ListCount - 1) As String
         
         strPreCode(cboRemark.ListCount - 1) = ADORS!CoA
      End If
      
      ADORS.MoveNext
   Loop
   
   If cboCategory = "Assets" Or cboCategory = "Expenses" Then
      optDC(0).Value = True
   Else
      optDC(1).Value = True
   End If
   
   cboRemark.ListIndex = 0
   lvB_CoAFrom.Enabled = True
End Sub

Private Sub cboRemark_Click()
   On Error GoTo ErrHandler
   
   txtCode = Left(strPreCode(cboRemark.ListIndex), 3)
   
ErrHandler:
   txtCode = "0"
End Sub

Private Sub Form_Load()
   cboCategory.AddItem "Assets"
   cboCategory.AddItem "Liabilities"
   cboCategory.AddItem "Equities"
   cboCategory.AddItem "Incomes"
   cboCategory.AddItem "Expenses"
   
   If Not bNewItem Then
      cboCategory = strCategory
      cboRemark = strRemark
      txtCode = strCode
      txtDescription = strDescription
      
      optDC(0).Value = bDebit
      optDC(1).Value = Not bDebit
      
      cboCategory.Enabled = False
      cboRemark.Enabled = False
      txtCode.Enabled = False
      optDC(0).Enabled = False
      optDC(1).Enabled = False
   End If
   
   SetMenuForm Me.hWnd
End Sub

Private Sub lvB_CoAFrom_Click()
   frmNewRemark.strCategory = cboCategory
   frmNewRemark.Show vbModeless, Me
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
   If Index = 0 Then GoTo UnloadMe
   
   If Len(Trim(txtCode)) < 5 Then
      ShowBalloonTip txtCode.hWnd, "Invalid Entry Value", "You should specify a correct entry value!", [2-Information Icon]
      Exit Sub
   End If
   
   If Trim(txtDescription) = vbNullString Then
      ShowBalloonTip txtDescription.hWnd, "Invalid Entry Value", "You should specify a correct entry value!", [2-Information Icon]
      Exit Sub
   End If
   
   If cboCategory.ListIndex < 0 Then
      xMsgBox Me.hWnd, "You should specify a correct Category value!", vbInformation, "Invalid Value"
      Exit Sub
   End If
   
   If cboRemark.ListIndex < 0 Then
      xMsgBox Me.hWnd, "You should specify a correct Remark value!", vbInformation, "Invalid Value"
      Exit Sub
   End If
   
   With ADORS
      If bNewItem Then
         CloseConn , Me.hWnd
         .Open "SELECT * FROM tblCoA;", ADOCnn
         
         Do Until .EOF
            If Trim(UCase(txtCode)) = UCase(!CoA) Then
               ShowBalloonTip txtCode.hWnd, "Invalid Entry Value", "This code number has already been used, try another!", [2-Information Icon]
               Exit Sub
            End If
            
            If Trim(UCase(txtDescription)) = UCase(!Description) And UCase(cboRemark) = UCase(!Remark) Then
               ShowBalloonTip txtDescription.hWnd, "Invalid Entry Value", "This description has already been used, try another!", [2-Information Icon]
               Exit Sub
            End If
            
            ADORS.MoveNext
         Loop
         
         If xMsgBox(Me.hWnd, "Are you sure to add item? " & vbCrLf & vbCrLf & _
            "Category: " & cboCategory & " " & vbCrLf & _
            "Code: " & txtCode & " " & vbCrLf & _
            "Description: " & StrConv(txtDescription, vbProperCase) & " " & vbCrLf & _
            "Remark: " & cboRemark & " " & vbCrLf & _
            "D/C Value: " & IIf(optDC(0).Value, "Debit", "Credit"), _
            vbQuestion Or vbYesNo, "Add Item") = vbYes Then
            
            CloseConn , Me.hWnd
            .Open "SELECT * FROM tblCoA;", ADOCnn, adOpenDynamic, adLockOptimistic
            
            .AddNew
            !CoA = txtCode
            !Remark = cboRemark
            !Description = StrConv(txtDescription, vbProperCase)
            !Category = cboCategory
            !IsDebt = optDC(0).Value
            .Update
            .UpdateBatch
         End If
      Else
         CloseConn , Me.hWnd
         .Open "SELECT Description FROM tblCoA WHERE Remark = '" & strRemark & "' AND Category = '" & strDescription & "';", ADOCnn
         
         Do Until .EOF
            If Trim(UCase(txtDescription)) = UCase(!Description) Then
               ShowBalloonTip txtDescription.hWnd, "Invalid Entry Value", "This field specify an existing account name, try another!", [2-Information Icon]
               Exit Sub
            End If
            
            .MoveNext
         Loop
         
         If xMsgBox(Me.hWnd, "Are you sure to change item? " & vbCrLf & vbCrLf & _
            "From... " & vbCrLf & _
            "Category: " & strCategory & " " & vbCrLf & _
            "Code: " & strCode & " " & vbCrLf & _
            "Description: " & strDescription & " " & vbCrLf & _
            "Remark: " & strRemark & " " & vbCrLf & _
            "Normal Pos: " & IIf(bDebit, "Debit", "Credit") & vbCrLf & vbCrLf & _
            "To..." & vbCrLf & _
            "Category: " & cboCategory & " " & vbCrLf & _
            "Code: " & txtCode & " " & vbCrLf & _
            "Description: " & StrConv(Trim(txtDescription), vbProperCase) & " " & vbCrLf & _
            "Remark: " & cboRemark & " " & vbCrLf & _
            "Normal Pos: " & IIf(optDC(0).Value, "Debit", "Credit") & " " & vbCrLf & vbCrLf & _
            "Note: This action affect all records using current account name. ", _
            vbQuestion Or vbYesNo, "Add Item") = vbYes Then
            
            CloseConn , Me.hWnd
            .Open "SELECT * FROM tblJournal WHERE (DebitCoA = '" & strDescription & "' OR CreditCoA = '" & strDescription & "') " & _
                  "AND Remark = '" & strRemark & "' AND Category = '" & strCategory & "';", ADOCnn, adOpenDynamic, adLockOptimistic
            
            If ADORS.RecordCount > 0 Then
               Do Until ADORS.EOF
                  If IsNull(!DebitCoA) Or !DebitCoA = vbNullString Then
                     !CreditCoa = StrConv(Trim(txtDescription), vbProperCase)
                  Else
                     !DebitCoA = StrConv(Trim(txtDescription), vbProperCase)
                  End If
                  
                  ADORS.MoveNext
               Loop
            End If
            
            CloseConn , Me.hWnd
            .Open "SELECT * FROM tblCoA WHERE ID = " & lngID & ";", ADOCnn, adOpenDynamic, adLockOptimistic
            
            !CoA = txtCode
            !Remark = cboRemark
            !Description = StrConv(Trim(txtDescription), vbProperCase)
            !Category = cboCategory
            !IsDebt = optDC(0).Value
            .Update
            .UpdateBatch
         End If
      End If
   End With
   
   MyTBDesc = vbNullString
   MyTBRem = vbNullString
   
   Call frmCoA.FillLV
      
UnloadMe:
   Unload Me
End Sub
