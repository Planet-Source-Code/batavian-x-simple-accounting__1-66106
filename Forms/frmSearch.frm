VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Record"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4365
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   4095
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   2925
      TabIndex        =   1
      Top             =   585
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
      Image           =   "frmSearch.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   1845
      TabIndex        =   2
      Top             =   585
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   714
      Caption         =   "&Search"
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
      Image           =   "frmSearch.frx":035E
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private intString2Search As Integer
Private intLVStartSubItem As Integer

Private Sub cboSearch_GotFocus()
   intLVStartSubItem = 1
End Sub

Private Sub cboSearch_KeyUp(KeyCode As Integer, Shift As Integer)
   Static iLast As Integer

   Dim i1st As Integer
   Dim sText As String
   
   If KeyCode = vbKeyDown Then
      SendMessage cboSearch.hWnd, &H14F, 1, ByVal 0&
      Exit Sub
   End If

   If (KeyCode >= 65 And KeyCode <= 90) Or _
      (KeyCode >= 97 And KeyCode <= 122) Or _
      (KeyCode >= 48 And KeyCode <= 57) Or _
      (KeyCode >= 96 And KeyCode <= 105) Or _
      (KeyCode >= 46 And KeyCode <= 48) Or _
      KeyCode = 222 Or KeyCode = vbKeySpace Then

      i1st = cboSearch.SelStart

      If iLast <> 0 Then
         cboSearch.SelStart = iLast
         i1st = iLast
      End If

      sText = CStr(Left(cboSearch.Text, i1st))
      cboSearch.ListIndex = SendMessage(cboSearch.hWnd, &H14C, -1, ByVal CStr(Left(cboSearch.Text, i1st)))

      If cboSearch.ListIndex = -1 Then
          iLast = Len(sText)
          cboSearch.Text = sText
      End If

      cboSearch.SelStart = i1st
      cboSearch.SelLength = Len(cboSearch)
      iLast = 0
   Else
      iLast = i1st
   End If
End Sub

Private Sub Form_Load()
   Dim intCount As Integer
   
   intLVStartSubItem = 1
   intString2Search = CInt(GetSetting("Batavian's Accounting Program", "SearchString", "StringCount", "0"))
   
   If intString2Search > 0 Then
      For intCount = 1 To intString2Search
         If Trim(GetSetting("Batavian's Accounting Program", "SearchString", CStr(intCount), vbNullString)) <> vbNullString Then
            cboSearch.AddItem GetSetting("Batavian's Accounting Program", "SearchString", CStr(intCount), vbNullString)
         End If
      Next intCount
   End If
   
   SetCBDropDownHeigth Me.hWnd, cboSearch, IIf(intString2Search > 20, 20, intString2Search)
   SetMenuForm Me.hWnd
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_OK_Click()
   Dim intCount As Integer
   Dim intLVCount As Integer
   Dim intLVHeader As Integer
   Dim strLV As String
   Dim bRecordExist As Boolean
   
   If cboSearch = vbNullString Then
      ShowBalloonTip GetComboStruckHandle(cboSearch), "Invalid String", "It is obvious that you should type some letters", [2-Information Icon]
      Exit Sub
   End If
   
   For intLVCount = intLVStartSubItem To frmMain.LV1.ListItems.Count
      strLV = frmMain.LV1.ListItems(intLVCount).Text
      
      For intCount = 1 To frmMain.LV1.ColumnHeaders.Count - 1
         strLV = strLV & " " & frmMain.LV1.ListItems(intLVCount).SubItems(intCount)
      Next intCount

      If InStr(UCase(strLV), UCase(cboSearch)) <> 0 Then
         intLVStartSubItem = intLVCount + 1
         
         frmMain.LV1.ListItems(intLVCount).EnsureVisible
         frmMain.LV1.ListItems(intLVCount).Selected = True
         frmMain.SetFocus
         
         bRecordExist = True
         
         Exit For
      End If
   Next intLVCount
   
   If Not bRecordExist Then
      xMsgBox Me.hWnd, "There are no more record containing the word " & cboSearch & ".", vbInformation, "No Record"
      intLVStartSubItem = 1
      Exit Sub
   End If
   
   If intLVStartSubItem >= frmMain.LV1.ListItems.Count Then
      intLVStartSubItem = 1
   End If
   
   For intCount = 0 To cboSearch.ListCount - 1
      If UCase(Trim(cboSearch.List(intCount))) = Trim(UCase(cboSearch)) Then
         Exit Sub
      End If
   Next intCount
   
   If cboSearch.ListCount > 20 Then
      cboSearch.RemoveItem 0
   End If
   
   cboSearch.AddItem cboSearch
   
   For intCount = 1 To cboSearch.ListCount
      SaveSetting "Batavian's Accounting Program", "SearchString", CStr(intCount), cboSearch.List(intCount - 1)
   Next intCount
   
   SaveSetting "Batavian's Accounting Program", "SearchString", "StringCount", cboSearch.ListCount
End Sub
