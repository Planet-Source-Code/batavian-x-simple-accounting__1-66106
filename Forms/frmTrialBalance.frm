VERSION 5.00
Begin VB.Form frmTrialBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Balance - Print Option"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   Icon            =   "frmTrialBalance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMonth 
      Enabled         =   0   'False
      Height          =   315
      Left            =   255
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1072
      Width           =   3240
   End
   Begin Project1.lvButtons_H lvB_Month 
      Height          =   495
      Index           =   0
      Left            =   255
      TabIndex        =   1
      Tag             =   "Report records for this month only."
      Top             =   270
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   873
      Caption         =   "This Month"
      CapAlign        =   2
      BackStyle       =   2
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
      cGradient       =   0
      CapStyle        =   2
      Mode            =   2
      Value           =   -1  'True
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Month 
      Height          =   495
      Index           =   1
      Left            =   1530
      TabIndex        =   2
      Tag             =   "Report records for another month."
      Top             =   270
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   873
      Caption         =   "Other Month"
      CapAlign        =   2
      BackStyle       =   2
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
      cGradient       =   0
      CapStyle        =   2
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   1620
      TabIndex        =   3
      Top             =   1695
      Width           =   1875
      _ExtentX        =   3307
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
      Image           =   "frmTrialBalance.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   255
      TabIndex        =   4
      Top             =   1695
      Width           =   1710
      _ExtentX        =   3016
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
      Image           =   "frmTrialBalance.frx":035E
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strDate As String

Private Sub PrintTrialBalance()
   Const strProgress As String = "Printing, please wait ...."
   
   Dim MyPrinter As New clsPrinter
   Dim strAccShadow() As String
   Dim strAccount() As String
   Dim strRemShadow() As String
   Dim strRemark() As String
   Dim strCategory() As String
   Dim bDebitPos() As Boolean
   Dim dblValue() As Double
   Dim bFlag() As Boolean
   Dim dblTotalD() As Double
   Dim dblTotalC() As Double
   Dim strSQL As String
   Dim strTemp As String
   Dim strTemp2 As String
   Dim lngCount As Long
   Dim intYear As Integer
   Dim intmonth As Integer
   Dim dblBalDebit As Double
   Dim dblBalCredit As Double
   Dim intPage As Integer
   Dim intCount As Integer
   Dim intStart As Integer
   Dim strExpense As String
   Dim dblExpense As Double
   Dim dblExpenseD As Double
   Dim dblExpenseC As Double
   Dim dblDebit As Double
   Dim dblCredit As Double
   Dim bExpensePrint As Boolean
   Dim strNumberD As String
   Dim strNumberC As String
   Dim strNumberB As String
   Dim bNegativeD As Boolean
   Dim bNegativeC As Boolean
   Dim bNegativeB As Boolean

   strExpense = "Total Admin Expenses/Cost"
   
   MyPrinter.AppName = App.Title
   MyPrinter.PageSize = eA4size
   MyPrinter.Orientation = qPortrait
   MyPrinter.ScaleMode = eTwip
   MyPrinter.MarginBottom = 300
   MyPrinter.MarginTop = 600
   MyPrinter.MarginLeft = 450
   MyPrinter.MarginRight = 450
   
   Set MyPrinter.PicBox = frmMain.picLogo
   
   MyPrinter.AddPic frmMain.picLogo, 1, 32, 20, 810, 810
   
   frmMain.picLogo.FontName = "Arial"
   frmMain.picLogo.FontBold = True
   frmMain.picLogo.FontSize = 12
   frmMain.picLogo.ScaleMode = vbTwips
   
   MyPrinter.AddText "<LINDENT=990>PT MAJUKO UTAMA INDONESIA<SIZE=12><LINDENT=" & 10500 - frmMain.picLogo.TextWidth("Trial Balance") & ">" & "Trial Balance", "Arial", 14, True
   MyPrinter.TextItem(MyPrinter.TextItem.Count).Top = -90 '210
   
   frmPrintProgress.Show vbModeless, Me
   
   strSQL = "SELECT DISTINCT AccountName, Remark FROM "
   strSQL = strSQL & "(SELECT DebitCoA AS AccountName, Remark "
   strSQL = strSQL & "FROM tblJournal "
   strSQL = strSQL & "WHERE " & strDate
   strSQL = strSQL & "AND IsNull(CreditCoA) "
   strSQL = strSQL & "UNION ALL "
   strSQL = strSQL & "SELECT CreditCoA AS AccountName, Remark "
   strSQL = strSQL & "FROM tblJournal "
   strSQL = strSQL & "WHERE " & strDate
   strSQL = strSQL & "AND IsNull(DebitCoA));"
   
'   Debug.Print strSQL
'   Set MyPrinter = Nothing
'   Exit Sub
   
   CloseConn , Me.hWnd
   ADORS.Open strSQL, ADOCnn
   
   If ADORS.RecordCount < 1 Then
      Unload frmPrintProgress
      xMsgBox Me.hWnd, "There are no record! ", vbInformation, "No Record"
      Exit Sub
   End If

   Do Until ADORS.EOF
      strTemp = strTemp & "|" & ADORS!AccountName
      strTemp2 = strTemp2 & "|" & ADORS!Remark
      ADORS.MoveNext
   Loop
   
   strTemp = Right(strTemp, Len(strTemp) - 1)
   strTemp2 = Right(strTemp2, Len(strTemp2) - 1)
   
   strAccShadow() = Split(strTemp, "|")
   strRemShadow() = Split(strTemp2, "|")
   
   ReDim strAccount(UBound(strAccShadow))
   ReDim strRemark(UBound(strAccShadow))
   ReDim strCategory(UBound(strAccShadow))
   ReDim dblValue(UBound(strAccShadow))
   ReDim dblTotalD(UBound(strAccShadow))
   ReDim dblTotalC(UBound(strAccShadow))
   ReDim bFlag(UBound(strAccShadow))
   ReDim bDebitPos(UBound(strAccShadow))
   
   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblCoA ORDER BY tblCoA.CoA;", ADOCnn
   
   Do Until ADORS.EOF
      For lngCount = 0 To UBound(strAccShadow)
         If ADORS!Description = strAccShadow(lngCount) And ADORS!Remark = strRemShadow(lngCount) Then
            Dim strClp As String
            
            strAccount(intCount) = ADORS!Description
            strRemark(intCount) = ADORS!Remark
            strCategory(intCount) = ADORS!Category
            bDebitPos(intCount) = ADORS!IsDebt
            
            intCount = intCount + 1
            Exit For
         End If
      Next lngCount
      
      ADORS.MoveNext
   Loop
   
   For lngCount = 0 To UBound(strAccount)
      frmPrintProgress.Label1 = strProgress & vbCrLf & Round(lngCount / UBound(strAccount) * 100) & "% complete."
      DoEvents
      
      strSQL = "SELECT SUM(DebitValue) As TotalD, SUM(CreditValue) AS TotalC FROM ("
      strSQL = strSQL & "SELECT Date, Category, Debit AS DebitValue, 0 AS CreditValue FROM tblJournal "
      strSQL = strSQL & "WHERE " & strDate & " AND (ISNULL(CreditCoA) OR CreditCoA = '') AND DebitCoA = '" & strAccount(lngCount) & "' AND Remark = '" & strRemark(lngCount) & "' "
      strSQL = strSQL & "UNION ALL "
      strSQL = strSQL & "SELECT Date, Category, 0 AS DebitValue, Credit AS CreditValue FROM tblJournal "
      strSQL = strSQL & "WHERE " & strDate & " AND (ISNULL(DebitCoA) OR DebitCoA = '') AND CreditCoA = '" & strAccount(lngCount) & "' AND Remark = '" & strRemark(lngCount) & "');"
      
      CloseConn
      ADORS.Open strSQL, ADOCnn
      
      If ADORS.RecordCount > 0 Then
         If Not IsNull(ADORS!TotalD) Then
            dblTotalD(lngCount) = ADORS!TotalD
         End If
         
         If Not IsNull(ADORS!TotalC) Then
            dblTotalC(lngCount) = ADORS!TotalC
         End If
         
         If bDebitPos(lngCount) Then
             dblValue(lngCount) = dblTotalD(lngCount) - dblTotalC(lngCount)
         Else
             dblValue(lngCount) = dblTotalC(lngCount) - dblTotalD(lngCount)
         End If
      End If
   Next lngCount
   
   frmMain.picLogo.FontName = "Arial"
   frmMain.picLogo.FontBold = True
   frmMain.picLogo.FontSize = 10
   
   strTemp = "As of " & cboMonth & " " & MaxDayOfMonth(Year(Date), cboMonth.ListIndex + 1) & GetOrdinal(MaxDayOfMonth(Year(Date), cboMonth.ListIndex + 1)) & ", " & Year(Date)
   
   MyPrinter.AddText " ", "Arial", 24
   MyPrinter.AddText "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strTemp) & ">" & strTemp, "Arial", 10, True
   MyPrinter.AddText " ", "Arial", 9
   
   MyPrinter.AddLine 30, 1120, 10500, 1120, 1, 3
   MyPrinter.AddLine 30, 1090, 10500, 1090, 1
   MyPrinter.AddLine 30, 1510, 10500, 1510, 1
   
   frmMain.picLogo.FontSize = 9
   
   MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth("Debit") & ">Debit<LINDENT=" & 3390 - frmMain.picLogo.TextWidth("Balance") & ">Balance<LINDENT=" & 3690 + ((3210 - frmMain.picLogo.TextWidth("Account Name")) / 2) & ">Account Name<LINDENT=" & 8700 - frmMain.picLogo.TextWidth("Balance") & ">Balance<LINDENT=" & 10500 - frmMain.picLogo.TextWidth("Credit") & ">Credit", "Arial", 9, True
   MyPrinter.AddText " ", "Arial", 9
   
   frmMain.picLogo.FontBold = False
   
   bExpensePrint = False
   
   For lngCount = 0 To UBound(strAccount)
      If strRemark(lngCount) = "Breakdown Of Admin Expense" Or strRemark(lngCount) = "Breakdown Of Cost" Then
         dblExpense = dblExpense + dblValue(lngCount)
         dblExpenseC = dblExpenseC + dblTotalC(lngCount)
         dblExpenseD = dblExpenseD + dblTotalD(lngCount)
      End If
   Next lngCount
   
   intStart = MyPrinter.TextItem.Count + 1
      
   For lngCount = 0 To UBound(strAccount)
      If strAccount(lngCount) <> vbNullString Then
         strNumberD = FormatNumber(dblTotalD(lngCount), 0, vbTrue, vbTrue, vbTrue)
         strNumberC = FormatNumber(dblTotalC(lngCount), 0, vbTrue, vbTrue, vbTrue)
         strNumberB = FormatNumber(dblValue(lngCount), 0, vbTrue, vbTrue, vbTrue)
         
         bNegativeD = dblTotalD(lngCount) < 0
         bNegativeC = dblTotalC(lngCount) < 0
         bNegativeB = dblValue(lngCount) < 0
         
         If bNegativeD Then
            strNumberD = Left(strNumberD, Len(strNumberD) - 1)
         End If
         
         If bNegativeC Then
            strNumberC = Left(strNumberC, Len(strNumberC) - 1)
         End If
         
         If bNegativeB Then
            strNumberB = Left(strNumberB, Len(strNumberB) - 1)
         End If
      
         If strRemark(lngCount) <> "Breakdown Of Admin Expense" And strRemark(lngCount) <> "Breakdown Of Cost" Then
            If strCategory(lngCount) = "Liabilities" And Not bExpensePrint Then
               bExpensePrint = True
               
'               strNumberD = FormatNumber(dblExpenseD, 0, vbTrue, vbTrue, vbTrue)
'               strNumberC = FormatNumber(dblExpenseC, 0, vbTrue, vbTrue, vbTrue)
'               strNumberB = FormatNumber(dblExpense, 0, vbTrue, vbTrue, vbTrue)
'
'               bNegativeD = dblExpenseD < 0
'               bNegativeC = dblExpenseC < 0
'               bNegativeB = dblExpense < 0
'
'               If bNegativeD Then
'                  strNumberD = Left(strNumberD, Len(strNumberD) - 1)
'               End If
'
'               If bNegativeC Then
'                  strNumberC = Left(strNumberC, Len(strNumberC) - 1)
'               End If
'
'               If bNegativeB Then
'                  strNumberB = Left(strNumberB, Len(strNumberB) - 1)
'               End If

               MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=3690>" & strExpense & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9
               MyPrinter.TextItem(MyPrinter.TextItem.Count).Top = 30
            
               strNumberD = FormatNumber(dblTotalD(lngCount), 0, vbTrue, vbTrue, vbTrue)
               strNumberC = FormatNumber(dblTotalC(lngCount), 0, vbTrue, vbTrue, vbTrue)
               strNumberB = FormatNumber(dblValue(lngCount), 0, vbTrue, vbTrue, vbTrue)
               
               bNegativeD = dblTotalD(lngCount) < 0
               bNegativeC = dblTotalC(lngCount) < 0
               bNegativeB = dblValue(lngCount) < 0
               
               If bNegativeD Then
                  strNumberD = Left(strNumberD, Len(strNumberD) - 1)
               End If
               
               If bNegativeC Then
                  strNumberC = Left(strNumberC, Len(strNumberC) - 1)
               End If
               
               If bNegativeB Then
                  strNumberB = Left(strNumberB, Len(strNumberB) - 1)
               End If
      
               If bDebitPos(lngCount) Then
                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9
               Else
                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9
               End If
            Else
               If bDebitPos(lngCount) Then
                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9
               Else
                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9
               End If
'               If bDebitPos(lngCount) Then
'                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", "") & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth("0") & ">0", "Arial", 9
'               Else
'                  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth("0") & ">0" & "<LINDENT=3690>" & strAccount(lngCount) & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", "") & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", ""), "Arial", 9
'               End If
            End If
               
            MyPrinter.TextItem(MyPrinter.TextItem.Count).Top = 30
         End If
      End If
   Next lngCount
   
   For intCount = intStart To MyPrinter.TextItem.Count
      If intCount Mod 2 = 0 Then
         MyPrinter.TextItem(intCount).BorderShading = RGB(239, 239, 239)
         MyPrinter.TextItem(intCount).BorderLine = 0
         MyPrinter.TextItem(intCount).ShowBorder = True
      End If
   Next intCount
   
'   Debug.Print MyPrinter.TextItem(MyPrinter.TextItem.Count).Text
'   MyPrinter.TextItem.Remove MyPrinter.TextItem.Count
   
   frmMain.picLogo.FontBold = True
   
   For lngCount = 0 To UBound(strAccount)
      If bDebitPos(lngCount) Then
         dblBalDebit = dblBalDebit + dblValue(lngCount)
      Else
         dblBalCredit = dblBalCredit + dblValue(lngCount)
      End If

      dblDebit = dblDebit + dblTotalD(lngCount)
      dblCredit = dblCredit + dblTotalC(lngCount)
   Next lngCount
   
   MyPrinter.AddText " ", "Arial", 12
'  MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(strNumberB) & ">" & strNumberB & IIf(bNegativeB, ")", "") & "<LINDENT=3690>" & strExpense & "<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", "") & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth("0") & ">0", "Arial", 9
   MyPrinter.AddText "<LINDENT=" & 1590 - frmMain.picLogo.TextWidth(FormatNumber(dblBalDebit, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblBalDebit, 0, vbTrue, vbTrue, vbTrue) & "<LINDENT=" & 3390 - frmMain.picLogo.TextWidth(FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue) & "<LINDENT=3690>Total<LINDENT=" & 8700 - frmMain.picLogo.TextWidth(FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue) & "<LINDENT=" & 10500 - frmMain.picLogo.TextWidth(FormatNumber(dblBalCredit, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblBalCredit, 0, vbTrue, vbTrue, vbTrue), "Arial", 9, True
   
   intPage = MyPrinter.Pages
   
   MyPrinter.AddLine 7110, 1075, 7110, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 300, 1
   MyPrinter.AddLine 3540, 1075, 3540, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 300, 1
   MyPrinter.AddLine 30, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip - 90, 10500, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip - 90, MyPrinter.TextItem(MyPrinter.TextItem.Count).StartPage
   MyPrinter.AddLine 30, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 300, 10500, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 300, MyPrinter.TextItem(MyPrinter.TextItem.Count).StartPage, 3
   MyPrinter.AddLine 30, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 330, 10500, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 330, MyPrinter.TextItem(MyPrinter.TextItem.Count).StartPage
   
   Unload frmPrintProgress
      
   MyPrinter.Preview Me, 120 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
   
   Set MyPrinter = Nothing
End Sub

Private Sub cboMonth_Click()
   strDate = "Date BETWEEN #1/1/" & Year(Date) & "# AND #" & cboMonth.ListIndex + 1 & "/" & MaxDayOfMonth(Year(Date), cboMonth.ListIndex + 1) & "/" & Year(Date) & "# "
End Sub

Private Sub Form_Load()
   cboMonth.AddItem "January"
   cboMonth.AddItem "February"
   cboMonth.AddItem "March"
   cboMonth.AddItem "April"
   cboMonth.AddItem "May"
   cboMonth.AddItem "June"
   cboMonth.AddItem "July"
   cboMonth.AddItem "August"
   cboMonth.AddItem "September"
   cboMonth.AddItem "October"
   cboMonth.AddItem "November"
   cboMonth.AddItem "December"
   
   cboMonth.ListIndex = Month(Date) - 1
   
   strDate = "Date BETWEEN #1/1/" & Year(Date) & "# AND #" & cboMonth.ListIndex + 1 & "/" & MaxDayOfMonth(Year(Date), cboMonth.ListIndex + 1) & "/" & Year(Date) & "# "
   
   SetMenuForm Me.hWnd
'   Me.Caption = "Print Option - " & cboMonth
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_Month_Click(Index As Integer)
   If lvB_Month(0).Value Then
      cboMonth.ListIndex = Month(Date) - 1
   End If
   
   cboMonth.Enabled = lvB_Month(1).Value
End Sub

Private Sub lvB_OK_Click()
   Me.Enabled = False
   PrintTrialBalance
   Me.Enabled = True
End Sub
