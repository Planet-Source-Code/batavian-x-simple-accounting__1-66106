VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmGL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Ledger - Print Option"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton lvButtons_H2 
      Caption         =   " This year to date"
      Height          =   495
      Index           =   1
      Left            =   398
      TabIndex        =   9
      Tag             =   "Report whole year records."
      Top             =   1425
      Width           =   2175
   End
   Begin VB.OptionButton lvButtons_H2 
      Caption         =   " Other month:"
      Height          =   495
      Index           =   2
      Left            =   398
      TabIndex        =   8
      Tag             =   "Report records for another month."
      Top             =   510
      Width           =   2175
   End
   Begin VB.OptionButton lvButtons_H2 
      Caption         =   " This month"
      Height          =   495
      Index           =   0
      Left            =   398
      TabIndex        =   7
      Tag             =   "Report records for this month."
      Top             =   135
      Value           =   -1  'True
      Width           =   2175
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   3735
      TabIndex        =   4
      Top             =   1020
      Width           =   255
      _ExtentX        =   476
      _ExtentY        =   503
      _Version        =   327681
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtYear 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2648
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1005
      Width           =   1125
   End
   Begin VB.ComboBox cboMonth 
      Enabled         =   0   'False
      Height          =   315
      Left            =   653
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1005
      Width           =   1935
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   1035
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Tag             =   "GL's table with 2 columns."
      Top             =   1980
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1826
      Caption         =   "2 Columns"
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
      ImgAlign        =   4
      Image           =   "frmGL.frx":000C
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   1035
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Tag             =   "GL's table for fast view of accounts, because of its simplicity."
      Top             =   1980
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1826
      Caption         =   "T-Accounts"
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
      ImgAlign        =   4
      Image           =   "frmGL.frx":0C5E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Cancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   1988
      TabIndex        =   5
      Top             =   3165
      Width           =   1890
      _ExtentX        =   3334
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
      Image           =   "frmGL.frx":18B0
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   293
      TabIndex        =   6
      Top             =   3165
      Width           =   2055
      _ExtentX        =   3625
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
      Image           =   "frmGL.frx":1C02
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strAccount() As String
Private strRemark() As String
Private bDebitPos() As Boolean
Private strYear As String
Private strDate As String
Private strConDate As String
Private intmonth As Integer

'Private Sub cboMonth_Click()
'   lvB_OK.Enabled = IsDBExist(cboMonth.ListIndex + 1, txtYear)
'   lvB_Cancel.Enabled = lvB_OK.Enabled
'End Sub

Private Sub Form_Load()
   txtYear = Val(Year(Date))

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
   SetCBDropDownHeigth Me.hWnd, cboMonth, 12
   UpDown1.Left = txtYear.Left + txtYear.Width - UpDown1.Width - 15
   SetMenuForm Me.hWnd
End Sub

Private Sub StartPrintGL()
   Dim strSQL As String
   Dim strTemp As String
   Dim strTemp2 As String
   Dim intYear As Integer
   Dim MyPrinter As New clsPrinter
   Dim intIndex As Integer
   Dim intCount As Integer
   Dim strAccShadow() As String
   Dim strRemShadow() As String

   intYear = Year(Date)
   intmonth = Month(Date)

   If lvButtons_H2(2).Value Then
      intYear = txtYear
      intmonth = cboMonth.ListIndex + 1
   End If

   strDate = "BETWEEN #" & IIf(lvButtons_H2(1).Value, 1, intmonth) & "/1/" & intYear & "# AND #" & intmonth & "/" & IIf(lvButtons_H2(1).Value, Day(Date), MaxDayOfMonth(intYear, intmonth)) & "/" & intYear & "#"
      
   If Not lvButtons_H2(1).Value And intmonth > 1 Then
      strConDate = "BETWEEN #1/1/" & intYear & "# AND #" & intmonth - 1 & "/" & MaxDayOfMonth(intYear, intmonth - 1) & "/" & intYear & "#"
   End If

   strSQL = "SELECT DISTINCT AccountName, Remark FROM ("
   strSQL = strSQL & "SELECT tblJournal.DebitCoA AS AccountName, tblJournal.Remark "
   strSQL = strSQL & "FROM tblJournal "
   strSQL = strSQL & "WHERE (tblJournal.Date " & strDate & ") "
   strSQL = strSQL & "AND IsNull(DebitCoA)=False "
   strSQL = strSQL & "UNION ALL "
   strSQL = strSQL & "SELECT tblJournal.CreditCoA AS AccountName, tblJournal.Remark "
   strSQL = strSQL & "FROM tblJournal "
   strSQL = strSQL & "WHERE (tblJournal.Date " & strDate & ") "
   strSQL = strSQL & "AND IsNull(CreditCoA)=False);"

'   Debug.Print strSQL
'   Unload Me
'   Exit Sub
   
   CloseConn , Me.hWnd
   ADORS.Open strSQL, ADOCnn

   If ADORS.RecordCount < 1 Then
      Unload frmPrintProgress
      xMsgBox Me.hWnd, "There are no record! ", vbInformation, "No Record"
      Exit Sub
   Else
      frmPrintProgress.Show vbModeless, Me
   End If

   Do Until ADORS.EOF
      strTemp = strTemp & "|" & ADORS!AccountName
      strTemp2 = strTemp2 & "|" & ADORS!Remark
      ADORS.MoveNext
   Loop

   CloseConn False, Me.hWnd

   strTemp = Right(strTemp, Len(strTemp) - 1)
   strTemp2 = Right(strTemp2, Len(strTemp2) - 1)

   strAccShadow() = Split(strTemp, "|")
   strRemShadow() = Split(strTemp2, "|")

   ReDim strAccount(UBound(strAccShadow))
   ReDim strRemark(UBound(strAccShadow))
   ReDim bDebitPos(UBound(strAccShadow))

   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblCoA ORDER BY CoA;", ADOCnn

   Do Until ADORS.EOF
      For intIndex = 0 To UBound(strAccShadow)
         If strAccShadow(intIndex) = ADORS!Description And strRemShadow(intIndex) = ADORS!Remark Then
            strAccount(intCount) = ADORS!Description
            strRemark(intCount) = ADORS!Remark
            bDebitPos(intCount) = ADORS!IsDebt

            intCount = intCount + 1
            Exit For
         End If
      Next intIndex

      ADORS.MoveNext
   Loop

   PrintGL IIf(lvButtons_H1(0).Value, 0, 1), MyPrinter

   MyPrinter.Preview Me, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY

   Set MyPrinter = Nothing
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_OK_Click()
   If lvButtons_H1(0).Enabled Then
      Me.Enabled = False
      StartPrintGL
      Me.Enabled = True
   End If
End Sub

Private Sub lvButtons_H1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If MyCoolTip.ParenthWnd = lvButtons_H1(Index).hWnd And _
      MyCoolTip.TipText = lvButtons_H1(Index).Tag Then Exit Sub

   MyCoolTip.Style = [Balloon Tip]
   MyCoolTip.ParenthWnd = lvButtons_H1(Index).hWnd
   MyCoolTip.TipText = lvButtons_H1(Index).Tag
   MyCoolTip.Create
End Sub

Private Sub lvButtons_H1_MouseOnButton(Index As Integer, OnButton As Boolean)
   If Not OnButton Then MyCoolTip.Destroy
End Sub

Private Sub lvButtons_H2_Click(Index As Integer)
   cboMonth.Enabled = lvButtons_H2(2).Value
   txtYear.Enabled = lvButtons_H2(2).Value
   UpDown1.Enabled = lvButtons_H2(2).Value

   If Index = 1 Then
      txtYear = Year(Date)
   ElseIf Index = 0 Then
      txtYear = Year(Date)
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

Private Sub PrintGL(Index As Integer, qPrinter As clsPrinter)
   Const strProgress As String = "Printing, please wait ...."

   Dim dblDebit As Double
   Dim dblCredit As Double
   Dim strSQL As String
   Dim strTemp As String
   Dim strTemp2 As String
   Dim strLongTmp As String
   Dim strLongTmp2 As String
   Dim intCount As Integer
   Dim intDC As Integer
   Dim intPage As Integer
   Dim intVLineX() As Integer
   Dim intVLineX2() As Integer
   Dim strDebit() As String
   Dim strCredit() As String
   Dim intStart() As Integer
   Dim bStartPage As Boolean
   Dim strNumber As String
   Dim strNumberD As String
   Dim strNumberC As String
   Dim bNegative As Boolean
   Dim bNegativeD As Boolean
   Dim bNegativeC As Boolean
   Dim dblTotalLast As Double
   Dim intPageStart() As Integer
   Dim intPageEnd() As Integer
   
   ReDim intVLineX(0)
   ReDim intPageStart(UBound(strAccount) + 1)
   ReDim intPageEnd(UBound(strAccount) + 1)

   frmPrintProgress.Label1 = strProgress & vbCrLf & "0% complete."
   AddReportLogo frmMain.picLogo, "General Ledgers", qPrinter, qPortrait, 1

   With qPrinter
      .AddText " ", "Arial", 6
      
      For intCount = 0 To UBound(strAccount) '- 1
         If strRemark(intCount) <> vbNullString Then
            ReDim strDebit(0)
            ReDim strCredit(0)

            frmPrintProgress.Label1 = strProgress & vbCrLf & Round(intCount / UBound(strAccount) * 100) & "% complete."
            DoEvents

            frmMain.picLogo.FontBold = True
            frmMain.picLogo.FontSize = 10

            dblDebit = 0
            dblCredit = 0
            
            strTemp = IIf(lvButtons_H2(1).Value, "As of " & GetMonthName(Month(Date)) & " " & Day(Date) & GetOrdinal(Day(Date)) & ", ", cboMonth.List(IIf(lvButtons_H2(2).Value, cboMonth.ListIndex + 1, Month(Date)) - 1) & ", ") & IIf(lvButtons_H2(2).Value, txtYear, Year(Date))
            strTemp = "<LINDENT=180>" & strRemark(intCount) & " - " & strAccount(intCount) & "<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strTemp) & ">" & strTemp
            
            .AddText strTemp, "Arial", 10, True
            .TextItem(.TextItem.Count).Top = IIf(bStartPage, -150, 300)

            frmMain.picLogo.FontSize = 9

            If Not bStartPage Then bStartPage = True
            
            If Index = 0 Then
               strTemp = "<LINDENT=180>Date<LINDENT=780>Description<LINDENT=4250>Customer<LINDENT=6480>Slip No.<LINDENT=" & 9150 - frmMain.picLogo.TextWidth("Debit") & ">Debit<LINDENT=" & 10650 - frmMain.picLogo.TextWidth("Credit") & ">Credit"
               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 225
            End If

            .AddText " ", "Arial", 6, False

            ReDim Preserve intVLineX2(UBound(intVLineX) + 1) As Integer
            intVLineX2(UBound(intVLineX2)) = .TextItem.Count

            frmMain.picLogo.FontBold = False

            strSQL = "SELECT * "
            strSQL = strSQL & "FROM ("
            strSQL = strSQL & "SELECT tblJournal.Date, tblJournal.Voucher, tblJournal.Description, tblJournal.Customer, tblJournal.DebitCoA AS AccountName, tblJournal.Remark, tblJournal.Debit AS AccValue, 'Debit' AS Flag "
            strSQL = strSQL & "FROM tblJournal "
            strSQL = strSQL & "WHERE (Date " & strDate & ") AND ISNULL(DebitCoA) = FALSE "
            strSQL = strSQL & "UNION ALL "
            strSQL = strSQL & "SELECT tblJournal.Date, tblJournal.Voucher, tblJournal.Description, tblJournal.Customer, tblJournal.CreditCoA AS AccountName, tblJournal.Remark, tblJournal.Credit AS AccValue, 'Credit' AS Flag "
            strSQL = strSQL & "FROM tblJournal "
            strSQL = strSQL & "WHERE (Date " & strDate & ") AND ISNULL(CreditCoA) = FALSE) "
            strSQL = strSQL & "WHERE AccountName = '" & strAccount(intCount) & "' AND Remark = '" & strRemark(intCount) & "' ORDER BY Date;"

            CloseConn , Me.hWnd
            ADORS.Open strSQL, ADOCnn

            Do Until ADORS.EOF
               strNumber = FormatNumber(ADORS!accvalue, 0, vbTrue, vbTrue, vbTrue)
               bNegative = ADORS!accvalue < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               If Index = 0 Then
                  strLongTmp = ADORS!Description
                  strLongTmp2 = IIf(IsNull(ADORS!Customer), " ", ADORS!Customer)

                  frmMain.picLogo.FontSize = 8

                  If frmMain.picLogo.TextWidth(strLongTmp) > 3380 Then
                     Do Until frmMain.picLogo.TextWidth(strLongTmp & "...") <= 3380
                        strLongTmp = Left(strLongTmp, Len(strLongTmp) - 1)
                     Loop

                     strLongTmp = strLongTmp & "..."
                  End If

                  If frmMain.picLogo.TextWidth(strLongTmp2) > 2140 Then
                     Do Until frmMain.picLogo.TextWidth(strLongTmp2 & "...") <= 2140
                        strLongTmp2 = Left(strLongTmp2, Len(strLongTmp2) - 1)
                     Loop

                     strLongTmp2 = strLongTmp2 & "..."
                  End If
                  
                  strTemp = "<LINDENT=180>" & Month(ADORS!Date) & "/" & Day(ADORS!Date) & "<LINDENT=780><SIZE=8>" & strLongTmp & "<LINDENT=4250>" & strLongTmp2 & "</SIZE><LINDENT=6480>" & ADORS!Voucher & "<LINDENT=" & IIf(ADORS!Flag = "Debit", 9150, 10650) - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               Else
                  strTemp = "<LINDENT=" & IIf(ADORS!Flag = "Debit", 180, 5325 + 300) & ">" & Month(ADORS!Date) & "/" & Day(ADORS!Date) & "/" & Year(ADORS!Date) & "<LINDENT=" & IIf(ADORS!Flag = "Debit", 5325 - 300, 10650) - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")

                  If ADORS!Flag = "Debit" Then
                     ReDim Preserve strDebit(UBound(strDebit) + 1)

                     strDebit(UBound(strDebit)) = strTemp
                  Else
                     ReDim Preserve strCredit(UBound(strCredit) + 1)

                     strCredit(UBound(strCredit)) = strTemp
                  End If
               End If

               dblDebit = dblDebit + CDbl(IIf(ADORS!Flag = "Debit", ADORS!accvalue, 0))
               dblCredit = dblCredit + CDbl(IIf(ADORS!Flag = "Debit", 0, ADORS!accvalue))

               If Index = 0 Then
                  .AddText strTemp, "Arial", 8
                  .TextItem(.TextItem.Count).Top = 90
               End If

               ADORS.MoveNext
            Loop

            If Index = 1 Then
               If UBound(strCredit) > UBound(strDebit) Then
                  For intDC = 1 To UBound(strCredit)
                     If intDC <= UBound(strDebit) Then
                        .AddText strDebit(intDC) & strCredit(intDC), "Arial", 9
                     Else
                        .AddText strCredit(intDC), "Arial", 9
                     End If

                     .TextItem(.TextItem.Count).Top = 90
                  Next intDC
               Else
                  For intDC = 1 To UBound(strDebit)
                     If intDC <= UBound(strCredit) Then
                        .AddText strDebit(intDC) & strCredit(intDC), "Arial", 9
                     Else
                        .AddText strDebit(intDC), "Arial", 9
                     End If

                     .TextItem(.TextItem.Count).Top = 90
                  Next intDC
               End If
            End If

            frmMain.picLogo.FontBold = True
            
            If Index = 0 Then
               strNumberD = IIf(dblDebit, FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue), "0")
               strNumberC = IIf(dblCredit, FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue), "0")
               
               bNegativeD = IIf(dblDebit, FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue), 0) < 0
               bNegativeC = IIf(dblCredit, FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue), 0) < 0
               
               If bNegativeD Then
                  strNumberD = Left(strNumberD, Len(strNumberD) - 1)
               End If
               
               If bNegativeC Then
                  strNumberC = Left(strNumberC, Len(strNumberC) - 1)
               End If

               strTemp = "<LINDENT=5280>Total<LINDENT=" & 9150 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", "")
            Else
               If bDebitPos(intCount) Then
                  strNumber = FormatNumber(dblDebit - dblCredit, 0, vbTrue, vbTrue, vbTrue)
                  bNegative = (dblDebit - dblCredit) < 0
                  
                  If bNegative Then
                     strNumber = Left(strNumber, Len(strNumber) - 1)
                  End If
               
                  strTemp = "<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                  dblCredit = dblDebit
               Else
                  strNumber = FormatNumber(dblCredit - dblDebit, 0, vbTrue, vbTrue, vbTrue)
                  bNegative = (dblCredit - dblDebit) < 0
                  
                  If bNegative Then
                     strNumber = Left(strNumber, Len(strNumber) - 1)
                  End If
                  
                  strTemp = "<LINDENT=" & 5325 - 300 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                  dblDebit = dblCredit
               End If

               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 300

               strNumberD = IIf(dblDebit, FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue), "0")
               strNumberC = IIf(dblCredit, FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue), "0")
               
               bNegativeD = IIf(dblDebit, FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue), 0) < 0
               bNegativeC = IIf(dblCredit, FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue), 0) < 0
               
               If bNegativeD Then
                  strNumberD = Left(strNumberD, Len(strNumberD) - 1)
               End If
               
               If bNegativeC Then
                  strNumberC = Left(strNumberC, Len(strNumberC) - 1)
               End If

               strTemp = "<LINDENT=180>Debit<LINDENT=" & 5325 - 300 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=5625>Credit<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", "")
            End If

            .AddText strTemp, "Arial", IIf(Index = 0, 8, 9), True
            .TextItem(.TextItem.Count).Top = IIf(Index = 0, 300, 165)

            ReDim Preserve intVLineX(UBound(intVLineX) + 1) As Integer
            intVLineX(UBound(intVLineX)) = .TextItem.Count

            If Not lvButtons_H2(1).Value And intmonth > 1 Then
               strSQL = "SELECT SUM(SumOfDebit) AS DebitValue, SUM(SumOfCredit) AS CreditValue  FROM ( "
               strSQL = strSQL & "SELECT SUM(Debit) AS SumOfDebit, 0 AS SumOfCredit "
               strSQL = strSQL & "FROM tblJournal "
               strSQL = strSQL & "WHERE (Date " & strConDate & ") AND DebitCoA = '" & strAccount(intCount) & "' AND Remark = '" & strRemark(intCount) & "' "
               strSQL = strSQL & "UNION ALL "
               strSQL = strSQL & "SELECT 0 AS SumOfDebit, SUM(Credit) AS SumOfCredit "
               strSQL = strSQL & "FROM tblJournal "
               strSQL = strSQL & "WHERE (Date " & strConDate & ") AND CreditCoA = '" & strAccount(intCount) & "' AND Remark = '" & strRemark(intCount) & "');"
            
               CloseConn , Me.hWnd
               ADORS.Open strSQL, ADOCnn
               
               If ADORS.RecordCount > 0 Then
                  If Index = 0 Then
                     If bDebitPos(intCount) Then
                        dblTotalLast = ADORS!DebitValue - ADORS!CreditValue
                        strNumber = FormatNumber(dblTotalLast, 0, vbTrue, vbTrue, vbTrue)
                        bNegative = (dblTotalLast) < 0
                        
                        If bNegative Then
                           strNumber = Left(strNumber, Len(strNumber) - 1)
                        End If
                     
                        strTemp = "<LINDENT=5280>Total Last Month<LINDENT=" & 9150 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     Else
                        dblTotalLast = ADORS!CreditValue - ADORS!DebitValue
                        strNumber = FormatNumber(dblTotalLast, 0, vbTrue, vbTrue, vbTrue)
                        bNegative = (dblTotalLast) < 0
                        
                        If bNegative Then
                           strNumber = Left(strNumber, Len(strNumber) - 1)
                        End If
                     
                        strTemp = "<LINDENT=5280>Total Last Month<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     End If
                     
                     .AddText strTemp, "Arial", 8, True
                     .TextItem(.TextItem.Count).Top = 90
                  End If
               End If
            End If
            
            If Index = 0 Then
               If bDebitPos(intCount) Then
                  dblDebit = dblDebit - dblCredit
                  dblDebit = dblDebit + dblTotalLast
                  
                  strNumber = FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue)
                  bNegative = (dblDebit) < 0
                  
                  If bNegative Then
                     strNumber = Left(strNumber, Len(strNumber) - 1)
                  End If
               
                  strTemp = "<LINDENT=5280>Grand Total<LINDENT=" & 9150 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               Else
                  dblCredit = dblCredit - dblDebit
                  dblCredit = dblCredit + dblTotalLast
                  
                  strNumber = FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue)
                  bNegative = (dblCredit) < 0
                  
                  If bNegative Then
                     strNumber = Left(strNumber, Len(strNumber) - 1)
                  End If
               
                  strTemp = "<LINDENT=5280>Grand Total<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               End If

               .AddText strTemp, "Arial", 8, True
               .TextItem(.TextItem.Count).Top = 90
            End If

            .TextItem(.TextItem.Count).NewPage = After_np
         End If
      Next intCount

      .TextItem(.TextItem.Count).NewPage = None_np

      intPage = .Pages
      
      For intCount = 1 To UBound(intVLineX2)
         .AddLine 180, IIf(intCount = 1, 1020, 210), 10650, IIf(intCount = 1, 1020, 210), .TextItem(intVLineX2(intCount)).StartPage, 3
         .AddLine 180, IIf(intCount = 1, 990, 180), 10650, IIf(intCount = 1, 990, 180), .TextItem(intVLineX2(intCount)).StartPage
         
         intPageStart(intCount) = .TextItem(intVLineX2(intCount)).StartPage

         If Index = 0 Then
            .AddLine 180, IIf(intCount = 1, 1410, 600), 10650, IIf(intCount = 1, 1410, 600), .TextItem(intVLineX2(intCount)).StartPage
         End If
      Next intCount
      
      Dim bSamePage As Boolean
      
      For intCount = 1 To UBound(intVLineX)
         bSamePage = .TextItem(intVLineX(intCount)).StartPage = .TextItem(intVLineX2(intCount)).StartPage
         
         If Index = 1 Then
            .AddLine 180, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 90, 360), 10650, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 90, 360), .TextItem(intVLineX(intCount)).StartPage
            .AddLine 180, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 450, 720), 10650, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 450, 720), .TextItem(intVLineX(intCount)).StartPage, 3
            .AddLine 180, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 480, 750), 10650, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 480, 750), .TextItem(intVLineX(intCount)).StartPage
         Else
            If .TextItem(intVLineX(intCount)).StartPage <> .TextItem(intVLineX(intCount) + 1).StartPage Then
               .TextItem(intVLineX(intCount)).NewPage = Before_np
               intPage = .Pages
            End If

            .AddLine 180, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 180, 450), 10650, .TextItem(intVLineX(intCount)).PositionStartTwip + IIf(bSamePage, 180, 450), .TextItem(intVLineX(intCount)).StartPage
            .AddLine 5280, .TextItem(intVLineX(intCount) + 1).PositionStartTwip + IIf(bSamePage, 390, 660) + IIf(Not lvButtons_H2(1).Value And intmonth > 1, 270, 0), 10650, .TextItem(intVLineX(intCount) + 1).PositionStartTwip + IIf(bSamePage, 390, 660) + IIf(Not lvButtons_H2(1).Value And intmonth > 1, 270, 0), .TextItem(intVLineX(intCount) + 1).StartPage, 3
            .AddLine 5280, .TextItem(intVLineX(intCount) + 1).PositionStartTwip + IIf(bSamePage, 420, 690) + IIf(Not lvButtons_H2(1).Value And intmonth > 1, 270, 0), 10650, .TextItem(intVLineX(intCount) + 1).PositionStartTwip + IIf(bSamePage, 420, 690) + IIf(Not lvButtons_H2(1).Value And intmonth > 1, 270, 0), .TextItem(intVLineX(intCount) + 1).StartPage
         End If
      
         intPageEnd(intCount) = .TextItem(intVLineX(intCount) + IIf(Index = 1, 0, 1)).StartPage
      Next intCount

      .Footer(EvenPage_hf).Text = "Page #pagenumber# of #pagetotal#"
      .Footer(OddPage_hf).Text = "Page #pagenumber# of #pagetotal#"
      .Footer(OddPage_hf).Alignment = eRight
      .Footer(EvenPage_hf).Alignment = eRight
      .Footer(OddPage_hf).FontSize = 7
      .Footer(EvenPage_hf).FontSize = 7
      .SetFooter(OddPage_hf) = True
      .SetFooter(EvenPage_hf) = True
   End With

   Unload frmPrintProgress
   
   Dim intIndex As Integer
   
   ReDim strPage(intPage)
   
   bCustomPage = True
   
   On Error Resume Next
   
   For intCount = 0 To intPage
      strPage(intCount) = strAccount(intIndex)
'      Debug.Print intPageEnd(intIndex + 1) & "-" & intPageStart(intIndex + 1)

      If intPageEnd(intIndex + 1) <> intPageStart(intIndex + 1) Then
         intPageEnd(intIndex + 1) = intPageEnd(intIndex + 1) - 1
      Else
         intIndex = intIndex + 1
      End If
   Next intCount
   
   On Error GoTo 0
End Sub

Private Sub txtYear_Change()
   If IsNumeric(txtYear) Then
      strYear = txtYear

'      Call cboMonth_Click
   Else
      txtYear = strYear
      ShowBalloonTip txtYear.hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
End Sub

Private Function IsDBExist(Optional intmonth As Integer = -1, Optional intYear As Integer = -1) As Boolean
   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblJournal WHERE Date BETWEEN #" & intmonth & "/1/" & intYear & _
              "# AND #" & IIf(lvButtons_H2(1).Value, 12, intmonth) & "/" & MaxDayOfMonth(intYear, IIf(lvButtons_H2(1).Value, 12, intmonth)) & "/" & intYear & "#;", ADOCnn

   IsDBExist = ADORS.RecordCount > 0
End Function

Private Sub UpDown1_DownClick()
   If txtYear > 2050 Then txtYear = 2050
   txtYear = txtYear - 1
   If txtYear < 1990 Then txtYear = 1990
End Sub

Private Sub UpDown1_UpClick()
   If txtYear < 1990 Then txtYear = 1990
   txtYear = txtYear + 1
   If txtYear > 2050 Then txtYear = 2050
End Sub

