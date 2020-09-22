VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmIncome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income Statement - Print Option"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3675
   Icon            =   "frmIncome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   3330
      TabIndex        =   0
      Top             =   285
      Width           =   270
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2287
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   263
      Width           =   1125
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   1635
      TabIndex        =   2
      Top             =   2010
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
      Image           =   "frmIncome.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvButtons_H1 
      Default         =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2010
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
      Image           =   "frmIncome.frx":035E
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Ex 
      Height          =   1035
      Index           =   2
      Left            =   247
      TabIndex        =   5
      Tag             =   "GL's table with 2 columns."
      Top             =   818
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1826
      Caption         =   "Simple"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "frmIncome.frx":0A70
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_Ex 
      Height          =   1035
      Index           =   3
      Left            =   1372
      TabIndex        =   6
      Tag             =   "GL's table for fast view of accounts, because of its simplicity."
      Top             =   818
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1826
      Caption         =   "Detail"
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
      cGradient       =   0
      CapStyle        =   2
      Mode            =   2
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmIncome.frx":16C2
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Report for year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   247
      TabIndex        =   4
      Top             =   308
      Width           =   1185
   End
End
Attribute VB_Name = "frmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strYear As String

Private Sub txtYear_Change()
   If IsNumeric(txtYear) Then
      strYear = txtYear
   Else
      txtYear = strYear
      ShowBalloonTip txtYear.hWnd, "Unacceptable Character", "You can only type a number here!", [2-Information Icon]
   End If
End Sub

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

Private Sub Form_Load()
   txtYear = Year(Date)
   UpDown1.Left = txtYear.Left + txtYear.Width - UpDown1.Width - 15
   SetMenuForm Me.hWnd
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
   Me.Enabled = False
   
   If Index = 1 Then
      Dim MyPrinter As New clsPrinter
      Dim strTemp As String
      Dim dblRevenue As Double
      Dim dblOtherIncomes As Double
      Dim dblExpense As Double
      Dim dblOtherExpenses As Double
      Dim dblTax As Double
      Dim strNumber As String
      Dim bNegative As Boolean
      Dim strIncomes() As String
      Dim dblValuesI() As Double
      Dim strOtherIncomes() As String
      Dim dblValuesOI() As Double
      Dim strExpenses() As String
      Dim dblValuesE() As Double
      Dim strOtherExpenses() As String
      Dim dblValuesOE() As Double
      Dim strSQL As String
      Dim intCount As Integer
      Dim intPage As Integer
      Dim intLine(1 To 4) As Integer
      
      ReDim strIncomes(0)
      ReDim dblValuesI(0)
      ReDim strOtherIncomes(0)
      ReDim dblValuesOI(0)
      ReDim strExpenses(0)
      ReDim dblValuesE(0)
      ReDim strOtherExpenses(0)
      ReDim dblValuesOE(0)
      
      CloseConn
      ADORS.Open "SELECT Sum(Credit)-Sum(Debit) AS AccValue From tblJournal WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Category='Incomes' AND Remark = 'Incomes';", ADOCnn
      
      If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
         dblRevenue = ADORS!accvalue
      End If
      
      CloseConn
      ADORS.Open "SELECT SUM(Credit) - SUM(Debit) AS AccValue FROM tblJournal WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Category = 'Incomes' AND Remark = 'Other Incomes';", ADOCnn
      
      If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
         dblOtherIncomes = ADORS!accvalue
      End If
      
      CloseConn
      ADORS.Open "SELECT SUM(Debit) - SUM(Credit) AS AccValue FROM tblJournal WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Category = 'Expenses' AND NOT Remark = 'Other Expenses';", ADOCnn
      
      If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
         dblExpense = ADORS!accvalue
      End If
      
      CloseConn
      ADORS.Open "SELECT SUM(Debit) - SUM(Credit) AS AccValue FROM tblJournal WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Category = 'Expenses' AND Remark = 'Other Expenses' AND NOT Description = 'Income Taxes';", ADOCnn
      
      If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
         dblOtherExpenses = ADORS!accvalue
      End If
      
      CloseConn
      ADORS.Open "SELECT SUM(Debit) AS AccValue FROM tblJournal WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Category = 'Expenses' AND Remark = 'Other Expenses' AND Description = 'Income Taxes';", ADOCnn
      
      If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
         dblTax = ADORS!accvalue
      End If
      
      If dblRevenue = 0 And dblExpense = 0 And dblOtherIncomes = 0 And dblOtherExpenses = 0 Then
         xMsgBox Me.hWnd, "There are no record! ", vbInformation, "No Record"
         Me.Enabled = True
         Exit Sub
      End If
         
      With MyPrinter
         AddReportLogo frmMain.picLogo, "Income Statement", MyPrinter, qPortrait
         
         frmMain.picLogo.FontBold = True
         frmMain.picLogo.FontSize = 10
         frmMain.picLogo.FontName = "Arial"
         
         .AddText " ", "Arial", 16, True
         .AddText "<LINDENT=" & 10650 - frmMain.picLogo.TextWidth("As of December 31st, " & strYear) & ">As of December 31st, " & strYear, "Arial", 10, True
         .AddText " ", "Arial", 12, True
         
         .AddLine 180, 985, 10650, 985, 1
         .AddLine 180, 1015, 10650, 1015, 1, 3
         .AddLine 180, 1465, 10650, 1465, 1, 2
            
         strTemp = "<LINDENT=180>Description<LINDENT=" & 10650 - frmMain.picLogo.TextWidth("Value (IDR)") & ">Value (IDR)"
         
         .AddText strTemp, "Arial", 10, True
         .AddText " ", "Arial", 16, True
         
         If lvB_Ex(2).Value Then
            frmMain.picLogo.FontSize = 9
            
            strNumber = FormatNumber(dblRevenue, 0, vbTrue, vbTrue, vbTrue)
            bNegative = dblRevenue < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=180>Total Revenue<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 10, True
            
            strNumber = FormatNumber(dblExpense * -1, 0, vbTrue, vbTrue, vbTrue)
            bNegative = (dblExpense * -1) < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=360>Total General & Admin Expenses<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 10, True
            
            strNumber = FormatNumber(dblOtherIncomes, 0, vbTrue, vbTrue, vbTrue)
            bNegative = dblOtherIncomes < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=180>Total Other Incomes<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 10, True
            
            strNumber = FormatNumber(dblOtherExpenses * -1, 0, vbTrue, vbTrue, vbTrue)
            bNegative = (dblOtherExpenses * -1) < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=360>Total Other Expenses<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 10, True
            
            strNumber = FormatNumber((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses), 0, vbTrue, vbTrue, vbTrue)
            bNegative = ((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=180>Total Income Before Tax<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 10, True
            
            strNumber = FormatNumber(dblTax, 0, vbTrue, vbTrue, vbTrue)
            bNegative = dblRevenue < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=360>Income Tax<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 9, True
            .AddText " ", "Arial", 16, True
            
            frmMain.picLogo.FontSize = 10
            
            strNumber = FormatNumber(((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) - dblTax, 0, vbTrue, vbTrue, vbTrue)
            bNegative = (((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) - dblTax) < 0
            
            If bNegative Then
               strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
            
            strTemp = "<LINDENT=180>Net Income (Loss)<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
            .AddText strTemp, "Arial", 10, True
         Else
            strSQL = "SELECT AccName, Remark, Category, SUM(AccDValue) AS AccDebit, SUM(AccCValue) AS AccCredit FROM ("
            strSQL = strSQL & "SELECT DebitCoA AS AccName, Remark, Category, Debit AS AccDValue, 0 AS AccCValue "
            strSQL = strSQL & "From tblJournal "
            strSQL = strSQL & "WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Not (DebitCoA = '' OR IsNull(DebitCoA)) And Category = 'Incomes' Or Category = 'Expenses'"
            strSQL = strSQL & "UNION ALL "
            strSQL = strSQL & "SELECT CreditCoA AS AccName, Remark, Category, 0 AS AccDValue, Credit AS AccCValue "
            strSQL = strSQL & "From tblJournal "
            strSQL = strSQL & "WHERE (Date BETWEEN #1/1/" & strYear & "# AND #12/31/" & strYear & "#) AND Not (CreditCoA = '' OR IsNull(CreditCoA)) And Category = 'Incomes' Or Category = 'Expenses') "
            strSQL = strSQL & "WHERE Not (AccName = '' OR IsNull(AccName)) "
            strSQL = strSQL & "GROUP BY AccName, Remark, Category;"
            
'            Debug.Print strSQL
'            Unload Me
'            Exit Sub

            CloseConn
            ADORS.Open strSQL, ADOCnn
            
            If ADORS.RecordCount > 0 Then
               Do Until ADORS.EOF
                  If ADORS!Category = "Incomes" Then
                     If ADORS!Remark = "Incomes" Then
                        ReDim Preserve strIncomes(UBound(strIncomes) + 1)
                        ReDim Preserve dblValuesI(UBound(dblValuesI) + 1)
                        
                        strIncomes(UBound(strIncomes)) = ADORS!AccName
                        dblValuesI(UBound(dblValuesI)) = ADORS!AccCredit - ADORS!AccDebit
                     Else
                        ReDim Preserve strOtherIncomes(UBound(strOtherIncomes) + 1)
                        ReDim Preserve dblValuesOI(UBound(dblValuesOI) + 1)
                        
                        strOtherIncomes(UBound(strOtherIncomes)) = ADORS!AccName
                        dblValuesOI(UBound(dblValuesOI)) = ADORS!AccCredit - ADORS!AccDebit
                     End If
                  Else
                     If InStr(UCase(ADORS!Remark), UCase("Breakdown of Admin")) > 0 Then
                        ReDim Preserve strExpenses(UBound(strExpenses) + 1)
                        ReDim Preserve dblValuesE(UBound(dblValuesE) + 1)
                        
                        strExpenses(UBound(strExpenses)) = ADORS!AccName
                        dblValuesE(UBound(dblValuesE)) = (ADORS!AccDebit - ADORS!AccCredit) * -1
                     Else
                        ReDim Preserve strOtherExpenses(UBound(strOtherExpenses) + 1)
                        ReDim Preserve dblValuesOE(UBound(dblValuesOE) + 1)
                        
                        strOtherExpenses(UBound(strOtherExpenses)) = ADORS!AccName
                        dblValuesOE(UBound(dblValuesOE)) = (ADORS!AccDebit - ADORS!AccCredit) * -1
                     End If
                  End If
                  
                  ADORS.MoveNext
               Loop
               
               frmMain.picLogo.FontSize = 9
               frmMain.picLogo.FontBold = False
               
               .AddText "<LINDENT=180>Revenues", "Arial", 9, True
               
               If UBound(strIncomes) > 0 Then
                  For intCount = 1 To UBound(strIncomes)
                     strNumber = FormatNumber(dblValuesI(intCount), 0, vbTrue, vbTrue, vbTrue)
                     bNegative = dblValuesI(intCount) < 0
                     
                     If bNegative Then
                        strNumber = Left(strNumber, Len(strNumber) - 1)
                     End If
                     
                     strTemp = "<LINDENT=960>" & strIncomes(intCount) & "<LINDENT=" & 9000 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     
                     .AddText strTemp, "Arial", 9
                     
                     If intCount Mod 2 Then
                        .TextItem(.TextItem.Count).BorderShading = RGB(239, 239, 239)
                        .TextItem(.TextItem.Count).ShowBorder = True
                     End If
                  Next intCount
               End If
                  
               strNumber = FormatNumber(dblRevenue, 0, vbTrue, vbTrue, vbTrue)
               bNegative = dblRevenue < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               frmMain.picLogo.FontBold = True
               strTemp = "<LINDENT=570>Total Revenue<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               frmMain.picLogo.FontBold = False
               
               intLine(1) = .TextItem.Count + 1
               
               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 30
               .AddText " ", "Arial", 8
               .AddText "<LINDENT=180>General & Admin Expenses", "Arial", 9, True
               
               If UBound(strExpenses) > 0 Then
                  For intCount = 1 To UBound(strExpenses)
                     strNumber = FormatNumber(dblValuesE(intCount), 0, vbTrue, vbTrue, vbTrue)
                     bNegative = dblValuesE(intCount) < 0
                     
                     If bNegative Then
                        strNumber = Left(strNumber, Len(strNumber) - 1)
                     End If
                     
                     strTemp = "<LINDENT=960>" & strExpenses(intCount) & "<LINDENT=" & 9000 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     
                     .AddText strTemp, "Arial", 9
                     
                     If intCount Mod 2 Then
                        .TextItem(.TextItem.Count).BorderShading = RGB(239, 239, 239)
                        .TextItem(.TextItem.Count).ShowBorder = True
                     End If
                  Next intCount
               End If
               
               strNumber = FormatNumber(dblExpense * -1, 0, vbTrue, vbTrue, vbTrue)
               bNegative = (dblExpense * -1) < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               frmMain.picLogo.FontBold = True
               strTemp = "<LINDENT=570>Total General & Admin Expenses<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               frmMain.picLogo.FontBold = False
               
               intLine(2) = .TextItem.Count + 1
               
               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 30
               .AddText " ", "Arial", 8
               .AddText "<LINDENT=180>Other Incomes", "Arial", 9, True
               
               If UBound(strOtherIncomes) > 0 Then
                  For intCount = 1 To UBound(strOtherIncomes)
                     strNumber = FormatNumber(dblValuesOI(intCount), 0, vbTrue, vbTrue, vbTrue)
                     bNegative = dblValuesOI(intCount) < 0
                     
                     If bNegative Then
                        strNumber = Left(strNumber, Len(strNumber) - 1)
                     End If
                     
                     strTemp = "<LINDENT=960>" & strOtherIncomes(intCount) & "<LINDENT=" & 9000 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     
                     .AddText strTemp, "Arial", 9
                     
                     If intCount Mod 2 Then
                        .TextItem(.TextItem.Count).BorderShading = RGB(239, 239, 239)
                        .TextItem(.TextItem.Count).ShowBorder = True
                     End If
                  Next intCount
               End If
                  
               frmMain.picLogo.FontBold = True
               strNumber = FormatNumber(dblOtherIncomes, 0, vbTrue, vbTrue, vbTrue)
               frmMain.picLogo.FontBold = False
               
               bNegative = dblOtherIncomes < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               strTemp = "<LINDENT=570>Total Other Incomes<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               
               intLine(3) = .TextItem.Count + 1
               
               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 30
               .AddText " ", "Arial", 8
               .AddText "<LINDENT=180>Other Expenses", "Arial", 9, True
               
               If UBound(strOtherExpenses) > 0 Then
                  For intCount = 1 To UBound(strOtherExpenses)
                     strNumber = FormatNumber(dblValuesOE(intCount), 0, vbTrue, vbTrue, vbTrue)
                     bNegative = dblValuesOE(intCount) < 0
                     
                     If bNegative Then
                        strNumber = Left(strNumber, Len(strNumber) - 1)
                     End If
                     
                     strTemp = "<LINDENT=960>" & strOtherExpenses(intCount) & "<LINDENT=" & 9000 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
                     
                     .AddText strTemp, "Arial", 9
                     
                     If intCount Mod 2 Then
                        .TextItem(.TextItem.Count).BorderShading = RGB(239, 239, 239)
                        .TextItem(.TextItem.Count).ShowBorder = True
                     End If
                  Next intCount
               End If
                  
               strNumber = FormatNumber(dblOtherExpenses * -1, 0, vbTrue, vbTrue, vbTrue)
               bNegative = (dblOtherExpenses * -1) < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               frmMain.picLogo.FontBold = True
               strTemp = "<LINDENT=570>Total Other Expenses<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
            
               intLine(4) = .TextItem.Count + 1
               
               .AddText strTemp, "Arial", 9, True
               .TextItem(.TextItem.Count).Top = 30
               .AddText " ", "Arial", 8
               
               strNumber = FormatNumber((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses), 0, vbTrue, vbTrue, vbTrue)
               bNegative = ((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               strTemp = "<LINDENT=180>Total Income Before Tax<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")  'strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               
               .AddText strTemp, "Arial", 9, True
               .AddText " ", "Arial", 8, True
               
               strNumber = FormatNumber(dblTax, 0, vbTrue, vbTrue, vbTrue)
               bNegative = dblRevenue < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               strTemp = "<LINDENT=180>Income Tax<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               
               .AddText strTemp, "Arial", 9, True
               .AddText " ", "Arial", 16, True
               
               frmMain.picLogo.FontSize = 10
               
               strNumber = FormatNumber(((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) - dblTax, 0, vbTrue, vbTrue, vbTrue)
               bNegative = (((dblRevenue + dblOtherIncomes) - (dblExpense + dblOtherExpenses)) - dblTax) < 0
               
               If bNegative Then
                  strNumber = Left(strNumber, Len(strNumber) - 1)
               End If
               
               strTemp = "<LINDENT=180>Net Income (Loss)<LINDENT=" & 10650 - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "")
               
               .AddText strTemp, "Arial", 10, True
            End If
         End If
         
         intPage = .Pages
         
         If lvB_Ex(3).Value Then
            For intCount = 1 To 4
               .AddLine 570, .TextItem(intLine(intCount)).PositionStartTwip + 15, 10650, .TextItem(intLine(intCount)).PositionStartTwip + 15, .TextItem(intLine(intCount)).StartPage
            Next
         End If
         
         .AddLine 180, .TextItem(.TextItem.Count).PositionStartTwip - 120, 10650, .TextItem(.TextItem.Count).PositionStartTwip - 120, .TextItem(.TextItem.Count).StartPage, 2
         .AddLine 180, .TextItem(.TextItem.Count).PositionEndTwip + 120, 10650, .TextItem(.TextItem.Count).PositionEndTwip + 120, .TextItem(.TextItem.Count).StartPage, 3
         .AddLine 180, .TextItem(.TextItem.Count).PositionEndTwip + 150, 10650, .TextItem(.TextItem.Count).PositionEndTwip + 150, .TextItem(.TextItem.Count).StartPage
         
         .Preview Me, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
      End With
      
      Set MyPrinter = Nothing
      Me.Enabled = True
   Else
      Unload Me
   End If
End Sub
