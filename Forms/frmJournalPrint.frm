VERSION 5.00
Begin VB.Form frmJournalPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Journal - Print Option"
   ClientHeight    =   2835
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJournalPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.ctlDatePicker DateTo 
      Height          =   315
      Left            =   2790
      TabIndex        =   5
      Top             =   1590
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Enabled         =   0   'False
   End
   Begin Project1.ctlDatePicker DateFrom 
      Height          =   315
      Left            =   255
      TabIndex        =   4
      Top             =   1590
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Enabled         =   0   'False
   End
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   495
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Tag             =   "Report records for this month only."
      Top             =   637
      Width           =   1665
      _ExtentX        =   2937
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
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   495
      Index           =   1
      Left            =   3315
      TabIndex        =   1
      Tag             =   "Report whole year records."
      Top             =   637
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   873
      Caption         =   "This Year"
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
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   495
      Index           =   2
      Left            =   1485
      TabIndex        =   2
      Tag             =   "Report records for another month."
      Top             =   637
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   873
      Caption         =   "Date in Range"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
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
      Left            =   2415
      TabIndex        =   8
      Top             =   2130
      Width           =   2550
      _ExtentX        =   4498
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
      Image           =   "frmJournalPrint.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   255
      TabIndex        =   9
      Top             =   2130
      Width           =   2505
      _ExtentX        =   4419
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
      Image           =   "frmJournalPrint.frx":035E
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2865
      TabIndex        =   7
      Top             =   1305
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date From :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   6
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Select date range :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1928
      TabIndex        =   3
      Top             =   255
      Width           =   1365
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   "Print{IMG:I14}"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close{IMG:I10}"
      End
   End
End
Attribute VB_Name = "frmJournalPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strDate As String
Private strDateTitle As String

Private Sub DateFrom_Change()
   strDate = "Date BETWEEN #" & DateFrom.DTPMonth & "/" & DateFrom.DTPDay & "/" & DateFrom.DTPYear & "# AND #" & DateTo.DTPMonth & "/" & DateTo.DTPDay & "/" & DateTo.DTPYear & "#"
   strDateTitle = GetMonthName(DateFrom.DTPMonth) & " " & DateFrom.DTPDay & GetOrdinal(DateFrom.DTPDay) & ", " & DateFrom.DTPYear & " to " & GetMonthName(DateTo.DTPMonth) & " " & DateTo.DTPDay & GetOrdinal(DateTo.DTPDay) & ", " & DateTo.DTPYear
   
   If DateFrom.DTPValue > DateTo.DTPValue Then
      DateTo.DTPValue = DateFrom.DTPValue
   End If
End Sub

Private Sub DateTo_Change()
   strDate = "Date BETWEEN #" & DateFrom.DTPMonth & "/" & DateFrom.DTPDay & "/" & DateFrom.DTPYear & "# AND #" & DateTo.DTPMonth & "/" & DateTo.DTPDay & "/" & DateTo.DTPYear & "#"
   strDateTitle = GetMonthName(DateFrom.DTPMonth) & " " & DateFrom.DTPDay & GetOrdinal(DateFrom.DTPDay) & ", " & DateFrom.DTPYear & " to " & GetMonthName(DateTo.DTPMonth) & " " & DateTo.DTPDay & GetOrdinal(DateTo.DTPDay) & ", " & DateTo.DTPYear
   
   If DateFrom.DTPValue > DateTo.DTPValue Then
      DateFrom.DTPValue = DateTo.DTPValue
   End If
End Sub

Private Sub Form_Load()
   DateFrom.DTPValue = Date
   DateTo.DTPValue = Date
   DateFrom.DTPDay = 1
   DateTo.DTPDay = MaxDayOfMonth(Year(Date), Month(Date))
   
   strDate = "Date BETWEEN #" & Month(Date) & "/1/" & Year(Date) & "# AND #" & Month(Date) & "/" & MaxDayOfMonth(Year(Date), Month(Date)) & "/" & Year(Date) & "#"
   strDateTitle = GetMonthName(Month(Date)) & ", " & Year(Date)

   SetMenuForm Me.hWnd, frmMain.IL1
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_OK_Click()
   Me.Enabled = False
   Call mnuPrint_Click
   Me.Enabled = True
End Sub

Private Sub lvButtons_H2_Click(Index As Integer)
   Select Case Index
      Case 2
         DateFrom.Enabled = True
         DateTo.Enabled = True
         strDate = "Date BETWEEN #" & DateFrom.DTPMonth & "/" & DateFrom.DTPDay & "/" & DateFrom.DTPYear & "# AND #" & DateTo.DTPMonth & "/" & DateTo.DTPDay & "/" & DateTo.DTPYear & "#"
         strDateTitle = GetMonthName(DateFrom.DTPMonth) & " " & DateFrom.DTPDay & GetOrdinal(DateFrom.DTPDay) & ", " & DateFrom.DTPYear & " to " & GetMonthName(DateTo.DTPMonth) & " " & DateTo.DTPDay & GetOrdinal(DateTo.DTPDay) & ", " & DateTo.DTPYear
         
      Case 1
         DateFrom.Enabled = False
         DateTo.Enabled = False
         strDate = "Date BETWEEN #1/1/" & Year(Date) & "# AND #12/31/" & Year(Date) & "#"
         strDateTitle = "January 1st, " & Year(Date) & " to December 31st, " & Year(Date)
         
      Case 0
         DateFrom.Enabled = False
         DateTo.Enabled = False
         strDate = "Date BETWEEN #" & Month(Date) & "/1/" & Year(Date) & "# AND #" & Month(Date) & "/" & MaxDayOfMonth(Year(Date), Month(Date)) & "/" & Year(Date) & "#"
         strDateTitle = GetMonthName(Month(Date)) & ", " & Year(Date)
         
   End Select
End Sub

Private Sub mnuClose_Click()
   Unload Me
End Sub

Private Sub mnuPrint_Click()
   Const strProgress As String = "Printing, please wait ...."
   
   Dim MyPrinter As New clsPrinter
   Dim strTemp As String
   Dim strTmp2 As String
   Dim strDateTmp As String
   Dim strCoAString As String
   Dim strDescription As String
   Dim strCustomer As String
   Dim dblDebit As Double
   Dim dblCredit As Double
   Dim intPage As Integer
   Dim bWithLine() As Boolean
   Dim intCount As Integer
   Dim int1stItem As Integer
   Dim intLastItem As Integer
   Dim intProgress As Integer
   Dim strNumber As String
   Dim strNumberC As String
   Dim strNumberD As String
   Dim bNegative As Boolean
   Dim bNegativeC As Boolean
   Dim bNegativeD As Boolean
   
   frmPrintProgress.Show vbModeless, Me
   intProgress = 1
   
   ReDim bWithLine(0)
   
   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM [SELECT tblJournal.IDJournal, tblJournal.Voucher, tblJournal.Date, tblJournal.DebitCoA AS CoAString, tblJournal.Debit AS CoAValue, tblJournal.Description, tblJournal.Customer, tblJournal.Notes, 'Debit' AS Flag FROM tblJournal WHERE IsNull(tblJournal.DebitCoA) = False UNION ALL SELECT tblJournal.IDJournal, tblJournal.Voucher, tblJournal.Date,  tblJournal.CreditCoA AS CoAString, tblJournal.Credit AS CoAValue, tblJournal.Description, tblJournal.Customer, tblJournal.Notes, 'Credit' AS Flag FROM tblJournal WHERE IsNull(tblJournal.CreditCoA) = False]. AS [%$##@_Alias] WHERE " & strDate & " ORDER BY Date, Voucher, IDJournal;", ADOCnn
   
   If ADORS.RecordCount < 1 Then
      Unload frmPrintProgress
      xMsgBox Me.hWnd, "There are no record! ", vbInformation, "No Records"
      Exit Sub
   End If
   
   With MyPrinter
      AddReportLogo frmMain.picLogo, "Journal", MyPrinter
      
      .MarginTop = 1200
      .TextItem(.TextItem.Count).Top = -690
      
      .AddLine 180, 90, 14850, 90, 1, 3
      .AddLine 180, 60, 14850, 60, 1, 1
      .AddLine 180, 480, 14850, 480, 1, 1

      frmMain.picLogo.FontName = "Arial"
      frmMain.picLogo.FontBold = True
      frmMain.picLogo.FontSize = 10
      
      strTemp = "<LINDENT=" & 14850 - frmMain.picLogo.TextWidth(strDateTitle) & ">" & strDateTitle
      
      .AddText strTemp, "Arial", 10, True
      .TextItem(.TextItem.Count).Top = 150
      
      strTemp = "<LINDENT=180>Date<LINDENT=1200>Slip No<LINDENT=2550>Account<LINDENT=4950>Description<LINDENT=" & 9750 - frmMain.picLogo.TextWidth("Debit") & ">Debit<LINDENT=" & 11550 - frmMain.picLogo.TextWidth("Credit") & ">Credit<LINDENT=11850>Customer"
      
      .AddText " ", "Arial", 24, True
      .AddText strTemp, "Arial", 10, True
      .TextItem(.TextItem.Count).Top = -375
      int1stItem = .TextItem.Count + 1
      
      strTemp = vbNullString
      
      frmMain.picLogo.FontBold = False
      frmMain.picLogo.FontSize = 9
      
      dblDebit = 0
      dblCredit = 0
      
      Do Until ADORS.EOF
         frmPrintProgress.Label1 = strProgress & vbCrLf & Round((intProgress / ADORS.RecordCount) * 100) & "% complete."
         DoEvents
         
         strCoAString = IIf(ADORS!Flag = "Debit", "", "   ") & ADORS!CoAString
         strDescription = IIf(strTmp2 <> ADORS!Voucher, ADORS!Description, "  ")
         strCustomer = IIf(IsNull(ADORS!Customer) Or ADORS!Customer = vbNullString, "-", ADORS!Customer)
         
         strCoAString = TrimmedString(frmMain.picLogo, strCoAString, 2250)
         strDescription = TrimmedString(frmMain.picLogo, strDescription, 2940)
         strCustomer = TrimmedString(frmMain.picLogo, strCustomer, 3000)
         
         If ADORS!Flag = "Debit" Then
            dblDebit = dblDebit + CDbl(ADORS!CoAValue)
         Else
            dblCredit = dblCredit + CDbl(ADORS!CoAValue)
         End If
         
         If strTmp2 <> ADORS!Voucher Then
            .AddText strTemp, "Arial", 9
            strTemp = vbNullString
            ReDim Preserve bWithLine(.TextItem.Count)
            bWithLine(.TextItem.Count) = True
            .TextItem(.TextItem.Count).Top = 75
         End If
         
'         intPage = .Pages
'
'         If .TextItem(.TextItem.Count).StartPage <> .TextItem(.TextItem.Count).EndPage Then
'            .TextItem(.TextItem.Count).NewPage = Before_np
'         End If

         strDateTmp = Month(ADORS!Date) & "/" & Day(ADORS!Date) & "/" & Right(Year(ADORS!Date), 2)
         
         strNumber = FormatNumber(ADORS!CoAValue, 0, vbTrue, vbTrue, vbTrue)
         bNegative = (ADORS!CoAValue) < 0
         
         If bNegative Then
            strNumber = Left(strNumber, Len(strNumber) - 1)
         End If
               
         strTemp = strTemp & "<LINDENT=180>" & IIf(strTmp2 <> ADORS!Voucher, strDateTmp, "  ") & "<LINDENT=1200>" & IIf(strTmp2 <> ADORS!Voucher, ADORS!Voucher, "  ") & "<LINDENT=2550>" & strCoAString & "<LINDENT=4950>" & strDescription & "<LINDENT=" & IIf(ADORS!Flag = "Debit", 9750, 11550) - frmMain.picLogo.TextWidth(strNumber) & ">" & strNumber & IIf(bNegative, ")", "") & "<LINDENT=11850>" & strCustomer & vbCrLf
         strTmp2 = ADORS!Voucher
      
         ADORS.MoveNext
         intProgress = intProgress + 1
      Loop
      
      .AddText strTemp, "Arial", 9
      
      ReDim Preserve bWithLine(.TextItem.Count)
      
      bWithLine(.TextItem.Count) = True
      .TextItem(.TextItem.Count).Top = 75
      
      .TextItem(int1stItem).Top = -90
      
      bWithLine(int1stItem) = False
      bWithLine(.TextItem.Count) = False
      
      intPage = .Pages
      
      For intCount = 5 To .TextItem.Count
         .TextItem(intCount).NewPage = None_np
      Next intCount
      
      For intCount = 5 To .TextItem.Count
         frmPrintProgress.Label1 = strProgress & vbCrLf & "Formatting table, " & Round((intCount / .TextItem.Count) * 100) & "% complete."
         DoEvents
'         intPage = .Pages
      
         If .TextItem(intCount).StartPage <> .TextItem(intCount).EndPage Or .TextItem(intCount).PositionEndTwip > 9765 Then
            .TextItem(intCount).NewPage = None_np
            .TextItem(intCount).NewPage = Before_np
            intPage = .Pages
         End If
      Next intCount
      
      intLastItem = .TextItem.Count
      
      frmMain.picLogo.FontBold = True
      
      intPage = .Pages
      
      If intPage > 1 Then
         For intPage = 2 To intPage
            .AddLine 180, -450, 14850, -450, intPage, 3
            .AddLine 180, -480, 14850, -480, intPage, 1
            .AddLine 180, -60, 14850, -60, intPage, 1

            frmMain.picLogo.FontBold = True
            frmMain.picLogo.FontSize = 10

            strTemp = "<LINDENT=180>Date<LINDENT=1200>Slip No<LINDENT=2550>Account<LINDENT=4950>Description<LINDENT=" & 9750 - frmMain.picLogo.TextWidth("Debit") & ">Debit<LINDENT=" & 11550 - frmMain.picLogo.TextWidth("Credit") & ">Credit<LINDENT=11850>Customer"
            
            .AddText strTemp, "Arial", 10, True
            .TextItem(.TextItem.Count).Top = -360
            .TextItem(.TextItem.Count).AbsPage = intPage
            .TextItem(.TextItem.Count).Absolute = True

            frmMain.picLogo.FontBold = False
            frmMain.picLogo.FontSize = 9
         Next intPage
      End If
      
      For intCount = 1 To UBound(bWithLine)
         If bWithLine(intCount) Then
            .AddLine 180, .TextItem(intCount).PositionEndTwip + 30, 14850, .TextItem(intCount).PositionEndTwip + 30, .TextItem(intCount).EndPage, 1
         End If
      Next intCount
      
      .AddLine 180, .TextItem(intLastItem).PositionEndTwip + 195, 14850, .TextItem(intLastItem).PositionEndTwip + 195, .TextItem(intLastItem).StartPage, 1
      .AddLine 180, .TextItem(intLastItem).PositionEndTwip + 165, 14850, .TextItem(intLastItem).PositionEndTwip + 165, .TextItem(intLastItem).StartPage, 3

'      .AddLine 1035, .TextItem(.TextItem.Count).PositionEndTwip + 195, 1035, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'      .AddLine 2385, .TextItem(.TextItem.Count).PositionEndTwip + 195, 2385, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'      .AddLine 4785, .TextItem(.TextItem.Count).PositionEndTwip + 195, 4785, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'      .AddLine 8085, .TextItem(.TextItem.Count).PositionEndTwip + 315, 8085, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'      .AddLine 9885, .TextItem(.TextItem.Count).PositionEndTwip + 315, 9885, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'      .AddLine 11685, .TextItem(.TextItem.Count).PositionEndTwip + 315, 11685, .TextItem(.TextItem.Count).PositionEndTwip, intPage, 1
'
'      .AddLine 1065, .TextItem(.TextItem.Count).PositionEndTwip + 195, 1065, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
'      .AddLine 2415, .TextItem(.TextItem.Count).PositionEndTwip + 195, 2415, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
'      .AddLine 4815, .TextItem(.TextItem.Count).PositionEndTwip + 195, 4815, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
'      .AddLine 8115, .TextItem(.TextItem.Count).PositionEndTwip + 255, 8115, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
'      .AddLine 9915, .TextItem(.TextItem.Count).PositionEndTwip + 255, 9915, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
'      .AddLine 11715, .TextItem(.TextItem.Count).PositionEndTwip + 255, 11715, .TextItem(.TextItem.Count).PositionEndTwip - 60, intPage, 1
      
      strNumberD = FormatNumber(dblDebit, 0, vbTrue, vbTrue, vbTrue)
      strNumberC = FormatNumber(dblCredit, 0, vbTrue, vbTrue, vbTrue)
      
      bNegativeD = dblDebit < 0
      bNegativeC = dblCredit < 0
      
      If bNegativeD Then
         strNumberD = Left(strNumberD, Len(strNumberD) - 1)
      End If
      
      If bNegativeC Then
         strNumberC = Left(strNumberC, Len(strNumberC) - 1)
      End If
      
      .AddText "<LINDENT=4950>Total<LINDENT=" & 9750 - frmMain.picLogo.TextWidth(strNumberD) & ">" & strNumberD & IIf(bNegativeD, ")", "") & "<LINDENT=" & 11550 - frmMain.picLogo.TextWidth(strNumberC) & ">" & strNumberC & IIf(bNegativeC, ")", ""), "Arial", 9, True
      .TextItem(.TextItem.Count).Top = 270
      
      .Footer(EvenPage_hf).Text = "Page #pagenumber# of #pagetotal#"
      .Footer(OddPage_hf).Text = "Page #pagenumber# of #pagetotal#"
      .Footer(OddPage_hf).Alignment = eRight
      .Footer(EvenPage_hf).Alignment = eRight
      .Footer(OddPage_hf).FontSize = 7
      .Footer(EvenPage_hf).FontSize = 7
      .SetFooter(OddPage_hf) = True
      .SetFooter(EvenPage_hf) = True

      Unload frmPrintProgress
      
      .Preview Me, 120 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
   End With
   
   Set MyPrinter = Nothing
End Sub

Private Function TrimmedString(picRuler As PictureBox, strNormal As String, sngWidth As Single) As String
   Dim strTemp As String
   Dim sngFontSize As Integer
   
   strTemp = strNormal
   
   If picRuler.TextWidth(strTemp) > sngWidth Then
      sngFontSize = picRuler.FontSize
      picRuler.FontSize = 8
   
      If picRuler.TextWidth(strTemp) > sngWidth Then
         Do Until picRuler.TextWidth(strTemp & "...") < sngWidth
            strTemp = Left(strTemp, Len(strTemp) - 1)
         Loop
         
         TrimmedString = "<SIZE=8>" & strTemp & "...</SIZE>"
      Else
         TrimmedString = "<SIZE=8>" & strTemp & "</SIZE>"
      End If
   
      picRuler.FontSize = sngFontSize
   Else
      TrimmedString = strTemp
   End If
   
End Function
