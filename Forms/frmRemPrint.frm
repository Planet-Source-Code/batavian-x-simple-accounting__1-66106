VERSION 5.00
Begin VB.Form frmRemPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Option"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   Icon            =   "frmRemPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboRemark 
      Height          =   315
      Left            =   323
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   262
      Width           =   3315
   End
   Begin VB.ComboBox cboMonth 
      Enabled         =   0   'False
      Height          =   315
      Left            =   323
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1687
      Width           =   3315
   End
   Begin Project1.lvButtons_H lvB_Month 
      Height          =   495
      Index           =   0
      Left            =   323
      TabIndex        =   1
      Tag             =   "Report records for this month only."
      Top             =   885
      Width           =   1755
      _ExtentX        =   3096
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
      Left            =   1628
      TabIndex        =   2
      Tag             =   "Report records for another month."
      Top             =   885
      Width           =   2010
      _ExtentX        =   3545
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
      Left            =   1718
      TabIndex        =   3
      Top             =   2310
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "frmRemPrint.frx":000C
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H lvB_OK 
      Default         =   -1  'True
      Height          =   405
      Left            =   323
      TabIndex        =   4
      Top             =   2310
      Width           =   1755
      _ExtentX        =   3096
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
      Image           =   "frmRemPrint.frx":035E
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmRemPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strRemark As String

Private strMonth() As String

Private Sub cboRemark_Click()
   strRemark = cboRemark.List(cboRemark.ListIndex)
   
   Me.Caption = strRemark & " - Print Option"
End Sub

Private Sub Form_Load()
   Me.Caption = strRemark & " - Print Option"
   
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
   SetMenuForm Me.hWnd
End Sub

Private Sub lvB_Cancel_Click()
   Unload Me
End Sub

Private Sub lvB_Month_Click(Index As Integer)
   cboMonth.Enabled = lvB_Month(1).Value
   
   If lvB_Month(0).Value Then
      cboMonth.ListIndex = Month(Date) - 1
   End If
End Sub

Private Sub lvB_OK_Click()
   Me.Enabled = False
   
   Const strProgress As String = "Printing, please wait ...."
   
   Dim MyPrinter As New clsPrinter
   Dim strItem() As String
   Dim strItemShadow() As String
   Dim bDebitPos() As Boolean
   Dim strTemp As String
   Dim strDate As String
   Dim strContraDate As String
   Dim intmonth As Integer
   Dim intCount As Long
   Dim dblPrev As Double
   Dim dblTotal As Double
   Dim intPage As Integer
   Dim dblGrTotalPrev As Double
   Dim dblGrTotal As Double
   Dim intAccCount As Integer
   Dim strNumberP As String
   Dim strNumberN As String
   Dim strNumberT As String
   Dim bNegativeP As Boolean
   Dim bNegativeN As Boolean
   Dim bNegativeT As Boolean
   
   intmonth = cboMonth.ListIndex + 1
   
   frmPrintProgress.Show vbModeless, Me
   
   AddReportLogo frmMain.picLogo, strRemark, MyPrinter, qPortrait
   
   strDate = "Date BETWEEN #1/1/" & Year(Date) & "# AND #" & intmonth & "/" & MaxDayOfMonth(Year(Date), intmonth) & "/" & Year(Date) & "# "
   
   intmonth = cboMonth.ListIndex
   
   If intmonth > 0 Then
      strContraDate = "Date BETWEEN #1/1/" & Year(Date) & "# AND #" & intmonth & "/" & MaxDayOfMonth(Year(Date), intmonth) & "/" & Year(Date) & "# "
   End If
   
   CloseConn
   ADORS.Open "SELECT DISTINCT AccName FROM (SELECT Date, CreditCoA AS AccName, Remark FROM tblJournal WHERE ISNULL(CreditCoA) = False UNION ALL SELECT Date, DebitCoA AS AccName, Remark FROM tblJournal WHERE ISNULL(DebitCoA) = False) WHERE Remark = '" & strRemark & "';", ADOCnn
   
'   Debug.Print "SELECT DISTINCT AccName FROM (SELECT Date, CreditCoA AS AccName, Remark FROM tblJournal WHERE ISNULL(CreditCoA) = False UNION ALL SELECT Date, DebitCoA AS AccName, Remark FROM tblJournal WHERE ISNULL(DebitCoA) = False) WHERE Remark = '" & strRemark & "';"
   
   If ADORS.RecordCount < 1 Then
      Unload frmPrintProgress
      xMsgBox Me.hWnd, "There are no record! ", vbInformation, "No Record"
      Me.Enabled = True
      Exit Sub
   End If
   
   Do Until ADORS.EOF
      strTemp = strTemp & "|" & ADORS!AccName
      ADORS.MoveNext
   Loop
   
   strTemp = Right(strTemp, Len(strTemp) - 1)
   strItemShadow() = Split(strTemp, "|")
   
   ReDim strItem(UBound(strItemShadow))
   ReDim bDebitPos(UBound(strItemShadow))
   
'   Debug.Print UBound(strItem)
'   Debug.Print UBound(strItemShadow)
   
'   Unload frmPrintProgress
'   Set MyPrinter = Nothing
'   Exit Sub
   
   CloseConn
   ADORS.Open "SELECT * FROM tblCoA WHERE Remark = '" & strRemark & "' ORDER BY CoA", ADOCnn
   
   Do Until ADORS.EOF
      For intCount = 0 To UBound(strItemShadow)
         If strItemShadow(intCount) = ADORS!Description Then
            strItem(intAccCount) = ADORS!Description
            bDebitPos(intAccCount) = ADORS!IsDebt
            intAccCount = intAccCount + 1
            Exit For
         End If
      Next intCount
      
      ADORS.MoveNext
   Loop
      
   intmonth = cboMonth.ListIndex + 1
   
   strDate = "Date BETWEEN #" & intmonth & "/1/" & Year(Date) & "# AND #" & intmonth & "/" & MaxDayOfMonth(Year(Date), intmonth) & "/" & Year(Date) & "# "
   
   With frmMain.picLogo
      .FontBold = True
      .FontSize = 10
      .FontName = "Arial"
      .ScaleMode = vbTwips

      MyPrinter.AddText " ", "Arial", 19
      MyPrinter.AddText "<LINDENT=" & 10650 - .TextWidth(cboMonth & ", " & Year(Date)) & ">" & cboMonth & ", " & Year(Date), "Arial", 10, True
      MyPrinter.AddText " ", "Arial", 8
      
      .FontSize = 9
      
      strTemp = "<LINDENT=180>Account Name<LINDENT=" & 5650 - .TextWidth("Total Last Month") & ">Total Last Month<LINDENT=" & 8150 - .TextWidth("Total This Month") & ">Total This Month<LINDENT=" & 10650 - .TextWidth("Total") & ">Total"
      
      MyPrinter.AddText strTemp, "Arial", 9, True
      MyPrinter.AddText "<LINDENT=0> ", "Arial", 6
      
      MyPrinter.AddLine 180, 990, 10650, 990, 1, 3
      MyPrinter.AddLine 180, 960, 10650, 960, 1
      MyPrinter.AddLine 180, 1320, 10650, 1320, 1
      
      For intCount = 0 To UBound(strItem)
         frmPrintProgress.Label1 = strProgress & vbCrLf & Round((intCount + 1) / (UBound(strItem) + 1) * 100) & "% complete."
         DoEvents
         
         If strContraDate <> vbNullString Then
            CloseConn
            
            If bDebitPos(intCount) Then
               ADORS.Open "SELECT SUM(Debit)-SUM(Credit) AS AccValue FROM tblJournal WHERE " & strContraDate & " AND (DebitCoA = '" & strItem(intCount) & "' OR CreditCoA = '" & strItem(intCount) & "') AND Remark = '" & strRemark & "';", ADOCnn
            Else
               ADORS.Open "SELECT SUM(Credit)-SUM(Debit) AS AccValue FROM tblJournal WHERE " & strContraDate & " AND (DebitCoA = '" & strItem(intCount) & "' OR CreditCoA = '" & strItem(intCount) & "') AND Remark = '" & strRemark & "';", ADOCnn
            End If
            
            If Not IsNull(ADORS!accvalue) Then
               dblPrev = ADORS!accvalue
               dblGrTotalPrev = dblGrTotalPrev + dblPrev
            Else
               dblPrev = 0
            End If
         End If
         
         CloseConn
         
         If bDebitPos(intCount) Then
            ADORS.Open "SELECT SUM(Debit)-SUM(Credit) AS AccValue FROM tblJournal WHERE " & strDate & " AND (DebitCoA = '" & strItem(intCount) & "' OR CreditCoA = '" & strItem(intCount) & "') AND Remark = '" & strRemark & "';", ADOCnn
         Else
            ADORS.Open "SELECT SUM(Credit)-SUM(Debit) AS AccValue FROM tblJournal WHERE " & strDate & " AND (DebitCoA = '" & strItem(intCount) & "' OR CreditCoA = '" & strItem(intCount) & "') AND Remark = '" & strRemark & "';", ADOCnn
         End If
         
         dblTotal = IIf(IsNull(ADORS!accvalue), 0, ADORS!accvalue)
         dblGrTotal = dblGrTotal + dblTotal
         
         .FontBold = False
         
         strNumberP = FormatNumber(dblPrev, 0, vbTrue, vbTrue, vbTrue)
         strNumberN = FormatNumber(dblTotal, 0, vbTrue, vbTrue, vbTrue)
         strNumberT = FormatNumber(dblPrev + dblTotal, 0, vbTrue, vbTrue, vbTrue)
         
         bNegativeP = dblPrev < 0
         bNegativeN = dblTotal < 0
         bNegativeT = (dblPrev + dblTotal) < 0
         
         If bNegativeP Then
            strNumberP = Left(strNumberP, Len(strNumberP) - 1)
         End If
         
         If bNegativeN Then
            strNumberN = Left(strNumberN, Len(strNumberN) - 1)
         End If
         
         If bNegativeT Then
            strNumberT = Left(strNumberT, Len(strNumberT) - 1)
         End If
      
         strTemp = "<LINDENT=180>" & strItem(intCount) & "<LINDENT=" & 5650 - .TextWidth(strNumberP) & ">" & strNumberP & IIf(bNegativeP, ")", "") & "<LINDENT=" & 8150 - .TextWidth(strNumberN) & ">" & strNumberN & IIf(bNegativeN, ")", "") & "<LINDENT=" & 10650 - .TextWidth(strNumberT) & ">" & strNumberT & IIf(bNegativeT, ")", "")
         
         MyPrinter.AddText strTemp, "Arial", 9
         MyPrinter.TextItem(MyPrinter.TextItem.Count).Top = 30
         
         If intCount Mod 2 <> 0 Then
            MyPrinter.TextItem(MyPrinter.TextItem.Count).BorderShading = RGB(239, 239, 239)
            MyPrinter.TextItem(MyPrinter.TextItem.Count).BorderLine = 0
            MyPrinter.TextItem(MyPrinter.TextItem.Count).ShowBorder = True
         End If
      Next intCount
      
      MyPrinter.AddText " ", "Arial", 8
      
      .FontBold = True
      .FontSize = 9
      .FontName = "Arial"
      .ScaleMode = vbTwips
   
      strNumberP = FormatNumber(dblGrTotalPrev, 0, vbTrue, vbTrue, vbTrue)
      strNumberN = FormatNumber(dblGrTotal, 0, vbTrue, vbTrue, vbTrue)
      strNumberT = FormatNumber(dblGrTotalPrev + dblGrTotal, 0, vbTrue, vbTrue, vbTrue)
      
      bNegativeP = dblGrTotalPrev < 0
      bNegativeN = dblGrTotal < 0
      bNegativeT = (dblGrTotalPrev + dblGrTotal) < 0
      
      If bNegativeP Then
         strNumberP = Left(strNumberP, Len(strNumberP) - 1)
      End If
      
      If bNegativeN Then
         strNumberN = Left(strNumberN, Len(strNumberN) - 1)
      End If
      
      If bNegativeT Then
         strNumberT = Left(strNumberT, Len(strNumberT) - 1)
      End If
      
'      strTemp = "<LINDENT=180>Grand Total<LINDENT=" & 5650 - .TextWidth(FormatNumber(dblGrTotalPrev, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblGrTotalPrev, 0, vbTrue, vbTrue, vbTrue) & "<LINDENT=" & 8150 - .TextWidth(FormatNumber(dblGrTotal, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblGrTotal, 0, vbTrue, vbTrue, vbTrue) & "<LINDENT=" & 10650 - .TextWidth(FormatNumber(dblGrTotalPrev + dblGrTotal, 0, vbTrue, vbTrue, vbTrue)) & ">" & FormatNumber(dblGrTotalPrev + dblGrTotal, 0, vbTrue, vbTrue, vbTrue)
      strTemp = "<LINDENT=180>Grand Total<LINDENT=" & 5650 - .TextWidth(strNumberP) & ">" & strNumberP & IIf(bNegativeP, ")", "") & "<LINDENT=" & 8150 - .TextWidth(strNumberN) & ">" & strNumberN & IIf(bNegativeN, ")", "") & "<LINDENT=" & 10650 - .TextWidth(strNumberT) & ">" & strNumberT & IIf(bNegativeT, ")", "")
      
      MyPrinter.AddText strTemp, "Arial", 9, True
   End With
   
   intPage = MyPrinter.Pages
   
   MyPrinter.AddLine 180, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip - 45, 10650, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip - 45, 1
   MyPrinter.AddLine 180, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 315, 10650, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 315, 1
   MyPrinter.AddLine 180, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 285, 10650, MyPrinter.TextItem(MyPrinter.TextItem.Count).PositionStartTwip + 285, 1, 3
   
   Unload frmPrintProgress
   
   MyPrinter.Preview Me, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
   
   Set MyPrinter = Nothing
   
   Me.Enabled = True
End Sub
