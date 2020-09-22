VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Simple Accounting"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   2940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   2940
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   735
      Picture         =   "frmMain.frx":57E2
      ScaleHeight     =   810
      ScaleWidth      =   810
      TabIndex        =   1
      Top             =   585
      Visible         =   0   'False
      Width           =   810
   End
   Begin ComctlLib.ListView LV1 
      Height          =   1770
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   3122
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Slip No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code of Account"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Debit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Credit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Note"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList IL1 
      Left            =   2115
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8160
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":833A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8514
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8866
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":925C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":996E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A012
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A364
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A6B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":AA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":AF5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain1 
      Caption         =   "&Report"
      Begin VB.Menu mnuSub1 
         Caption         =   "Journal Report{IMG:I7}"
         Index           =   0
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Account Report{IMG:I7}"
         Index           =   1
         Begin VB.Menu mnuSub11 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "General Ledger{IMG:I5}"
         Index           =   3
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Trial Balance{IMG:I6}"
         Index           =   4
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Income Statement{IMG:I8}"
         Index           =   5
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Balance Sheet"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "Exit{IMG:I10}"
         Index           =   9
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuSub1 
         Caption         =   "About program...{DEFAULT}"
         Index           =   11
      End
   End
   Begin VB.Menu mnuMain2 
      Caption         =   "&Journal"
      Begin VB.Menu mnuOpenBal 
         Caption         =   "Opening balance{IMG:I5}"
      End
      Begin VB.Menu mnuSub2Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSub2 
         Caption         =   "Add new transaction{IMG:I11}"
         Index           =   0
      End
      Begin VB.Menu mnuSub2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSub2 
         Caption         =   "Open{DEFAULT|IMG:I13}"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuSub2 
         Caption         =   "Delete{IMG:I12}"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu mnuMain3 
      Caption         =   "CoA"
      Visible         =   0   'False
      Begin VB.Menu mnuSub3 
         Caption         =   "Add new{IMG:I11}"
         Index           =   0
      End
      Begin VB.Menu mnuSub3 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSub3 
         Caption         =   "Edit{DEFAULT|IMG:I13}"
         Index           =   2
      End
      Begin VB.Menu mnuSub3 
         Caption         =   "Delete{IMG:I12}"
         Index           =   3
      End
   End
   Begin VB.Menu mnuMain4 
      Caption         =   "&Misc"
      Begin VB.Menu mnuSub4 
         Caption         =   "Search phrase in the list{IMG:I15}"
         Index           =   0
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "View by date{IMG:I16}"
         Index           =   1
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "Print CoA list{IMG:I14}"
         Index           =   3
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "Show Calculator"
         Index           =   5
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSub4 
         Caption         =   "Database backup"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' Jakarta, May 2006
'==============================================================
' Project Title : Simple Accounting
' Author        : Batavian
' Email         : batavian_codes@yahoo.com
' Date          : May '06
'==============================================================
' Thanks to:
' ~~~~~~~~~~
' 1. LaVolpe
'    The greatest VB Author in the Left$("PSC", 1)
'    There will be no enough thanks for you.
' 2. edward moth and qbd software ltd
' 3. Randy Birch
' 4. Miroslav Milak
' 5. Everyone in PSC
'==============================================================
' Objects which used in this project:
' - LaVolpe's "modified formatted codes" Menu Classes
' - LaVolpe's La Volpe Buttons VH.1
' - edward moth's "modified formatted codes" Printer Classes
'==============================================================
' Warning:
' ~~~~~~~~
' This project contains messy -my part- codes caused by
' limited time, resources and a so called minimum payment. :(
' I had to finished it within 2 week with no accounting background,
' I wasted most of the time to learn accountancy,
' so give me a break please?!
' Next project (if some what happen :) will be better, that's 4 sure!
'==============================================================
' Request:
' ~~~~~~~~
' I've wrote a module that create a "SysListView32" object at
' runtime without any Windows Common Control OCX requirement,
' so it will be very usefull for a small program, which does't
' require any other file except the program file it-self.
'
' It will do:
' - Create Listviews at runtime
' - Create any amount of Headers/Columns and resize them
' - Set/Get string to any Listview's Item/Sub Item
' - [un]Check and/or [un]select Listview's Items
' - Much more...
'
' Problem:
' - When I click to one item then click it again, it'll surely
'   crash, it will even randomly crash the whole system.
' - My solution was to subclassed the Listview, yes there are
'   no crash happened, but yet it acted with a strange behaviour.
' - Occurred on a non check-box style Listview only.
'
' Request:
' - Correct solution for the problem
' - Classes from the module (would be easy)
' - Post it in PSC (no credit for me are needed)
'
' I'll post a project which using the module if there any
' request, or you can contact me by e-mail.
'
' Another request:
' - Job offers
'==============================================================

Option Explicit

Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean

Private Const FLASHW_CAPTION   As Long = &H1
Private Const FLASHW_TRAY      As Long = &H2
Private Const FLASHW_ALL       As Long = (FLASHW_CAPTION Or FLASHW_TRAY)
Private Const FLASHW_TIMER     As Long = &H4
Private Const FLASHW_TIMERNOFG As Long = &HC

Private Type FLASHWINFO
    cbSize    As Long
    hWnd      As Long
    dwFlags   As Long
    uCount    As Long
    dwTimeout As Long
End Type

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private bLVSelected As Boolean
Private bGenerating As Boolean

Public strDate As String

Private Sub Form_Load()
   Dim intCount As Integer
   Dim intLoop As Integer
   Dim bExist As Boolean
   Dim bCalc As Boolean
   Dim lngPrevIns As Long
   Dim FlashInfo As FLASHWINFO
   Dim strCaption As String
   Dim intOBYear As Integer
   Dim bBackup As Boolean

   Me.Width = GetSetting("Batavian's Accounting Program", "Setting", "Width", 640 * Screen.TwipsPerPixelX)
   Me.Height = GetSetting("Batavian's Accounting Program", "Setting", "Height", 480 * Screen.TwipsPerPixelY)
   Me.Left = GetSetting("Batavian's Accounting Program", "Setting", "X", (Screen.Width - Me.Width) / 2)
   Me.Top = GetSetting("Batavian's Accounting Program", "Setting", "Y", (Screen.Height - Me.Height) / 2)
   Me.WindowState = IIf(GetSetting("Batavian's Accounting Program", "Setting", "Maximized", False), vbMaximized, vbNormal)
   
   ' =============================
   ' This codes check the existence of manifest resource.
   ' If manifest resource exist, don't create manifest file.
   ' =============================
   Dim strRes As String
   Dim bManifestExist As Boolean
   
   On Error Resume Next
   
   strRes = LoadResData(1, 24)
   
   If Err.Number = 326 Then
      bManifestExist = False
   Else
      If strRes <> vbNullString Then
         bManifestExist = True
      End If
   End If
   ' =============================
   ' End of check the existence of manifest resource
   ' =============================
   
   On Error GoTo 0
   
   If Not bManifestExist Then
      If App.EXEName <> "prjAccounting" Then
         If Dir(AppPath & App.EXEName & ".exe.manifest", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
            Dim intFile As Integer
            
            intFile = FreeFile
            
            Open AppPath & App.EXEName & ".exe.manifest" For Output As #intFile
            
            Print #intFile, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
            Print #intFile, "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">"
            Print #intFile, "<dependency>"
            Print #intFile, "    <dependentAssembly>"
            Print #intFile, "        <assemblyIdentity"
            Print #intFile, "            type=""win32"""
            Print #intFile, "            name=""Microsoft.Windows.Common-Controls"""
            Print #intFile, "            version=""6.0.0.0"""
            Print #intFile, "            processorArchitecture=""X86"""
            Print #intFile, "            publicKeyToken=""6595b64144ccf1df"""
            Print #intFile, "            language=""*"""
            Print #intFile, "        />"
            Print #intFile, "    </dependentAssembly>"
            Print #intFile, "</dependency>"
            Print #intFile, "</assembly>"
            
            Close #intFile
            
            xMsgBox 0, "Manifest file created, please restart this application! "
            Unload Me
            Exit Sub
         End If
      End If
   End If

   If App.PrevInstance Then
      strCaption = Me.Caption
      Me.Caption = "Not me that you gonna flash!"
      lngPrevIns = FindWindow("ThunderRT6FormDC", strCaption)
      
      FlashInfo.cbSize = Len(FlashInfo)
      FlashInfo.dwFlags = FLASHW_ALL
      FlashInfo.dwTimeout = 0
      FlashInfo.hWnd = lngPrevIns
      FlashInfo.uCount = 5
      FlashWindowEx FlashInfo
      
      xMsgBox 0, "Another instance of this application is already running!", vbInformation, strCaption
      BringWindowToTop lngPrevIns
      
      Unload Me
      Exit Sub
   End If
   
   If Dir(AppPath & "Acc.mdb", vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) = vbNullString Then
      If xMsgBox(0, "Database file is missing, the program cannot continue until database created or restored! " & vbCrLf & vbCrLf & "Do you want to created a new blank database? ", vbCritical Or vbYesNo, "Fatal Error") = vbYes Then
         If xMsgBox(0, "Do you want to restore database instead of create a new one? " & vbCrLf & vbCrLf & "Note: If there aren't any existing backup files, you should choose no! ", vbQuestion Or vbYesNo, "Database Error") = vbYes Then
            frmDBList.Show vbModal, Me
         Else
            CreateBlankDB
         End If
      Else
         Unload Me
         Exit Sub
      End If
   End If
   
   If Dir(AppPath & "Acc.mdb", vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) = vbNullString Then
      CreateBlankDB
   End If
   
   If Dir(AppPath & "Acc.mdb", vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) = vbNullString Then
      xMsgBox 0, "No database file created, the program cannot continue! "
      Unload Me
      Exit Sub
   End If
   
   If Dir(AppPath & "DB_Backup", vbDirectory) = vbNullString Then
      MkDir AppPath & "DB_Backup"
   End If
   
   For intCount = 1 To 20
      If Dir(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbHidden Or vbArchive Or vbReadOnly Or vbSystem Or vbDirectory) <> vbNullString Then
         If DateValue(FileDateTime(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb")) = DateValue(Date) Then
            bBackup = True
         End If
      End If
   Next intCount
   
   On Error GoTo ErrDB
   
   If Not bBackup Then
      For intCount = 1 To 20
         If Dir(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbHidden Or vbArchive Or vbReadOnly Or vbSystem Or vbDirectory) = vbNullString Then
            frmPrintProgress.Show vbModeless
            frmPrintProgress.Label1 = "Database backup process, please wait..."
            DoEvents
            
            CloseConn False, Me.hWnd
            FileCopy AppPath & "Acc.mdb", AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
            ChangeFileTime AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
            bBackup = True
            
            Unload frmPrintProgress
            Exit For
         End If
      Next intCount
   End If
   
   If Not bBackup Then
      frmPrintProgress.Show vbModeless
      frmPrintProgress.Label1 = "Database backup process, please wait..."
      DoEvents
      
      For intCount = 1 To 19
         SetAttr AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbNormal
         Kill AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
         FileCopy AppPath & "DB_Backup\DB_Backup_" & intCount + 1 & ".mdb", AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
      Next intCount
      
      CloseConn False, Me.hWnd
      FileCopy AppPath & "Acc.mdb", AppPath & "DB_Backup\DB_Backup_" & 20 & ".mdb"
      ChangeFileTime AppPath & "DB_Backup\DB_Backup_" & 20 & ".mdb"
            
      Unload frmPrintProgress
   End If
   
   On Error GoTo 0
   
ErrPass:
   CloseConn , Me.hWnd
   ADORS.Open "SELECT Date, InternalID FROM tblJournal WHERE InternalID= 'OB';", ADOCnn
   
   If ADORS.RecordCount < 1 Then
      xMsgBox 0, "Please add opening balance entries! ", vbInformation, "Opening Balance"
      frmOpeningBalance.Show vbModeless, Me
   Else
      intOBYear = Year(ADORS!Date)
      
      If intOBYear < Year(Date) Then
         CloseConn , Me.hWnd
         ADORS.Open "SELECT InternalID FROM tblJournal WHERE InternalID= 'OBY' AND Date BETWEEN #1/1/" & Year(Date) & "# AND #31/12/" & Year(Date) & "#;", ADOCnn
         
         If ADORS.RecordCount < 1 Then
            If xMsgBox(0, "The program cannot continue until opening balance for this year is generated! " & vbCrLf & "Do you want the program to generate opening balance now? ", vbQuestion Or vbYesNo, "Opening Balance of Year") = vbYes Then
               addOPoYear
            Else
               Unload Me
               Exit Sub
            End If
         End If
      End If
   End If
   
   For intCount = 0 To LV1.ColumnHeaders.Count - 1
      If GetSetting("Batavian's Accounting Program", "Setting", "MainLV" & intCount) <> vbNullString Then
         SendMessage LV1.hWnd, LVM_SETCOLUMNWIDTH, intCount, CLng(GetSetting("Batavian's Accounting Program", "Setting", "MainLV" & intCount))
      End If
   Next intCount

   SetLVExtendedStyle LV1, True
   SetLVColumnOrder LV1, GetSetting("Batavian's Accounting Program", "Setting", "ColumnOrder", "0;1;2;3;4;5;6;7;8")
   
   bCalc = GetSetting("Batavian's Accounting Program", "Setting", "ShowCalc", "0") = "1"

   If bCalc Then
      mnuSub4(5).Checked = True
      frmCalc.Show vbModeless, Me
   End If

   MyCoolTip.BackColor = RGB(31, 31, 31)
   MyCoolTip.ForeColor = RGB(223, 223, 223)

   CloseConn , Me.hWnd
   ADORS.Open "SELECT DISTINCT Remark FROM tblCoA ORDER BY Remark;", ADOCnn

   Do Until ADORS.EOF
      mnuSub11(mnuSub11.UBound).Visible = True
      mnuSub11(mnuSub11.UBound).Enabled = True
      mnuSub11(mnuSub11.UBound).Caption = CreateMenuCaption(ADORS!Remark, lv_ImgListIndex, "7")

      Load mnuSub11(mnuSub11.UBound + 1)

      ADORS.MoveNext
   Loop

   Unload mnuSub11(mnuSub11.UBound)
   
   strDate = "BETWEEN #" & Month(Date) & "/1/" & Year(Date) & "# AND #" & Month(Date) & "/" & MaxDayOfMonth(Year(Date), Month(Date)) & "/" & Year(Date) & "#"
   
   FillMainLV
   SetMenuForm Me.hWnd, IL1
   SetMinMaxInfo Me.hWnd, -1, -1, 0, 0, -1, -1, 640, 480
   
   Exit Sub
' === Remind me to comment below lines
'   addOPoYear
ErrDB:
   xMsgBox Me.hWnd, "Database backup has failed, check if any other program using this file! ", vbCritical, "Backup Failed"
   GoTo ErrPass
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim MyAllForm As Form
   Dim intCount As Integer
   
   CloseConn False

   Set ADORS = Nothing
   Set ADOCnn = Nothing

   Set MyTBDesc = Nothing
   Set MyCoolTip = Nothing
   
   Set picLogo.Picture = LoadPicture("")
   
   SaveSetting "Batavian's Accounting Program", "Setting", "Maximized", IIf(Me.WindowState = vbMaximized, True, False)

   If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
   
   For intCount = 1 To mnuSub11.UBound
      Unload mnuSub11(intCount)
   Next intCount

   SaveSetting "Batavian's Accounting Program", "Setting", "Width", Me.Width
   SaveSetting "Batavian's Accounting Program", "Setting", "Height", Me.Height
   SaveSetting "Batavian's Accounting Program", "Setting", "X", Me.Left
   SaveSetting "Batavian's Accounting Program", "Setting", "Y", Me.Top
   SaveSetting "Batavian's Accounting Program", "Setting", "ShowCalc", IIf(mnuSub4(5).Checked, "1", "0")
   SaveSetting "Batavian's Accounting Program", "Setting", "ColumnOrder", GetLVColumnOrder(LV1)

   For intCount = 0 To LV1.ColumnHeaders.Count - 1
      SaveSetting "Batavian's Accounting Program", "Setting", "MainLV" & intCount, SendMessage(LV1.hWnd, LVM_GETCOLUMNWIDTH, intCount, 0&)
   Next intCount

   For Each MyAllForm In Forms
      If MyAllForm.Name <> Me.Name Then Unload MyAllForm
   Next MyAllForm
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub

   Me.LV1.Left = 30
   Me.LV1.Top = 30

   If Me.ScaleWidth >= 300 Then
      Me.LV1.Width = Me.ScaleWidth - 60
   End If

   If Me.ScaleHeight >= 300 Then
      Me.LV1.Height = Me.ScaleHeight - 60
   End If
End Sub

Private Sub LV1_DblClick()
   If bLVSelected Then
      Call mnuSub2_Click(2)
   End If
End Sub

Private Sub LV1_ItemClick(ByVal Item As ComctlLib.ListItem)
   bLVSelected = True
   mnuSub2(2).Enabled = True
   mnuSub2(3).Enabled = True
End Sub

Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      mnuOpenBal.Visible = False
      mnuSub2Sep.Visible = False

      If bLVSelected Then
         mnuSub2(2).Enabled = True
         mnuSub2(3).Enabled = True
      Else
         mnuSub2(2).Enabled = False
         mnuSub2(3).Enabled = False
      End If

      PopupMenu mnuMain2

      mnuOpenBal.Visible = True
      mnuSub2Sep.Visible = True
   End If
End Sub

Private Sub mnuMain2_Click()
   bLVSelected = False
End Sub

Private Sub mnuOpenBal_Click()
   CloseConn , Me.hWnd
   ADORS.Open "SELECT InternalID FROM tblJournal WHERE InternalID = 'OB';", ADOCnn ' WHERE Date >= #1/1/" & year(date) & "# AND Date <= #31/12/" & year(date) & ";", ADOCnn

   If ADORS.RecordCount > 0 Then
      If xMsgBox(Me.hWnd, "Opening balance entries is exist in the database, " & vbCrLf & _
         "if you decided to enter new opening balance entries, " & vbCrLf & _
         "all previous Opening balance entries will be deleted. " & vbCrLf & vbCrLf & _
         "Are you sure to do this?", vbExclamation Or vbYesNo, "Opening Balance Entry") <> vbYes Then
         Exit Sub
      End If
   End If

   frmOpeningBalance.Show vbModeless, Me
End Sub

Private Sub mnuSub1_Click(Index As Integer)
   Select Case Index
      Case 0, 3, 4, 5
         PrintReport Index
      Case 9
         Unload Me
      Case 11
         xMsgBox Me.hWnd, "Simple Accounting Program  " & vbCrLf & "Version " & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "Licensed to:" & vbCrLf & "PT Majuko Utama Indonesia  ", vbInformation, "About Program"
   End Select

   bLVSelected = False
End Sub

Private Sub mnuSub11_Click(Index As Integer)
   Dim intCount As Integer
   
   frmRemPrint.strRemark = Left(mnuSub11(Index).Caption, InStr(mnuSub11(Index).Caption, "{") - 1)
   
   For intCount = 0 To mnuSub11.UBound
      frmRemPrint.cboRemark.AddItem Left(mnuSub11(intCount).Caption, InStr(mnuSub11(intCount).Caption, "{") - 1)
   Next intCount
   
   frmRemPrint.cboRemark = Left(mnuSub11(Index).Caption, InStr(mnuSub11(Index).Caption, "{") - 1)
   frmRemPrint.Show vbModeless, Me
End Sub

Private Sub mnuSub2_Click(Index As Integer)
   Dim lngIndex As Long
'   Dim strJournalDate As String
   Dim strSlip As String

   With frmJournal
      Select Case Index
         Case 0
            .bNewTransaction = True
            .Top = Me.Top + 300
            .Left = Me.Left + 300
            .Show vbModeless, Me
         Case 2
            .bNewTransaction = False
            .strSlipNo = LV1.SelectedItem.Tag
            .Top = Me.Top + 300
            .Left = Me.Left + 300
            .Show vbModeless, Me
         Case 3
            strSlip = LV1.SelectedItem.Tag
'            strJournalDate = LV1.SelectedItem.SubItems(1)

            If xMsgBox(Me.hWnd, "Are You sure to delete this record?" & vbCrLf & vbCrLf & _
               "Slip No: " & strSlip & vbCrLf, _
               vbQuestion Or vbYesNo, "Delete Record") = vbYes Then

               CloseConn , Me.hWnd
               ADORS.Open "SELECT * FROM tblJournal WHERE Voucher = '" & strSlip & "';", ADOCnn, adOpenDynamic, adLockOptimistic

               Do Until ADORS.EOF
                  ADORS.Delete
                  ADORS.MoveNext
               Loop

               For lngIndex = LV1.ListItems.Count To 1 Step -1
                  If LV1.ListItems.Item(lngIndex).Tag = strSlip Then
                     LV1.ListItems.Remove lngIndex
                  End If
               Next lngIndex
            End If
      End Select
   End With
End Sub

Private Sub mnuSub3_Click(Index As Integer)
   Select Case Index
      Case 0: Call frmCoA.lvButtons_H3_Click
      Case 2: Call frmCoA.lvButtons_H4_Click
      Case 3: Call frmCoA.lvButtons_H5_Click
   End Select
End Sub

Public Sub FillMainLV()
   Dim LV As ListItem
   Dim strTemp As String

   LV1.ListItems.Clear

   CloseConn , Me.hWnd
   ADORS.Open "SELECT * FROM tblJournal WHERE Date " & strDate & " ORDER BY Date, Voucher, IDJournal;", ADOCnn

   Do Until ADORS.EOF
      Set LV = LV1.ListItems.Add(, , "")

      If LV1.ListItems.Count < 1 Or Not (strTemp = ADORS!Voucher) Then
         strTemp = ADORS!Voucher
         LV.Text = ADORS!Voucher
         LV.SubItems(1) = ADORS!Date
         LV.SubItems(3) = ADORS!Description
         LV.SubItems(9) = IIf(IsNull(ADORS!Notes), vbNullString, ADORS!Notes)
      End If

      LV.Tag = ADORS!Voucher
      LV.SubItems(2) = IIf(IsNull(ADORS!DebitCoA), "    " & ADORS!CreditCoa, ADORS!DebitCoA)
      LV.SubItems(4) = ADORS!Remark
      LV.SubItems(5) = ADORS!Category
      LV.SubItems(6) = IIf(ADORS!Debit = 0, "-", FormatNumber(ADORS!Debit, 0, vbTrue, vbTrue, vbTrue)) ' "#,###"))
      LV.SubItems(7) = IIf(ADORS!Credit = 0, "-", FormatNumber(ADORS!Credit, 0, vbTrue, vbTrue, vbTrue)) ' "#,###"))
      LV.SubItems(8) = IIf(IsNull(ADORS!Customer), vbNullString, ADORS!Customer)

      ADORS.MoveNext
   Loop

   bLVSelected = False
End Sub

Private Sub PrintReport(Index As Integer, Optional intAditionalIndex As Integer)
      Select Case Index
         Case 0
            frmJournalPrint.Show vbModeless, Me
         Case 3
            frmGL.Show vbModeless, Me
         Case 4
            frmTrialBalance.Show vbModeless, Me
         Case 5
            frmIncome.Show vbModeless, Me
      End Select
End Sub

Private Sub mnuSub4_Click(Index As Integer)
   Select Case Index
      Case 0
         frmSearch.Show vbModeless, Me
      Case 1
         frmByDate.Show vbModeless, Me
      Case 3
         Const strProgress As String = "Printing, please wait ...."
   
         Dim MyPrinter As New clsPrinter
         Dim intCount As Integer
         
         frmPrintProgress.Show vbModeless, Me
         frmPrintProgress.Label1 = strProgress & vbCrLf & "0% complete."
         DoEvents

         AddReportLogo picLogo, "Chart of Accounts", MyPrinter, qPortrait
         
         MyPrinter.MarginBottom = 600
         MyPrinter.AddText " ", "Arial", 24
         MyPrinter.AddText "<LINDENT=210>Category<LINDENT=1450>Code<LINDENT=2400>Description<LINDENT=7500>Remark", "Arial", 12, True
         MyPrinter.AddText " ", "Arial", 8
         MyPrinter.AddLine 180, 1150, 10650, 1150, 1
         MyPrinter.AddLine 180, 1105, 10650, 1105, 1, 2
         MyPrinter.Footer(EvenPage_hf).Text = "Page #pagenumber# of #pagetotal#"
         MyPrinter.Footer(OddPage_hf).Text = "Page #pagenumber# of #pagetotal#"
         MyPrinter.Footer(OddPage_hf).Alignment = eCentre
         MyPrinter.Footer(EvenPage_hf).Alignment = eCentre
         MyPrinter.Footer(OddPage_hf).FontSize = 7
         MyPrinter.Footer(EvenPage_hf).FontSize = 7
         MyPrinter.SetFooter(OddPage_hf) = True
         MyPrinter.SetFooter(EvenPage_hf) = True

         CloseConn , Me.hWnd
         ADORS.Open "SELECT Category, CoA, Description, Remark From tblCoA ORDER BY CoA;", ADOCnn

         Do Until ADORS.EOF
            frmPrintProgress.Label1 = strProgress & vbCrLf & Round((intCount / ADORS.RecordCount) * 100) & "% complete."
            MyPrinter.AddText "<LINDENT=210>" & ADORS!Category & "<LINDENT=1450>" & ADORS!CoA & "<LINDENT=2400>" & ADORS!Description & "<LINDENT=7500>" & ADORS!Remark, "Arial", 10
            ADORS.MoveNext
            intCount = intCount + 1
            DoEvents
         Loop

         MyPrinter.Preview Me, 99 * Screen.TwipsPerPixelX, 20 * Screen.TwipsPerPixelY
         
         Unload frmPrintProgress
         
         Set MyPrinter = Nothing
      Case 5
         mnuSub4(5).Checked = Not mnuSub4(5).Checked

         If mnuSub4(5).Checked Then
            frmCalc.Show vbModeless, Me
         Else
            Unload frmCalc
         End If
      Case 7
         frmDBList.Show vbModeless, Me
   End Select
End Sub

Private Sub addOPoYear()
   Const strProgress As String = "Generating opening balance, please wait ...."
   
   Dim dblRetainedGainLoss As Double
   Dim strAcc() As String
   Dim strRem() As String
   Dim strCat() As String
   Dim bDC() As Boolean
   Dim strAccX() As String
   Dim strRemX() As String
   Dim dblValue() As Double
   Dim strTemp As String
   Dim strTempX As String
   Dim strVoucher As String
   Dim strPreDate As String
   Dim strSQL As String
   Dim intCount As Integer
   Dim intCountX As Integer
   Dim bRetained As Boolean
   Dim intRetained As Integer
   Dim intProgress As Integer
   
   Me.Enabled = False
   Screen.MousePointer = vbArrowHourglass
   
   frmPrintProgress.Show vbModeless, Me
   frmPrintProgress.Label1 = strProgress & vbCrLf & "Calculate recent gain/loss..."
   DoEvents
   
   strVoucher = Right(Year(Date), 2) & ".01.01." & String(3 - Len(CStr(GenerateSlipNo(Year(Date), "1", "1", Me.hWnd))), "0") & CStr(GenerateSlipNo(Year(Date), "1", "1", Me.hWnd))
   strPreDate = "Date BETWEEN #1/1/" & Year(Date) - 1 & "# AND #12/31/" & Year(Date) - 1 & "#"
   
   strSQL = "SELECT SUM(AccIncome) - SUM(AccExpense) AS AccValue FROM "
   strSQL = strSQL & "(SELECT SUM(Credit) AS AccIncome, SUM(Debit) AS AccExpense FROM "
   strSQL = strSQL & "tblJournal WHERE " & strPreDate & " AND Category = 'Incomes' "
   strSQL = strSQL & "UNION ALL "
   strSQL = strSQL & "SELECT SUM(Credit) AS AccIncome, SUM(Debit) AS AccExpense FROM "
   strSQL = strSQL & "tblJournal WHERE " & strPreDate & " AND Category = 'Expenses');"
   
   CloseConn
   ADORS.Open strSQL, ADOCnn
   
   If ADORS.RecordCount > 0 And Not IsNull(ADORS!accvalue) Then
      dblRetainedGainLoss = ADORS!accvalue
   End If
   
   frmPrintProgress.Label1 = strProgress & vbCrLf & "Retrieve accounts."
   DoEvents
   
   strSQL = "SELECT DISTINCT AccountName, Remark "
   strSQL = strSQL & "FROM (SELECT DebitCoA AS AccountName, Remark FROM tblJournal WHERE " & strPreDate & " AND ISNULL(CreditCoA) AND Category <> 'Expenses' AND Category <> 'Incomes'"
   strSQL = strSQL & "UNION ALL "
   strSQL = strSQL & "SELECT CreditCoA AS AccountName, Remark FROM tblJournal WHERE " & strPreDate & " AND ISNULL(DebitCoA) AND Category <> 'Expenses' AND Category <> 'Incomes')"
   
'   Debug.Print strSQL
   
   CloseConn
   ADORS.Open strSQL, ADOCnn
   
   intProgress = 1
   
   Do Until ADORS.EOF
      frmPrintProgress.Label1 = strProgress & vbCrLf & "Retrieving accounts..." & Round(intProgress / ADORS.RecordCount * 100) & "% complete."
      DoEvents
      
      strTemp = strTemp & "|" & ADORS!AccountName
      strTempX = strTempX & "|" & ADORS!Remark
      
      intProgress = intProgress + 1
      ADORS.MoveNext
   Loop
   
   strTemp = Right(strTemp, Len(strTemp) - 1)
   strTempX = Right(strTempX, Len(strTempX) - 1)
   
   strAccX() = Split(strTemp, "|")
   strRemX() = Split(strTempX, "|")
   
   ReDim strAcc(UBound(strAccX))
   ReDim strRem(UBound(strAccX))
   ReDim bDC(UBound(strAccX))
   ReDim dblValue(UBound(strAccX))
   ReDim strCat(UBound(strAccX))
   
   frmPrintProgress.Label1 = strProgress & vbCrLf & "Retrieve value of accounts."
   DoEvents
   
   CloseConn
   ADORS.Open "SELECT * FROM tblCoA ORDER BY CoA;", ADOCnn
   
   intProgress = 1
   
   Do Until ADORS.EOF
      frmPrintProgress.Label1 = strProgress & vbCrLf & "Retrieve value of accounts..." & Round(intProgress / ADORS.RecordCount * 100) & "% complete."
      DoEvents
      
      For intCountX = 0 To UBound(strAccX)
         If ADORS!Description = strAccX(intCountX) And ADORS!Remark = strRemX(intCountX) Then
            Dim strClp As String
            
            strAcc(intCount) = ADORS!Description
            strRem(intCount) = ADORS!Remark
            bDC(intCount) = ADORS!IsDebt
            strCat(intCount) = ADORS!Category
            
            intCount = intCount + 1
            Exit For
         End If
      Next intCountX
      
      intProgress = intProgress + 1
      ADORS.MoveNext
   Loop
   
   For intCount = 0 To UBound(strAcc)
      CloseConn
      
      If bDC(intCount) Then
         ADORS.Open "SELECT SUM(Debit)-SUM(Credit) AS AccValue FROM tblJournal WHERE " & strPreDate & " AND (DebitCoA = '" & strAcc(intCount) & "' OR CreditCoA = '" & strAcc(intCount) & "');", ADOCnn
      Else
         ADORS.Open "SELECT SUM(Credit)-SUM(Debit) AS AccValue FROM tblJournal WHERE " & strPreDate & " AND (DebitCoA = '" & strAcc(intCount) & "' OR CreditCoA = '" & strAcc(intCount) & "');", ADOCnn
      End If
      
'      Debug.Print "SELECT SUM(Debit)-SUM(Credit) AS AccValue FROM tblJournal WHERE " & strPreDate & " AND (DebitCoA = '" & strAcc(intCount) & "' OR CreditCoA = '" & strAcc(intCount) & "');"
   
      If strAcc(intCount) = "Retained Earning" Or strAcc(intCount) = "Retained Loss" Then
         dblRetainedGainLoss = dblRetainedGainLoss + ADORS!accvalue
      End If
      
      If strAcc(intCount) <> "Retained Earning" And strAcc(intCount) <> "Retained Loss" Then
         dblValue(intCount) = ADORS!accvalue
      End If
   Next intCount
   
   frmPrintProgress.Label1 = strProgress & vbCrLf & "Writing opening balance entries."
   DoEvents
   
   CloseConn
   ADORS.Open "SELECT * FROM tblJournal;", ADOCnn, adOpenDynamic, adLockOptimistic
      
   For intCount = 0 To UBound(strAcc)
      If strAcc(intCount) <> "Retained Earning" And strAcc(intCount) <> "Retained Loss" _
         And dblValue(intCount) <> 0 Then
         ADORS.AddNew
         ADORS!Voucher = strVoucher
         ADORS!Date = DateSerial(Year(Date), 1, 1)
         ADORS!Description = "Opening Balance of Year"
         
         If bDC(intCount) Then
            ADORS!DebitCoA = strAcc(intCount)
            ADORS!Debit = dblValue(intCount)
         Else
            ADORS!CreditCoa = strAcc(intCount)
            ADORS!Credit = dblValue(intCount)
         End If
         
         ADORS!Remark = strRem(intCount)
         ADORS!Category = strCat(intCount)
         ADORS!InternalID = "OBY"
      End If
   Next intCount
   
   ADORS.AddNew
   ADORS!Voucher = strVoucher
   ADORS!Date = DateSerial(Year(Date), 1, 1)
   ADORS!Description = "Opening Balance of Year"
   ADORS!Credit = dblRetainedGainLoss
   ADORS!Remark = "Retained Earnings & Losses"
   ADORS!Category = "Equities"
   ADORS!InternalID = "OBY"
         
   If dblRetainedGainLoss < 0 Then
      ADORS!CreditCoa = "Retained Loss"
   Else
      ADORS!CreditCoa = "Retained Earning"
   End If
   
   ADORS.Update
   ADORS.UpdateBatch
   
   Unload frmPrintProgress
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

