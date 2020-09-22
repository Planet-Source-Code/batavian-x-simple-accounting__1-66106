VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDBList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database BackUp List"
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView LV1 
      Height          =   5640
      Left            =   52
      TabIndex        =   0
      Top             =   52
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   9948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "IL_Large"
      SmallIcons      =   "IL_Small"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Backup Date - Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList IL_Small 
      Left            =   2880
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDBList.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList IL_Large 
      Left            =   2880
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDBList.frx":035E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBU 
         Caption         =   "Backup database"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuIcons 
         Caption         =   "&Icons"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Details"
      End
   End
   Begin VB.Menu mnuX 
      Caption         =   "x"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
   End
End
Attribute VB_Name = "frmDBList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim intCount As Integer
   
   For intCount = 1 To 20
      If Dir(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) <> vbNullString Then
         LV1.ListItems.Add , , "DB_Backup_" & intCount
         LV1.ListItems(LV1.ListItems.Count).SubItems(1) = FileDateTime(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb")
         LV1.ListItems(LV1.ListItems.Count).Icon = 1
         LV1.ListItems(LV1.ListItems.Count).SmallIcon = 1
      End If
   Next
   
   SetLVExtendedStyle LV1
   SetMenuForm Me.hWnd
End Sub

Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim bSelected As Boolean
   Dim intCount As Integer
   
   For intCount = 1 To LV1.ListItems.Count
      If LV1.ListItems(intCount).Selected Then
         bSelected = True
         Exit For
      End If
   Next intCount
      
   mnuRestore.Enabled = bSelected
   Me.PopupMenu mnuX
End Sub

Private Sub mnuBU_Click()
   Dim bBackup As Boolean
   Dim intCount As Integer
   
   On Error GoTo ErrDB
   
   If xMsgBox(Me.hWnd, "Are you sure to backup database now? ", vbQuestion Or vbYesNo, "Database Backup") = vbYes Then
      For intCount = 1 To 20
         If Dir(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbHidden Or vbArchive Or vbReadOnly Or vbSystem Or vbDirectory) = vbNullString Then
            CloseConn False, Me.hWnd
            FileCopy AppPath & "Acc.mdb", AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
            ChangeFileTime AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
            bBackup = True
            Exit For
         End If
      Next intCount
   
      If Not bBackup Then
         For intCount = 1 To 19
            SetAttr AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbNormal
            Kill AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
            FileCopy AppPath & "DB_Backup\DB_Backup_" & intCount + 1 & ".mdb", AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb"
         Next intCount
         
         CloseConn False, Me.hWnd
         FileCopy AppPath & "Acc.mdb", AppPath & "DB_Backup\DB_Backup_" & 20 & ".mdb"
         ChangeFileTime AppPath & "DB_Backup\DB_Backup_" & 20 & ".mdb"
      End If
   
      LV1.ListItems.Clear
   
      For intCount = 1 To 20
         If Dir(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb", vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) <> vbNullString Then
            LV1.ListItems.Add , , "DB_Backup_" & intCount
            LV1.ListItems(LV1.ListItems.Count).SubItems(1) = FileDateTime(AppPath & "DB_Backup\DB_Backup_" & intCount & ".mdb")
            LV1.ListItems(LV1.ListItems.Count).Icon = 1
            LV1.ListItems(LV1.ListItems.Count).SmallIcon = 1
         End If
      Next
   End If
   
   Exit Sub
   
ErrDB:
   xMsgBox Me.hWnd, "Database backup has failed, check if any other program using this file! ", vbCritical, "Backup Failed"
End Sub

Private Sub mnuClose_Click()
   Unload Me
End Sub

Private Sub mnuDetails_Click()
   LV1.View = lvwReport
End Sub

Private Sub mnuIcons_Click()
   LV1.View = lvwIcon
End Sub

Private Sub mnuRestore_Click()
   On Error GoTo ErrHandler
   
   If xMsgBox(Me.hWnd, "Are you sure to restore this file, " & vbCrLf & vbCrLf & _
      "File name: " & LV1.SelectedItem.Text & " " & vbCrLf & vbCrLf & _
      "Created: " & LV1.SelectedItem.SubItems(1) & " " & vbCrLf & vbCrLf & _
      "Caution: Original database will be replaced, and cannot be undone! ", vbQuestion Or vbYesNo, "Database Restore") = vbYes Then
      
      CloseConn False, Me.hWnd
      
      If Dir(AppPath & "Acc.mdb", vbHidden Or vbArchive Or vbReadOnly Or vbSystem Or vbDirectory) <> vbNullString Then
         SetAttr AppPath & "Acc.mdb", vbNormal
         Kill AppPath & "Acc.mdb"
      End If
      
      FileCopy AppPath & "DB_Backup\" & LV1.SelectedItem.Text & ".mdb", AppPath & "Acc.mdb"
      xMsgBox Me.hWnd, "Database has been restored! ", vbInformation, "Restore Success"
   End If
   
   Call frmMain.FillMainLV
   Exit Sub
   
ErrHandler:
   xMsgBox Me.hWnd, "Database restoration has failed, check if any other program using this file! ", vbCritical, "Restore Failed"
End Sub
