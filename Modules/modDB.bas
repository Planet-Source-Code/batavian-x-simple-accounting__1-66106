Attribute VB_Name = "modDB"
Option Explicit

Public ADORS As New ADODB.Recordset
Public ADOCnn As New ADODB.Connection

'***************************
'Database BAS Creator module generator
'Made by Miroslav Milak, mmilak@net4u.hr
'Module created: 13/05/2006 23:56:42
'Note:
'References to include in your product:
'  - Microsoft scripting runtime
'  - Microsoft ADO Extensions 2.x for DDL and Security Object Model
'  - Microsoft ActiveX Data Object Library 2.x
'***************************
'Modified by Batavian
'***************************

Private Cat As New ADOX.Catalog
Private Col As Column
Private Tbl As Table
Private Key As Key
Private Idx As Index

Public Sub CreateBlankDB()
   Dim strCoA() As String
   Dim strDes() As String
   Dim strRem() As String
   Dim strCat() As String
   Dim strBol() As String
   
   Dim intCount As Integer
   
   If Not CreateDatabase Then
      Exit Sub
   End If
        
   CreateTables
   CreateIndexes
   
   strCoA = Split(LoadResString(101), "|")
   strDes = Split(LoadResString(102), "|")
   strRem = Split(LoadResString(103), "|")
   strCat = Split(LoadResString(104), "|")
   strBol = Split(LoadResString(105), "|")
   
   CloseConn
   ADORS.Open "SELECT * FROM tblCoA;", ADOCnn, adOpenDynamic, adLockOptimistic
   
   For intCount = 0 To UBound(strCoA)
      ADORS.AddNew
      
      ADORS!CoA = strCoA(intCount)
      ADORS!Description = strDes(intCount)
      ADORS!Remark = strRem(intCount)
      ADORS!Category = strCat(intCount)
      ADORS!IsDebt = CBool(strBol(intCount))
   Next intCount
   
   ADORS.Update
   ADORS.UpdateBatch
   
   Set Cat = Nothing
End Sub

Private Function CreateDatabase() As Boolean
   On Error GoTo PROC_ERR

   Cat.Create "Provider=Microsoft.jet.oledb.4.0; Data Source =" & AppPath & "Acc.mdb;Jet OLEDB:Database Password='B@t@v!@n@J@k@rt@T!mu';"
   CreateDatabase = True
   
   Exit Function
   
PROC_ERR:
   If Err.Number = -2147217897 Then
      MsgBox "Database already exists."
      Exit Function
   Else
      MsgBox Err.Number & vbNewLine & Err.Description
   End If
End Function

Private Sub CreateTables()
   On Error GoTo PROC_ERR

   Set Tbl = New ADOX.Table
   
   With Tbl
      .Name = "tblCoA"
      
      Set .ParentCatalog = Cat
      
      With .Columns
         .Append "ID", adInteger, 0
         .Item("ID").Precision = "10"
         .Item("ID").Properties("Autoincrement").Value = True
         .Append "CoA", adVarWChar, 50
         .Append "Description", adVarWChar, 255
         .Append "Remark", adVarWChar, 255
         .Append "Category", adVarWChar, 50
         .Append "IsDebt", adBoolean, 2
      End With
   End With
   
   Cat.Tables.Append Tbl
   
   Set Tbl = Nothing
   Set Tbl = New ADOX.Table
   
   With Tbl
      .Name = "tblJournal"
      
      Set .ParentCatalog = Cat
      
      With .Columns
         .Append "IDJournal", adInteger, 0
         .Item("IDJournal").Precision = "10"
         .Item("IDJournal").Properties("Autoincrement").Value = True
         .Append "Voucher", adVarWChar, 50
         .Item("Voucher").Attributes = adColNullable
         .Append "Date", adDate, 0
         .Item("Date").Attributes = adColNullable
         .Append "Description", adVarWChar, 255
         .Item("Description").Attributes = adColNullable
         .Append "CreditCoA", adVarWChar, 255
         .Item("CreditCoA").Attributes = adColNullable
         .Append "Credit", adDouble, 0
         .Item("Credit").Precision = "15"
         .Item("Credit").Properties("Default").Value = "0"
         .Append "DebitCoA", adVarWChar, 255
         .Item("DebitCoA").Attributes = adColNullable
         .Append "Debit", adDouble, 0
         .Item("Debit").Precision = "15"
         .Item("Debit").Properties("Default").Value = "0"
         .Append "Remark", adVarWChar, 255
         .Item("Remark").Attributes = adColNullable
         .Append "Category", adVarWChar, 255
         .Item("Category").Attributes = adColNullable
         .Append "Customer", adVarWChar, 255
         .Item("Customer").Attributes = adColNullable
         .Append "Notes", adLongVarWChar, 0
         .Item("Notes").Attributes = adColNullable
         .Append "InternalID", adVarWChar, 50
         .Item("InternalID").Attributes = adColNullable
      End With
   End With
   
   Cat.Tables.Append Tbl
   
   Set Tbl = Nothing
   
   Exit Sub
   
PROC_ERR:
   MsgBox Err.Number & vbNewLine & Err.Description
End Sub

Private Sub CreateIndexes()
   On Error GoTo PROC_ERR

   Set Idx = New ADOX.Index
   
   Idx.Name = "PrimaryKey"
   Idx.Columns.Append "ID"
   Idx.PrimaryKey = True
   Idx.Unique = True
   Idx.Clustered = False
   Idx.IndexNulls = 1
   
   Cat.Tables("tblCoA").Indexes.Append Idx

   Set Idx = New ADOX.Index
   
   Idx.Name = "PrimaryKey"
   Idx.Columns.Append "IDJournal"
   Idx.PrimaryKey = True
   Idx.Unique = True
   Idx.Clustered = False
   Idx.IndexNulls = 1
   
   Cat.Tables("tblJournal").Indexes.Append Idx

   Set Idx = Nothing

   Exit Sub
   
PROC_ERR:
   MsgBox Err.Number & vbNewLine & Err.Description
End Sub

Public Sub CloseConn(Optional bOpenNew As Boolean = True, Optional lngFormHwnd As Long = -1)
   Dim sConn As String
   
   On Error GoTo ErrHandler
   
   If lngFormHwnd = -1 Then lngFormHwnd = frmMain.hWnd
   
   sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath & "Acc.mdb" & ";Jet OLEDB:Database Password='B@t@v!@n@J@k@rt@T!mu';"
   
   If ADOCnn.State <> adStateClosed Then ADOCnn.Close
   If ADORS.State <> adStateClosed Then ADORS.Close
   
   If bOpenNew Then
      If Dir(AppPath & "Acc.mdb", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
         xMsgBox lngFormHwnd, "Database file doesn't exist. ", vbCritical, "Fatal Error"
      Else
         If GetAttr(AppPath & "Acc.mdb") And vbReadOnly = vbReadOnly Or _
            GetAttr(AppPath & "Acc.mdb") And vbHidden = vbHidden Or _
            GetAttr(AppPath & "Acc.mdb") And vbSystem = vbSystem Then
            SetAttr AppPath & "Acc.mdb", vbArchive
         End If
         
         ADOCnn.CursorLocation = adUseClient
         ADOCnn.Open sConn
      End If
   End If
   
ErrHandler:
   If Err.Number = -2147467259 Then
      If xMsgBox(lngFormHwnd, "Database file corrupted, do you want to pick backup database?", vbQuestion Or vbYesNo, "Database Corrupted") = vbYes Then
         frmDBList.Show vbModal, frmMain
      Else
         Unload frmMain
      End If
   End If
End Sub

Public Function GenerateSlipNo(strYear As String, strMonth As String, strDay As String, Optional lngFormHwnd As Long = -1) As Integer
   Dim intTemp As Integer
   
   CloseConn , lngFormHwnd
   ADORS.Open "SELECT DISTINCT Int(Max(Right(Voucher,3))) AS intUBound FROM tblJournal WHERE Voucher LIKE '" & Right(strYear, 2) & "." & Format(strMonth, "00") & "." & Format(strDay, "00") & "%';", ADOCnn
      
   Do Until ADORS.EOF
      If intTemp < ADORS!intUBound Then
         intTemp = ADORS!intUBound
      End If
      
      ADORS.MoveNext
   Loop
   
   GenerateSlipNo = intTemp + 1
End Function

Public Function GenerateFullSlip(strYear As String, strMonth As String, strDate As String) As String
   GenerateFullSlip = Right(strYear, 2) & "." & Format(strMonth, "00") & "." & Format(strDate, "00") & "." & String(3 - Len(CStr(intCurrRecord)), "0") & CStr(intCurrRecord)
End Function
