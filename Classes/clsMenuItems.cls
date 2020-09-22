VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private imgList As Control
Private PrevProc As Long
Private TipCallBackProc As Long
Private PrevTipCallBack As Long
Private hSysMenu As Long
Private hMDIclient As Long
Private bMDI As Boolean
Private bNoChildMenus As Boolean
Private uMinMax As MINMAXINFO
Private vItems() As MenuComponentData
Private vPanels() As PanelData
Private cItems As Collection
Private cPanels As Collection
Private gMenus As Collection
Private Const LB_GETTEXT As Long = &H189
Private Const LB_GETTEXTLEN As Long = &H18A

Private Type CHOOSECOLORSTRUCT
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As Long
   Flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Private Type MENUITEMINFO_string
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As CHOOSECOLORSTRUCT) As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLCID As Long, ByVal lLCTYPE As Long, ByVal strLCData As String, ByVal lDataLen As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Sub lstrcpyn Lib "kernel32" (ByVal strDest As String, ByVal strSrc As Any, ByVal lBytes As Long)

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal sFile As String, ByVal lpIgnored As Long, ByVal lpSize As Long, ByVal lpBuf As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByVal lpBuf As Long, ByVal szReceive As String, lpBufPtr As Long, lLen As Long) As Long

Private FormHwnd As Long
Private tipDefault As String

Public Sub InitializeSubMenu(hForm As Long, Optional oImgList As Control, Optional CallBacks As Long, _
   Optional bNoSysMenu As Boolean, Optional TimeConsumingMenu As Long = 0, Optional DefaultTip As String, Optional Flags As Long = 0)
   Dim sysMenu As Long
   
   FormHwnd = hForm
   
   If CallBacks Then
      If CallBacks <> PrevTipCallBack Then
         PrevTipCallBack = TipCallBackProc
         TipCallBackProc = CallBacks
      End If
   Else
      TipCallBackProc = PrevTipCallBack
   End If
   
   If Not oImgList Is Nothing Then Set imgList = oImgList
   
   If hSysMenu = 0 And bNoSysMenu = False Then
      hSysMenu = GetSystemMenu(hForm, 0)
      
      If hSysMenu Then IsWindowList hSysMenu, True
   End If
   
   If TimeConsumingMenu Then
      Dim mPanel As PanelData
      
      mPanel.Hourglass = True
      
      ReDim vPanels(0 To 1)
      
      vPanels(1) = mPanel
      cPanels.Add 1, "p" & TimeConsumingMenu
   End If

   tipDefault = DefaultTip
End Sub

Public Property Get MDIClient() As Long
   MDIClient = hMDIclient
End Property

Public Property Let MDIClient(hwnd As Long)
   hMDIclient = hwnd
End Property

Public Property Get IsMDIclient() As Boolean
   IsMDIclient = bMDI
End Property

Public Property Let IsMDIclient(bYesNo As Boolean)
   bMDI = bYesNo
End Property

Public Property Get IsMenuLess() As Boolean
   IsMenuLess = bNoChildMenus
End Property

Public Property Let IsMenuLess(bYesNo As Boolean)
   bNoChildMenus = bYesNo
End Property

Public Property Get SystemMenu() As Long
   SystemMenu = hSysMenu
End Property

Public Property Get ShowTips() As Long
   ShowTips = TipCallBackProc
End Property

Public Function Tips(MenuID As Long, hMenu As Long) As String
   Dim mIndex As Integer
   
   On Error Resume Next
   
   mIndex = cItems(MenuID & "." & hMenu)
   
   If Err = 0 Then
      Tips = vItems(mIndex).Tip
      Exit Function
   End If
   
   Tips = gMenus("g" & hMenu).Tips(MenuID, hMenu)
   Err.Clear
End Function

Public Property Get ImageListObject() As Control
   Set ImageListObject = imgList
End Property

Public Property Let hPrevProc(hProc As Long)
   PrevProc = hProc
End Property

Public Property Get hPrevProc() As Long
   hPrevProc = PrevProc
End Property

Public Function RestrictSize(lParam As Long, bSet As Boolean) As Boolean
   If bSet Then
      CopyMemory uMinMax, ByVal lParam, Len(uMinMax)
   Else
      If uMinMax.ptMaxSize.x < 0 Then Exit Function
      
      CopyMemory ByVal lParam, uMinMax, Len(uMinMax)
      RestrictSize = True
   End If
End Function

Public Function MenuSelected(Index As Long, hMenu As Long, lParam As Long) As Boolean
   Dim mIndex As Long
   Dim pParam1 As Long
   Dim pParam2 As Long
   Dim msgParam As String
   Dim sUserID As String
   Dim sDefValue As String
   Dim i As Integer
   Dim sTarget As String
   Dim sValue As String
   Dim wCaption As String

   Const CB_SETCURSEL As Long = &H14E
   Const LB_SETCURSEL As Long = &H186
   Const LB_SETSEL As Long = &H185
   Const CC_RGBINIT As Long = &H1
   Const CC_FULLOPEN As Long = &H2
   Const CC_ANYCOLOR As Long = &H100

   On Error Resume Next
   
   If lParam Then
      mIndex = cItems(Index & "." & hMenu)
      
      If Err = 0 Then
         MenuSelected = True
         SeparateCaption vItems(mIndex).Caption, "", wCaption, ""
         
         For i = 1 To 7
            sTarget = Choose(i, "LColor:", "LMonth:", "LState:", "LDay:", "LDate:", "LFont:", "LDrv:")
            ReturnComponentValue wCaption, sTarget, sValue
            
            If Len(sValue) Then Exit For
         Next
         
         If Len(sValue) Then
            ReturnComponentValue wCaption, "lvID:", sUserID
            
            Dim oTipClass As clsTips
            
            CopyMemory oTipClass, TipCallBackProc, 4&
            
            Select Case sTarget
               Case "LColor:"
                  If Val(sValue) = -1 Then
                     Dim CC As CHOOSECOLORSTRUCT
                     Dim cCusGray As Long
                     Dim lGrays(0 To 15) As Long
                        
                     For cCusGray = 240 To 15 Step -15
                        lGrays((cCusGray \ 15) - 1) = RGB(cCusGray, cCusGray, cCusGray)
                     Next
                        ReturnComponentValue wCaption, "LDefClr:", sDefValue
                        
                        With CC
                          .hInstance = App.hInstance
                          
                          If Val(sDefValue) > -1 Then
                            .Flags = CC_RGBINIT
                            .rgbResult = Val(sDefValue)
                          Else
                            .Flags = CC_ANYCOLOR
                          End If
                          
                          .Flags = .Flags Or CC_FULLOPEN
                          .lStructSize = Len(CC)
                          .hwndOwner = FormHwnd
                          .lpCustColors = VarPtr(lGrays(0))
                        End With
                        
                        If ChooseColor(CC) = 0 Then
                           MenuSelected = False
                        Else
                           oTipClass.SendCustomSelection sUserID, "Color", CC.rgbResult
                        End If
                    Else
                        oTipClass.SendCustomSelection sUserID, "Color", Val(sValue)
                    End If
                    
                Case "LDay:": oTipClass.SendCustomSelection sUserID, "WeekDay", Val(sValue)
                Case "LMonth:": oTipClass.SendCustomSelection sUserID, "Month", Val(sValue)
                Case "LState:": oTipClass.SendCustomSelection sUserID, "State", sValue
                Case "LDate:": oTipClass.SendCustomSelection sUserID, "DayOfMonth", CDate(sValue)
                Case "LFont:": oTipClass.SendCustomSelection sUserID, "Font", sValue
                Case "LDrv:": oTipClass.SendCustomSelection sUserID, "Drive", Left$(sValue, 3)
            End Select
            
            CopyMemory oTipClass, 0&, 4&
            
            Set oTipClass = Nothing
         Else
            If IsNumeric(vItems(mIndex).gControl) = False Then Exit Function
            If vItems(mIndex).gControl = 0 Then Exit Function
            If IsWindow(vItems(mIndex).gControl) = 0 Then Exit Function
            
            ReturnComponentValue wCaption, "LIndex:", sValue
            
            If vItems(mIndex).ControlType = 2 Then
               pParam2 = Val(sValue)
               
               If ((vItems(mIndex).Status And lv_mChk) <> lv_mChk) Then pParam1 = 1
            Else
               pParam1 = Val(sValue)
            End If
            
            If Err Then Exit Function
            
            Select Case vItems(mIndex).ControlType
               Case 0: msgParam = CB_SETCURSEL
               Case 1: msgParam = LB_SETCURSEL
               Case 2: msgParam = LB_SETSEL
            End Select
            
            With vItems(mIndex)
               SendMessage .gControl, msgParam, pParam1, ByVal pParam2
               SendMessage GetParent(.gControl), WM_COMMAND, MakeLong(CInt(GetWindowLong(.gControl, GWL_ID)), 1), ByVal .gControl
            End With
         End If
      End If
   Else
      MenuSelected = gMenus("g" & hMenu).MenuSelected(Index, hMenu, 1)
   End If
   
   Err.Clear
End Function

Public Sub GetMenuItem(ID As Long, hMenu As Long)
   Dim MenuData As MenuComponentData
   Dim Index As Long

   On Error Resume Next
   
   XferMenuData = MenuData
   Index = cItems(ID & "." & hMenu)
   
   If Err = 0 Then
      XferMenuData = vItems(Index)
   Else
      gMenus("g" & hMenu).GetMenuItem ID, hMenu
   End If

   Err.Clear
End Sub

Public Sub UpdateMenuItems(hMenu As Long)
   Dim MIS As MENUITEMINFO_string
   Dim Looper As Integer
   Dim pIndex As Long
   Dim mMode As Integer
   Dim bUpdated As Boolean
   Dim nrMenus As Integer
   Dim Index As Long

   On Error Resume Next
   
   pIndex = cPanels("p" & hMenu)
   
   If Err Then
      Err.Clear
      Exit Sub
   End If

   If vPanels(pIndex).SidebarMenuItem = 0 Then Exit Sub

   On Error GoTo AbortMenuUpdate

   Index = cItems(vPanels(pIndex).SidebarMenuItem & "." & hMenu)
   MIS.cbSize = Len(MIS)
   MIS.fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE
   MIS.fType = 0
   
   If Len(vItems(Index).Cached) Then
      MIS.dwTypeData = vItems(Index).Cached
   Else
      MIS.dwTypeData = vItems(Index).Caption
   End If
   
   MIS.cch = Len(MIS.dwItemData)
   MIS.dwItemData = hMenu
   MIS.wID = vItems(Index).ID

   If ((vItems(Index).Status And lv_mDisabled) = lv_mDisabled) Then MIS.fState = MF_DISABLED
   
   If ((vItems(Index).Status And lv_mSBarHidden) = lv_mSBarHidden) Then
      vItems(Index).Status = vItems(Index).Status And Not lv_mSBarHidden
   Else
      RemoveMenu hMenu, MIS.wID, 0
   End If

   If InsertMenuItem(hMenu, 0, True, MIS) Then bUpdated = True
   If bUpdated Then DrawMenuBar FormHwnd

AbortMenuUpdate:
End Sub

Public Sub GetPanelItem(hMenu As Long)
   Dim PanelInfo As PanelData
   Dim Index As Long
   
   On Error Resume Next

   XferPanelData = PanelInfo
   Index = cPanels("p" & hMenu)
   
   If Err = 0 Then
      XferPanelData = vPanels(Index)
   Else
      gMenus("g" & hMenu).GetPanelItem hMenu
   End If
   
   Err.Clear
End Sub

Public Function IsWindowList(hSubMenu As Long, bIsSysMenu As Boolean) As Boolean
   Dim lMenus As Long
   Dim lState As Long
   Dim Looper As Integer
   
   On Error Resume Next

   lMenus = gMenus("g" & hSubMenu).ShowTips
   
   If Err = 0 Then
      gMenus("g" & hSubMenu).IsWindowList hSubMenu, False
      Exit Function
   End If
   
   Err.Clear

   Dim OverallHeight As Long, Height As Long, MaxHeight As Long
   Dim MI() As Byte, Items2Check() As Long
   Dim MII As MENUITEMINFO
   Dim lType As Long, lStatus As Long, CachedMenus As Long
   Dim sCaption As String, sNoScroll As String
   Dim bBarBreak As Boolean, NewSubmenuID As Long
   Dim bSideBarItem As Boolean, bHasSideBar As Boolean
   Dim PageIDs() As Long, bNoOwnerDrawn As Boolean
   
   On Error Resume Next
   
   CachedMenus = modMenus.MenuCaptionListBox
   
   If CachedMenus Then
      If IsWindow(CachedMenus) = 0 Then CachedMenus = 0
   End If
   
   lMenus = GetMenuItemCount(hSubMenu)
   
   ReDim Items2Check(0 To lMenus)
   ReDim PageIDs(0 To 0)
   
   MaxHeight = Screen.Height / Screen.TwipsPerPixelY - 7
   
   For Looper = 0 To lMenus - 1
      bNoOwnerDrawn = False
      
      ReDim MI(0 To 1023)
      
      MII.cbSize = Len(MII)
      MII.fType = 0
      MII.fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU Or MIIM_DATA
      MII.dwTypeData = VarPtr(MI(0))
      MII.cch = UBound(MI)
      
      GetMenuItemInfo hSubMenu, Looper, True, MII
    
      If MII.wID >= &HF000& And MII.hSubMenu = 0 Then
         bIsSysMenu = True
      Else
         If MII.wID > &H7FFF And MII.hSubMenu = 0 Then IsWindowList = True
      End If
   
      Items2Check(Looper + 1) = MII.wID
      lStatus = 0
      
      If ((MII.fState And MF_DISABLED) = MF_DISABLED) Or ((MII.fState And MF_GRAYED) = MF_GRAYED) Then lStatus = lStatus Or lv_mDisabled
      If ((MII.fState And MF_CHECKED) = MF_CHECKED) Then lStatus = lStatus Or lv_mChk
      If ((MII.fState And MF_DEFAULT) = MF_DEFAULT) Then lStatus = lStatus Or lv_mDefault
      If MII.hSubMenu Then lStatus = lStatus Or lv_mSubmenu
      
      If ((MII.fType And MF_MENUBARBREAK) = MF_MENUBARBREAK) Then
         MII.fType = MII.fType And Not MF_MENUBARBREAK
         bNoOwnerDrawn = True
      End If
   
      If ((MII.fType And MF_MENUBREAK) = MF_MENUBREAK) Then
         MII.fType = MII.fType And Not MF_MENUBREAK
         bNoOwnerDrawn = True
      End If
      
      If bIsSysMenu Then bNoOwnerDrawn = True
      
      lType = MII.fType
      lState = MII.fState
      sCaption = Trim$(Replace$(StrConv(MI, vbUnicode), Chr$(0), ""))
   
      GetMenuMetrix sCaption, lStatus, lType, lState, hSubMenu, MII.wID, Looper, bSideBarItem, Height, NewSubmenuID, CachedMenus
      
      If bBarBreak Then
         ReDim Preserve PageIDs(0 To UBound(PageIDs) + 1)
      
         PageIDs(UBound(PageIDs)) = MII.wID
         bBarBreak = False
         OverallHeight = 0
         bNoOwnerDrawn = True
      End If
   
      If bSideBarItem And Not bHasSideBar Then
         bBarBreak = True
         PageIDs(0) = MII.wID
         bHasSideBar = True
         bNoOwnerDrawn = True
         Height = 0
      End If
   
      If Not bNoOwnerDrawn Then lType = lType Or MF_OWNERDRAW
      
      OverallHeight = OverallHeight + Height
      
      If OverallHeight > MaxHeight And bHasSideBar Then
        ReDim Preserve PageIDs(0 To UBound(PageIDs) + 1)
        
        PageIDs(UBound(PageIDs)) = MII.wID
        If UBound(PageIDs) = 2 Then MaxHeight = OverallHeight - Height
        OverallHeight = Height
      End If
      
      lFlag = OverallHeight
   
      MII.fMask = 0
      
      If NewSubmenuID <> 0 Or MII.hSubMenu <> 0 Then
         MII.fMask = MIIM_SUBMENU
         If NewSubmenuID <> 0 Then MII.hSubMenu = NewSubmenuID
      End If
     
      MII.dwItemData = hSubMenu
      MII.fState = lState
      MII.fType = lType
      MII.fMask = MII.fMask Or MIIM_ID Or MIIM_STATE Or MIIM_DATA Or MIIM_TYPE
      
      SetMenuItemInfo hSubMenu, MII.wID, False, MII
     
      If bNoOwnerDrawn Then
         MII.fType = MII.fType Or MF_OWNERDRAW
         SetMenuItemInfo hSubMenu, MII.wID, False, MII
      End If
   Next
   
   If bHasSideBar Then
      With vItems(cItems(PageIDs(0) & "." & hSubMenu))
         SeparateCaption .Caption, "", sCaption
         ReturnComponentValue sCaption, "NOSCROLL", sNoScroll
         
         If UBound(PageIDs) > 1 And Len(sNoScroll) > 0 Then
            .Status = .Status Or lv_mSBarHidden
            RemoveMenu hSubMenu, PageIDs(0), 0
            
            ReDim PageIDs(0 To 0)
         End If
      End With
   
      For Looper = 1 To UBound(PageIDs)
         ReDim MI(0 To 1023)
         
         MII.cch = UBound(MI)
         MII.fMask = MIIM_TYPE
         MII.fType = 0
         
         GetMenuItemInfo hSubMenu, PageIDs(Looper), False, MII
      
         If Looper = 1 Then
            MII.fType = MII.fType Or MF_OWNERDRAW Or MF_MENUBREAK
         Else
            MII.fType = MII.fType Or MF_OWNERDRAW Or MF_MENUBARBREAK
         End If
      
         SetMenuItemInfo hSubMenu, PageIDs(Looper), False, MII
      Next
   End If
   
   GetPanelMetrix hSubMenu, bIsSysMenu, Items2Check(), PageIDs()
   Erase Items2Check
   Erase PageIDs
   DrawMenuBar FormHwnd       ' not needed yet, next version will take control of the menu bar
End Function
   
Private Sub GetMenuMetrix(Caption As String, Status As Long, NewType As Long, NewState As Long, hMenu As Long, _
   MenuID As Long, menuPos As Integer, IsSideBar As Boolean, ItemHeight As Long, NewSubmenuID As Long, AltMenuSource As Long)
   Dim bChildClass As Boolean
   Dim IdX As Integer
   Dim NewIndex As Integer
   Dim wCaption As String
   Dim sTarget As String
   Dim sValue As String
   Dim CachedCaption As String
   Dim mComponent As MenuComponentData
   Dim frmObject As Control
   
   On Error Resume Next
   
   NewIndex = cItems(MenuID & "." & hMenu)
   IsSideBar = False
   NewSubmenuID = 0
   
   If AltMenuSource And Len(Caption) > 0 Then
      Dim sHotKey As String
      
      SeparateCaption Caption, CachedCaption, wCaption, sHotKey
      ReturnComponentValue wCaption, "Cache:", sValue
      
      If Len(sValue) Then
         mComponent.Cached = Caption
         IdX = SendMessage(AltMenuSource, LB_GETTEXTLEN, Val(sValue), ByVal 0&)
         sTarget = String$(IdX, 0)
         SendMessage AltMenuSource, LB_GETTEXT, Val(sValue), ByVal sTarget
         Caption = CachedCaption & StringFromBuffer(sTarget)
         
         If Len(sHotKey) Then Caption = Caption & vbTab & sHotKey
      End If
   End If
   
   If NewIndex > 0 Then
      mComponent = vItems(NewIndex)
      
      If StrComp(Caption, mComponent.Caption, vbBinaryCompare) = 0 Or Caption = "" Then
         Caption = mComponent.Caption
         SeparateCaption Caption, "", wCaption, ""
     
         If (((Status And lv_mDefault) = lv_mDefault) Or ((mComponent.Status And lv_mDefault) = (Status And lv_mDefault))) Then
            With mComponent
               If ((.Status And lv_mDisabled) = lv_mDisabled) Then .Status = .Status And Not lv_mDisabled
               If ((.Status And lv_mChk) = lv_mChk) Then .Status = .Status And Not lv_mChk
               If ((.Status And lv_mDefault) = lv_mDefault) Then .Status = .Status And Not lv_mDefault
               .Status = .Status Or Status
            End With
            
            If mComponent.hControl = 0 And ((mComponent.Status And lv_mCustom) <> lv_mCustom) Then
               If ((mComponent.Status And lv_mSBar) = lv_mSBar) Then
                  IsSideBar = True
                  mComponent.HotKey = ""
               Else
                  If ((mComponent.Status And lv_mSep) = lv_mSep) Then NewType = NewType Or MF_SEPARATOR
               End If
             
               ItemHeight = mComponent.Dimension.y
               vItems(NewIndex) = mComponent
               GoTo DoIconReference
            Else
               GoTo DoCustomMenus
            End If
         End If
      End If
   End If
   
   With mComponent
      .Caption = Caption
      .Index = menuPos
      .Status = Status
      SeparateCaption Caption, .Display, wCaption, sValue
      
'      If .Display = "&Restore" Then .Display = "No&rmal"
'      If .Display = "&Move" Then .Display = "&Pindah"
'      If .Display = "&Size" Then .Display = "&Ukuran"
'      If .Display = "Mi&nimize" Then .Display = "Mi&nimal"
'      If .Display = "Ma&ximize" Then .Display = "Ma&ksimal"
'      If .Display = "&Close" Then .Display = "&Tutup"
      
      .Display = Replace$(.Display, "& ", "&& ")
      
      If Len(sValue) Then .HotKey = sValue
      
      ReturnComponentValue wCaption, "TIP:", sValue
      
      If Len(sValue) Then
         .Tip = sValue
      Else
         .Tip = tipDefault
      End If
      
      ReturnComponentValue wCaption, "SIDEBAR", sValue
      
      If Len(sValue) Then
         If menuPos = 0 Then
            .Status = 0
            
            If ((NewType And MF_SEPARATOR) = MF_SEPARATOR) Then NewType = NewType And Not MF_SEPARATOR
            If ((NewState And MF_DISABLED) = MF_DISABLED) Then .Status = .Status Or lv_mDisabled
            If ((NewState And MF_CHECKED) = MF_CHECKED) Then .Status = .Status Or lv_mChk
         
            .Status = .Status Or lv_mSBar
            ReturnComponentValue wCaption, "SBDisabled", sValue
            
            If Len(sValue) Then
               .Status = .Status Or lv_mDisabled
               NewState = NewState Or MF_DISABLED
            End If
            
            .HotKey = "Remeasure"
            IsSideBar = True
            GoTo Check4ControlReference
         Else
            .Display = ""
         End If
      End If
      
      If Len(.Display) Then
         If Left$(.Display, 1) = "-" Then
            .Display = Mid$(.Display, 2)
            NewType = NewType Or MF_SEPARATOR
         End If
      Else
         NewType = NewType Or MF_SEPARATOR
      End If
   
      If ((NewType And MF_SEPARATOR) = MF_SEPARATOR) Then
         .Status = 0
         NewState = 0
         ReturnComponentValue wCaption, "RAISED", sValue
         
         If Len(sValue) Then .Status = .Status Or lv_mSepRaised
         
         .Status = .Status Or lv_mSep Or lv_mDisabled
         .HotKey = ""
         .Icon = ""
         .hControl = 0
         NewState = NewState Or MF_DISABLED
         GoTo GetMeasurements
      Else
         ReturnComponentValue wCaption, "DEFAULT", sValue
         
         If Len(sValue) Then .Status = .Status Or lv_mDefault
      End If
      
      If Len(wCaption) = 0 Then GoTo GetMeasurements
         ReturnComponentValue wCaption, "HotKey:", sValue
         
         If Len(sValue) Then .HotKey = sValue
         
         If ((.Status And lv_mSubmenu) = lv_mSubmenu) And .hControl <> 0 Or ((.Status And lv_mSubmenu) <> lv_mSubmenu) Then
            ReturnComponentValue wCaption, "LB:", sValue
            
            If Len(sValue) Then
               .ControlType = 1
            Else
               ReturnComponentValue wCaption, "CB:", sValue
            End If
         
            If Len(sValue) Then
               If IsNumeric(sValue) Then
                  If IsWindow(CLng(sValue)) Then .hControl = CLng(sValue)
               Else
                  Set frmObject = SplitControlIndex(sValue)
                  
                  If Not frmObject Is Nothing Then
                  .hControl = frmObject.hwnd
                  
                  Set frmObject = Nothing
               End If
            End If
         End If
      End If
   
Check4ControlReference:
      ReturnComponentValue wCaption, "gControl:", sValue
      
      If Len(sValue) Then
         bChildClass = True
         .gControl = Val(sValue)
         ReturnComponentValue wCaption, "gType:", sValue
         
         If Len(sValue) Then .ControlType = Val(sValue)
         ReturnComponentValue wCaption, "LColor:", sValue
        
         If Len(sValue) Then .Status = .Status Or lv_mColor
         ReturnComponentValue wCaption, "LFont:", sValue
        
         If Len(sValue) Then .Status = .Status Or lv_mFont
         ReturnComponentValue wCaption, "LState:", sValue
        
         If Len(sValue) Then .HotKey = sValue
         ReturnComponentValue wCaption, "gEXE:", sValue
        
         If Len(sValue) Then
            .Icon = GetShortFileName(sValue & .Display)
            sValue = StripFile(.Display, "E")
            .Display = StripFile(.Display, "m")
            
            If StrComp(LCase(.Display), .Display, vbBinaryCompare) = 0 Then .Display = UCase(.Display)
            
            .Display = .Display & "." & LCase(sValue)
            .Tip = GetFileDescription(.Icon)
            
            If .Tip = "" Then .Tip = tipDefault
         End If
         
         ReturnComponentValue wCaption, "LDrv:", sValue
         
         If Len(sValue) Then
            IdX = InStr(sValue, "\")
            
            If IdX Then .Icon = Left$(sValue, IdX) Else .Icon = sValue
            
            If .Tip = "Removable Media" Then
               .Icon = GetFloppyIcon(.Icon)
               .Display = .Display & " Removable"
            Else
               If Len(.Icon) > IdX Then IdX = 1 Else IdX = 0
               sTarget = String$(255, 0)
               sValue = String$(55, 0)
               GetVolumeInformation .Icon, sTarget, 255, 0, 0, 0, sValue, 55
               sTarget = StringFromBuffer(sTarget)
               
               If Len(sTarget) Then
                  If .Tip = "Network/Mapped Drive" Then
                     .Display = .Display & "  (" & sTarget & ")"
                  Else
                     .Display = .Display & " " & sTarget
                  End If
               End If
            End If
         End If
      End If
   
      If ((.Status And lv_mSubmenu) = lv_mSubmenu) Then
         .HotKey = ""
      End If
   End With
   
GetMeasurements:
   If ((mComponent.Status And lv_mSBar) = lv_mSBar) = False Then
      Dim tDC As Long
      Dim tRect As RECT
      
      tDC = GetDC(FormHwnd)
   
      With mComponent
         If ((.Status And lv_mSep) = lv_mSep) Then
            ApplyMenuFont 3, tDC
         Else
            ApplyMenuFont 1 + Abs(CInt((.Status And lv_mDefault) = lv_mDefault)), tDC
         End If
      
         If ((.Status And lv_mFont) = lv_mFont) Then
            sTarget = .Display
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            .Dimension.x = tRect.Right
            .Dimension.y = tRect.Bottom
            ApplyMenuFont 0, tDC
            ReturnComponentValue wCaption, "LFont:", sValue
            CreateDestroyMenuFont True, False, sValue
            ApplyMenuFont 6, tDC
            sTarget = .HotKey
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            tRect.Right = tRect.Right + .Dimension.x
         
            If tRect.Bottom > .Dimension.y Then tRect.Bottom = .Dimension.y
         Else
            sTarget = .Display & .HotKey
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
         End If
     
         If ((.Status And lv_mSep) = lv_mSep) = True And Len(.Display) = 0 Then
            tRect.Bottom = 4
            .Dimension.x = tRect.Right
         Else
            If ((.Status And lv_mSep) <> lv_mSep) Then
               If tRect.Bottom < 20 Then tRect.Bottom = 20
            End If
         
            .Dimension.x = tRect.Right + 40
         End If
     
         .Dimension.y = tRect.Bottom + 2
         ApplyMenuFont 0, tDC
      End With
   
      ReleaseDC FormHwnd, tDC
   End If
   
DoCustomMenus:
   With mComponent
      For IdX = 1 To 7
         sTarget = Choose(IdX, "lvColors:", "lvMonths:", "lvDays:", "lvStates:", "lvMonth:", "lvFonts:", "lvDrives:")
         ReturnComponentValue wCaption, sTarget, sValue
         
         If Len(sValue) Then Exit For
      Next
      
      If Len(sValue) Then
         CreateSubMenuCustom MenuID, menuPos, hMenu, NewSubmenuID, sTarget, sValue, .Tip
         .Status = .Status Or lv_mCustom
      Else
         If .hControl Then
            CreateSubMenu MenuID, menuPos, hMenu, .ControlType, .hControl, NewSubmenuID, wCaption
         End If
      End If
      
      If NewSubmenuID Then
         .Status = .Status Or lv_mSubmenu
         
         If Len(.HotKey) Then
            sTarget = .HotKey
            tDC = GetDC(FormHwnd)
            ApplyMenuFont 1, tDC
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            .Dimension.x = .Dimension.x - tRect.Right
            ApplyMenuFont 0, tDC
            ReleaseDC FormHwnd, tDC
         End If
         
         .HotKey = ""
      End If
   End With

DoIconReference:
   With mComponent
      .ShowBKG = False
      
      If bChildClass = False And ((.Status And lv_mSep) <> lv_mSep) Then
         .Icon = CStr(ValidateImage(wCaption, IdX))
      
         If .Icon = "0" Then
            .Icon = ""
         Else
            ReturnComponentValue wCaption, "IMGBKG", sValue
            If Len(sValue) And IdX = 1 Then .ShowBKG = True
            End If
      End If
   End With

   If NewIndex = 0 Then
      NewIndex = UBound(vItems) + 1
   
      ReDim Preserve vItems(0 To NewIndex)
   
      cItems.Add NewIndex, MenuID & "." & hMenu
   End If

   mComponent.ID = MenuID
   vItems(NewIndex) = mComponent

   If NewSubmenuID Then cItems.Add NewIndex, NewSubmenuID & "." & hMenu
   
   ItemHeight = mComponent.Dimension.y
End Sub
   
Private Sub GetPanelMetrix(hMenu As Long, bIsSysMenu As Boolean, Items2Check() As Long, PageBreaks() As Long)
   Dim Looper As Integer
   Dim lItem As Long
   Dim NewIndex As Long
   Dim NextPage As Long
   Dim IdxStart As Long
   Dim IdxStop As Long
   Dim tIndex As Integer
   Dim lSideBar As Long
   Dim Xoffset As Integer
   Dim Win98MEoffset As Integer
   Dim pData As PanelData
   Dim MaxHeight As Integer
   Dim MaxWidth As Integer
   Dim bHasTabs As Boolean
   Dim bHasSubs As Boolean
   Dim bHasHotKey As Boolean
   Dim bHideSidebarCheck As Boolean
   Dim bNonSbar As Boolean
   Dim bReprocessSideBar As Boolean
   
   On Error Resume Next
   
   Win98MEoffset = modMenus.Win98MEoffset
   NewIndex = cPanels("p" & hMenu)
   pData.SubmenuID = hMenu
   
   If Err Then
      NewIndex = 0
      Err.Clear
   Else
      pData.PanelIcon = vPanels(NewIndex).PanelIcon
   End If
   
   IdxStart = 1
   IdxStop = UBound(Items2Check)
   Items2Check(0) = 0
   
   If UBound(PageBreaks) > 1 Then NextPage = 2
   If UBound(PageBreaks) > 1 Then NextPage = 2
   
   For Looper = 1 To UBound(Items2Check)
      tIndex = cItems(Items2Check(Looper) & "." & hMenu)
      
      With vItems(tIndex)
         If ((.Status And lv_mSBar) <> lv_mSBar) Then
            If (Items2Check(Looper) = PageBreaks(NextPage)) And NextPage > 1 Then
               IdxStop = Looper - 1
               GoSub UpdateItemOffsets
               bHasSubs = False
               bHasHotKey = False
               MaxWidth = 0
               
               If NextPage = 2 Then Items2Check(0) = MaxHeight
               If NextPage > UBound(PageBreaks) Then NextPage = 0 Else NextPage = NextPage + 1
               
               IdxStart = Looper
               IdxStop = UBound(Items2Check)
               MaxHeight = 0
            End If
            
            If Len(.Icon) Then pData.HasIcons = True
            If ((.Status And lv_mSubmenu) = lv_mSubmenu) Then bHasSubs = True
            If .ID >= &HF000& And ((.Status And lv_mSubmenu) <> lv_mSubmenu) Then pData.IsSystem = True
            If Len(.HotKey) Then bHasHotKey = True
            If .Dimension.x > MaxWidth Then MaxWidth = .Dimension.x
         
            MaxHeight = MaxHeight + .Dimension.y
         End If
      End With
   Next
   
   If Items2Check(0) > 0 Then MaxHeight = Items2Check(0)
   If NextPage = 0 Then GoSub UpdateItemOffsets
   If bIsSysMenu Then pData.IsSystem = True
   
   If NewIndex = 0 Then
      NewIndex = UBound(vPanels) + 1
      
      ReDim Preserve vPanels(0 To NewIndex)
      
      cPanels.Add NewIndex, "p" & hMenu
   End If
   
   If UBound(Items2Check) > 50 Then pData.Hourglass = True
   
   vPanels(NewIndex) = pData
   
   If bReprocessSideBar = True Then
      lItem = cItems(PageBreaks(0) & "." & hMenu)
      
      Dim wCaption As String
      Dim sValue As String
   
      SeparateCaption vItems(lItem).Caption, "", wCaption
      ReturnComponentValue wCaption, "IMG:", sValue
      
      If Len(sValue) Then
         CreateImageSidebar lItem, NewIndex, wCaption
      Else
         CreateTextSidebar lItem, NewIndex, wCaption
      End If
      
      With vItems(lItem)
         .OffsetCx = -12
         
         If InStr(.Caption, vbTab) Then .OffsetCx = .OffsetCx - Win98MEoffset
      End With
   End If
   
   Exit Sub
   
UpdateItemOffsets:
   Xoffset = 0
   
   If bHasHotKey Then
      If Win98MEoffset = 0 Then Xoffset = 15
      If bHasSubs = True And Win98MEoffset = 0 Then Xoffset = Xoffset + 20
   End If
   
   If pData.IsSystem = True And Win98MEoffset = 0 Then Xoffset = Xoffset + 15
   
   For lItem = IdxStart To IdxStop
      tIndex = cItems(Items2Check(lItem) & "." & hMenu)
      
      With vItems(tIndex)
         If ((.Status And lv_mSBar) = lv_mSBar) Then
            pData.SidebarMenuItem = .ID
            
            If .Dimension.y <> MaxHeight Then bReprocessSideBar = True
            .Dimension.y = MaxHeight
            
            If .HotKey = "Remeasure" Then bReprocessSideBar = True
            .HotKey = ""
            pData.Accelerators = pData.Accelerators & Chr$(246)
         Else
            If IsNumeric(.Icon) = False And Len(.Icon) > 0 Then pData.Hourglass = True
            If InStr(.Caption, "lFont:") > 0 And UBound(Items2Check) > 20 Then pData.Hourglass = True
            .OffsetCx = Xoffset
            If ((.Status And lv_mColor) = lv_mColor) Then .OffsetCx = .OffsetCx + 10
            
            If ((.Status And lv_mSep) = lv_mSep) Then
               pData.Accelerators = pData.Accelerators & Chr$(246)
            Else
               pData.Accelerators = pData.Accelerators & ReturnAcceleratorKey(.Display)
            End If
         End If
      End With
   Next
   
   Return
End Sub
   
Private Function ReturnAcceleratorKey(sInKey As String) As String
   Dim i As Integer
   Dim Index As Integer
   
   i = InStr(Index + 1, sInKey, "&")
   
   Do While i
      If Mid$(sInKey, i, 2) = "&&" Then
         Index = i + 1
         i = InStr(Index + 1, sInKey, "&")
      Else
         Exit Do
      End If
   Loop
   
   If i Then
      ReturnAcceleratorKey = UCase(Mid$(sInKey, i + 1, 1))
   Else
      ReturnAcceleratorKey = Chr$(246)
   End If
End Function
Private Sub CreateSubMenu(MenuID As Long, mnuPos As Integer, ParentMenu As Long, ControlType As Byte, hControl As Long, NewSubmenuID As Long, sCustomFlag As String)
   If hControl = 0 Then Exit Sub
   
   Dim msgListCount As Long
   Dim msgListIndex As Long
   Dim msgTextLen As Long
   Dim msgText As Long
   Dim msgItemData As Long
   Dim ItemData As Long
   Dim ListCount As Long
   Dim ListIndex As Long
   Dim lLen As Long
   Dim lStyle As Long
   Dim Looper As Integer
   Dim SelCount As Long
   Dim bMultiSelect As Boolean
   Dim bShowHourglass As Boolean
   Dim sBuffer As String
   Dim sParent As String
   Dim sIcon As String
   
   Const CB_GETCOUNT As Long = &H146
   Const CB_GETLBTEXT As Long = &H148
   Const CB_GETLBTEXTLEN As Long = &H149
   Const CB_GETCURSEL As Long = &H147
   Const CB_GETITEMDATA As Long = &H150
   Const CBS_OWNERDRAWFIXED As Long = &H10&
   Const CBS_OWNERDRAWVARIABLE As Long = &H20&
   Const LBS_MULTIPLESEL As Long = &H8&
   Const LBS_EXTENDEDSEL As Long = &H800&
   Const LB_GETSEL As Long = &H187
   Const LB_GETCURSEL As Long = &H188
   Const LB_GETCOUNT As Long = &H18B
   Const LB_GETITEMDATA As Long = &H199
   Const LBS_MULTICOLUMN As Long = &H200&
   
   lStyle = GetWindowLong(hControl, GWL_STYLE)
   
   If ControlType Then
      If ((lStyle And LBS_MULTICOLUMN) = LBS_MULTICOLUMN) Then Exit Sub
      
      bMultiSelect = (((lStyle And LBS_MULTIPLESEL) = LBS_MULTIPLESEL) Or ((lStyle And LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL))
   Else
      If ((lStyle And CBS_OWNERDRAWVARIABLE) = CBS_OWNERDRAWVARIABLE) Then Exit Sub
      If ((lStyle And CBS_OWNERDRAWFIXED) = CBS_OWNERDRAWFIXED) Then Exit Sub
   End If
   
   NewSubmenuID = GetSubMenu(ParentMenu, mnuPos)
   
   On Error Resume Next
   
   If NewSubmenuID Then
      gMenus.Remove "g" & NewSubmenuID
      DestroyMenu NewSubmenuID
      cItems.Remove NewSubmenuID & "." & ParentMenu
      Err.Clear
   End If
   
   NewSubmenuID = CreatePopupMenu()
   
   If NewSubmenuID = 0 Then Exit Sub
   sParent = "gControl:" & hControl & "|LIndex:x"
   ReturnComponentValue sCustomFlag, "Files:", sBuffer
   
   If Len(sBuffer) > 0 And ControlType <> 0 Then
      If sBuffer = "-1" Then
         sIcon = "Paths Included"
      Else
         If Right$(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
         
         If Len(Dir(sBuffer, vbDirectory)) Then
            sParent = sParent & "|gEXE:" & sBuffer
            bShowHourglass = True
         Else
            sIcon = "Paths Included"
         End If
      End If
   End If
   
   sParent = sParent & "|gType:" & ControlType + Abs(CInt(bMultiSelect))
   msgListCount = Choose(ControlType + 1, CB_GETCOUNT, LB_GETCOUNT)
   msgListIndex = Choose(ControlType + 1, CB_GETCURSEL, LB_GETCURSEL)
   msgTextLen = Choose(ControlType + 1, CB_GETLBTEXTLEN, LB_GETTEXTLEN)
   msgText = Choose(ControlType + 1, CB_GETLBTEXT, LB_GETTEXT)
   msgItemData = Choose(ControlType + 1, CB_GETITEMDATA, LB_GETITEMDATA)
   
   ListCount = SendMessage(hControl, msgListCount, 0&, ByVal 0&) - 1
   ListIndex = SendMessage(hControl, msgListIndex, 0, ByVal 0&)
   
   For Looper = 0 To ListCount
      lLen = SendMessage(hControl, msgTextLen, Looper, ByVal 0&)
      sBuffer = Space$(lLen) & Chr$(0)
      SendMessage hControl, msgText, Looper, ByVal sBuffer
      lStyle = 0
     
      If bMultiSelect Then
         If SendMessage(hControl, LB_GETSEL, Looper, ByVal 0&) > 0 Then lStyle = MF_CHECKED
      Else
         If Looper = ListIndex Then lStyle = MF_CHECKED
      End If
     
      sBuffer = Trim(Replace$(sBuffer, Chr$(0), ""))
      
      If InStr(sBuffer, "{") Then
         sBuffer = Left$(sBuffer, InStr(sBuffer, "{")) & sParent & "|" & Mid(sBuffer, InStr(sBuffer, "{") + 1)
      Else
         If Len(sIcon) Then
            sIcon = StripFile(GetShortFileName(sBuffer), "P")
            sBuffer = StripFile(sBuffer, "N")
            
            If Len(Dir(sIcon & sBuffer, vbDirectory)) = 0 Then
               sIcon = "|LaVolpe2"
            Else
               sIcon = "|gEXE:" & sIcon
               bShowHourglass = True
            End If
         End If
         
         sBuffer = sBuffer & "{" & sParent & sIcon & "}"
      End If
     
      sBuffer = Replace$(sBuffer, "LIndex:x", "LIndex:" & Looper)
      AppendMenu NewSubmenuID, lStyle, 32500 - Looper, Replace$(sBuffer, "&", "&&")
   Next
   
   If ListCount = -1 Then
      lStyle = MF_DISABLED Or MF_GRAYED
      AppendMenu NewSubmenuID, lStyle, 32500, "-No Selections Available{" & sParent & "|Raised}"
   End If
   
   Dim myMenus As clsMenuItems
   
   Set myMenus = New clsMenuItems
   
   If bShowHourglass Then lStyle = NewSubmenuID Else lStyle = 0
   ReturnComponentValue sCustomFlag, "Tip:", sBuffer
   myMenus.InitializeSubMenu FormHwnd, imgList, TipCallBackProc, True, lStyle, sBuffer
   myMenus.hPrevProc = PrevProc
   gMenus.Add myMenus, "g" & NewSubmenuID
   
   Set myMenus = Nothing
End Sub
   
Private Sub CreateSubMenuCustom(MenuID As Long, mnuPos As Integer, ParentMenu As Long, NewSubmenuID As Long, CustomTag As String, CurCustomVal As String, ParentTip As String)
   Dim lStyle As Long
   Dim lValue As Long
   Dim lOffset As Long
   Dim Looper As Integer
   Dim CheckValue As Long
   Dim sCheckValue As String
   Dim sText As String
   Dim sParent As String
   Dim CustomID As String
   Dim bShowHourglass As Boolean
   
   NewSubmenuID = GetSubMenu(ParentMenu, mnuPos)
   
   On Error Resume Next
   
   If NewSubmenuID Then
      gMenus.Remove "g" & NewSubmenuID
      DestroyMenu NewSubmenuID
      cItems.Remove NewSubmenuID & "." & ParentMenu
      Err.Clear
   End If
   
   NewSubmenuID = CreatePopupMenu()
   
   If NewSubmenuID = 0 Then Exit Sub
   ReturnComponentValue CurCustomVal & "|", "ID:", CustomID
   lOffset = InStr(CustomID, ":")
   
   If lOffset Then CustomID = Left$(CustomID, lOffset - 1)
   If Len(CustomID) Then CurCustomVal = Replace$(CurCustomVal, ":ID:" & CustomID, "")
   
   sParent = "{gControl:0|lvID:" & CustomID
   
   Select Case CustomTag
      Case "lvDrives:"
         Dim vDrives() As String
         Dim sTip As String
     
         lOffset = InStr(CurCustomVal, ":")
         
         If lOffset > 1 Then sCheckValue = Left(CurCustomVal, lOffset) Else sCheckValue = CurCustomVal
         ListAllDrives vDrives
         
         If UBound(vDrives) Then bShowHourglass = True
     
         For Looper = 1 To UBound(vDrives)
            CustomID = Left(vDrives(Looper), InStr(vDrives(Looper), "|") - 1)
            lOffset = InStr(CustomID, "\")
            
            If lOffset = 0 Then
               lOffset = InStr(CustomID, ":")
               
               If lOffset = 0 Then lOffset = Len(CustomID)
            End If
            
            sText = Left(CustomID, lOffset)
         
            If Right$(sText, 1) <> "\" Then sText = sText & "\"
            If InStr(sText, ":") Then sText = "[" & UCase(Left$(sText, InStr(sText, ":") - 1)) & "]"
         
            sText = sText & Mid$(CustomID, lOffset + 1)
            
            If Left$(sText, Len(sCheckValue)) = sCheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            sText = sText & sParent & "|LDrv:" & CustomID & "|Tip:" & Mid$(vDrives(Looper), InStr(vDrives(Looper), "|") + 1) & "}"
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
         Next
      
      Case "lvFonts:"
         Dim vFonts() As String
         Dim sLtrA As String
         Dim sLtrZ As String
         Dim bInclude As Boolean
         
         ReturnComponentValue CurCustomVal & "|", "Type:", sText
         
         If Len(sText) Then lOffset = InStr(sText, ":")
         If lOffset Then sText = Left$(sText, lOffset - 1)
         
         Select Case sText
            Case "TrueType": lStyle = TRUETYPE_FONTTYPE
            Case "System": lStyle = RASTER_FONTTYPE
            Case Else: lStyle = 0
         End Select
     
         lOffset = InStr(CurCustomVal, ":")
         
         If lOffset Then sCheckValue = Left(CurCustomVal, lOffset - 1)
         ReturnComponentValue CurCustomVal & "|", "Group:", CustomTag
         lOffset = InStr(CustomTag, ":")
         
         If lOffset Then CustomTag = Left(CustomTag, lOffset - 1)
         
         If Len(CustomTag) = 3 Then
            sLtrA = Left(CustomTag, 1)
            sLtrZ = Right(CustomTag, 1)
         End If
         
         LoadFontMenu vFonts(), lStyle
         lOffset = 1
         
         For Looper = 1 To UBound(vFonts)
            sText = StrConv(vFonts(Looper), vbProperCase)
            bInclude = False
            
            If Len(CustomTag) Then
               If Left$(vFonts(Looper), 1) >= sLtrA And Left$(vFonts(Looper), 1) <= sLtrZ Then bInclude = True
            Else
               bInclude = True
            End If
         
            If bInclude = True Then
               sText = sText & sParent & "|LFont:" & vFonts(Looper) & "}" & vbTab & "Sample"
               
               If sCheckValue = vFonts(Looper) Then lStyle = MF_CHECKED Else lStyle = 0
               AppendMenu NewSubmenuID, lStyle, 32500 - lOffset, sText
               lOffset = lOffset + 1
            End If
         Next
      
         If UBound(vFonts) = 0 Or lOffset = 1 Then AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-No Fonts Available{Raised}"
         If lOffset > 20 Then bShowHourglass = True
         
         Erase vFonts
      
      Case "lvColors:"
         If Len(CurCustomVal) > 0 And Val(CurCustomVal) <> -1 Then
            CheckValue = ExchangeVBcolor(CurCustomVal, -1)
         Else
            CheckValue = -1
         End If
     
         For Looper = 1 To 23
            sText = Choose(Looper, "Black", "Blue", "Brown", "Cyan / Aqua", "Gold", "Gray", "Green / Lime", "Green Dark / Olive", "Ivory", "Magenta / Fuchsia", "Maroon", "Navy", "Orange", "Peach", "Purple", "Red", "Silver / Light Gray", "Tan", "Teal", "Turquoise", "Violet", "White", "Yellow")
            lValue = Choose(Looper, vbBlack, vbBlue, vbBrown, vbCyan, vbGold, vbGray, vbGreen, vbOlive, vbIvory, vbMagenta, vbMaroon, vbNavy, vbOrange, vbPeach, vbPurple, vbRed, vbSilver, vbTan, vbTeal, vbTurquoise, vbViolet, vbWhite, vbYellow)
            sText = sText & sParent & "|LColor:" & lValue & "}"
            
            If lValue = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
         Next
     
         AppendMenu NewSubmenuID, MF_DISABLED, 32500 - Looper, "-{Raised}"
         sText = "Select a custom color" & sParent & "|LDefClr:" & CheckValue & "|LColor:-1}"
         AppendMenu NewSubmenuID, 0, 32500 - Looper - 1, sText
      
      Case "lvMonths:"
         CheckValue = Val(CurCustomVal)
         
         If CheckValue = 0 Then CheckValue = Month(Date)
         ReturnComponentValue CurCustomVal & "|", "Group:", CustomTag
         lOffset = InStr(CustomTag, ":")
         
         If lOffset Then CustomTag = Left$(CustomTag, lOffset - 1)
         
         Select Case CustomTag
            Case "CYQtr"
               AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-First Quarter{Raised}"
            
            Case "FYQtr"
               AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-1st Fiscal Quarter{Raised}"
         
         End Select
     
         lValue = 1
         
         For Looper = 1 To 12
            If CustomTag = "FYQtr" Then
               lOffset = Choose(Looper, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9)
            Else
               lOffset = Looper
            End If
         
            If (Looper - 1) Mod 3 = 0 And Looper > 1 Then
               Select Case CustomTag
                  Case "CYQtr"
                     sText = Choose((Looper - 1) / 3, "Second", "Third", "Fourth") & " Quarter{Raised}"
                     AppendMenu NewSubmenuID, MF_DISABLED, 32500 - lValue, "-" & sText
                     lValue = lValue + 1
                  
                  Case "FYQtr"
                     sText = Choose((Looper - 1) / 3, "2nd", "3rd", "4th") & " Fiscal Quarter{Raised}"
                     AppendMenu NewSubmenuID, MF_DISABLED, 32500 - lValue, "-" & sText
                     lValue = lValue + 1
               End Select
            End If
         
            If lOffset = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            sText = Format(DateSerial(1962, lOffset, 7), "mmmm") & sParent & "|LMonth:" & lOffset & "}"
            AppendMenu NewSubmenuID, lStyle, 32500 - lValue, sText
            lValue = lValue + 1
         Next
      
      Case "lvDays:"
         CheckValue = Val(CurCustomVal)
         
         If CheckValue = 0 Then CheckValue = Weekday(Date)
         lOffset = Weekday(Date) - 1
         
         If lOffset Then lOffset = 7 - lOffset
         
         For Looper = 1 To 7
            sText = Format(DateAdd("d", lOffset, Date), "dddd")
            
            If Looper = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            sText = sText & sParent & "|LDay:" & Looper & "}"
            lOffset = lOffset + 1
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
         Next
      
      Case "lvStates:"
         Dim sAbbr As String
         
         If Len(CurCustomVal) > 2 Then
            sCheckValue = Left$(CurCustomVal, 2)
         Else
            If Len(CurCustomVal) < 2 Then sCheckValue = "LaVolpe" Else sCheckValue = CurCustomVal
         End If
         
         bShowHourglass = True
         
         For Looper = 1 To 51
            Select Case Looper
               Case Is < 11
                  sText = Choose(Looper, "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida")
                  sAbbr = Choose(Looper, "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL")
                  
               Case Is < 21
                  sText = Choose(Looper - 10, "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine")
                  sAbbr = Choose(Looper - 10, "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME")
                  
               Case Is < 31
                  sText = Choose(Looper - 20, "Maryland", "Massachusettes", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire")
                  sAbbr = Choose(Looper - 20, "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH")
                  
               Case Is < 41
                  sText = Choose(Looper - 30, "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island")
                  sAbbr = Choose(Looper - 30, "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI")
                  
               Case Is < 51
                  sText = Choose(Looper - 40, "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin")
                  sAbbr = Choose(Looper - 40, "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI")
                  
               Case Is < 61
                  sText = Choose(Looper - 50, "Wyoming")
                  sAbbr = Choose(Looper - 50, "WY")
                  
            End Select
         
            sText = sText & sParent & "|LState:" & sAbbr & "}"
            
            If sAbbr = sCheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
         Next
      
      Case "lvMonth:"
         Dim lvYear As Long
         Dim lDate As Long
         
         lValue = Val(CurCustomVal)
         
         If lValue = 0 Then lValue = Month(Date)
         
         If lValue > 0 And lValue < 13 Then
            ReturnComponentValue CurCustomVal & "|", "Year:", sText
            
            If Len(sText) Then lvYear = Val(sText) Else lvYear = Year(Date)
            If lvYear < 1 Then lvYear = Year(Date)
            ReturnComponentValue CurCustomVal & "|", "Day:", sText
            CheckValue = Val(sText)
            
            If CheckValue = 0 Then CheckValue = Day(Date)
            
            For Looper = 1 To Day(DateSerial(lvYear, lValue + 1, 0))
               lDate = CDate(DateSerial(lvYear, lValue, Looper))
               sText = Format(lDate, "mmm d, dddd")
               sText = sText & sParent & "|LDate:" & lDate & "}"
               
               If Looper = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
               AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
            Next
         End If
      
   End Select
   
   Dim myMenus As clsMenuItems
   
   Set myMenus = New clsMenuItems
   
   If bShowHourglass Then lValue = NewSubmenuID Else lValue = 0
   myMenus.InitializeSubMenu FormHwnd, imgList, TipCallBackProc, True, lValue, ParentTip
   gMenus.Add myMenus, "g" & NewSubmenuID
   
   Set myMenus = Nothing
End Sub
   
Private Function FormIndex() As Integer
   Dim i As Integer
   
   For i = Forms.Count - 1 To 0 Step -1
      If Forms(i).hwnd = FormHwnd Then
         FormIndex = i
         Exit For
      End If
   Next
End Function
   
Private Function SplitControlIndex(sControlName As String) As Control
   If Len(sControlName) = 0 Then Exit Function
   
   Dim Index As Long
   Dim sName As String
   
   sControlName = Trim$(sControlName)
   
   If Right$(sControlName, 1) = ")" Then
      Index = Val(Mid$(sControlName, InStrRev(sControlName, "(") + 1))
      sName = Left$(sControlName, InStrRev(sControlName, "(") - 1)
      
      Set SplitControlIndex = Forms(FormIndex).Controls(sName)(Index)
   Else
      Set SplitControlIndex = Forms(FormIndex).Controls(sControlName)
   End If
End Function
   
Private Sub Class_Terminate()
   Dim mIndex As Long
   
   On Error Resume Next
   
   For mIndex = 1 To UBound(vPanels)
      If vPanels(mIndex).PanelIcon Then
         DeleteObject vPanels(mIndex).PanelIcon
      End If
   Next
   
   Erase vItems
   Erase vPanels
   
   Set imgList = Nothing
   Set gMenus = Nothing
   Set cItems = Nothing
   Set cPanels = Nothing
End Sub
   
Private Sub Class_Initialize()
   Set gMenus = New Collection
   Set cItems = New Collection
   Set cPanels = New Collection
   
   ReDim vItems(0 To 0)
   ReDim vPanels(0 To 0)
   
   uMinMax.ptMaxSize.x = -1
End Sub
   
Public Sub RestoreMenus()
   Dim mIndex As Integer
   Dim Looper As Integer
   Dim hMenu As Long
   Dim typeRemoval As Long
   Dim nrItems As Long
   Dim mItem As Long
   Dim MII As MENUITEMINFO
   Dim MIS As MENUITEMINFO_string
   Dim MI() As Byte
   Dim mRestore() As Long
   Dim mSubMenus() As Long
   
   On Error Resume Next
   
   For Looper = 1 To UBound(vPanels)
      hMenu = vPanels(Looper).SubmenuID
      
      If IsMenu(hMenu) Then
         ReDim mRestore(0)
         ReDim mSubMenus(0)
      
         nrItems = GetMenuItemCount(hMenu)
         
         For mIndex = 0 To nrItems - 1
            ReDim MI(0 To 1023)
            
            MII.cbSize = Len(MII)
            MII.fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
            MII.fType = 0
            MII.dwTypeData = VarPtr(MI(0))
            MII.cch = UBound(MI)
         
            GetMenuItemInfo hMenu, mIndex, True, MII
            
            If ((MII.fType And MF_OWNERDRAW) = MF_OWNERDRAW) Then
               mItem = cItems(MII.wID & "." & hMenu)
               
               ReDim Preserve mRestore(0 To UBound(mRestore) + 1)
               
               mRestore(UBound(mRestore)) = mItem
               
               ReDim Preserve mSubMenus(0 To UBound(mSubMenus) + 1)
               
               mSubMenus(UBound(mSubMenus)) = MII.hSubMenu
            End If
         Next
     
         For mIndex = UBound(mRestore) To 1 Step -1
            RemoveMenu hMenu, vItems(mRestore(mIndex)).ID, 0
         Next
         
         For mIndex = 1 To UBound(mRestore)
            MIS.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            MIS.fState = 0
            MIS.fType = 0
            
            With vItems(mRestore(mIndex))
               If Len(.Cached) Then
                  MIS.dwTypeData = .Cached
               Else
                  MIS.dwTypeData = .Caption
               End If
             
               If ((.Status And lv_mSep) = lv_mSep) Then
                  If MIS.dwTypeData = "" Then
                     MIS.dwItemData = "-"
                  Else
                     .Caption = "LaVolpe" & Chr$(1) & Chr$(2) & Chr$(3) & mIndex
                  End If
               Else
                  If ((.Status And lv_mChk) = lv_mChk) Then MIS.fState = MIS.fState Or MF_CHECKED
                  .Caption = "LaVolpe" & Chr$(1) & Chr$(2) & Chr$(3) & mIndex
               End If
               
               MIS.cch = Len(MIS.dwTypeData)
               
               If ((.Status And lv_mDisabled) = lv_mDisabled) Then
                  MIS.fState = MIS.fState Or MF_DISABLED
                  MIS.fState = MIS.fState Or MF_GRAYED
               End If
               
               MIS.hSubMenu = mSubMenus(mIndex)
               MIS.cbSize = Len(MIS)
               MIS.wID = .ID
               
               InsertMenuItem hMenu, .Index, True, MIS
            End With
         Next
      End If
   Next
   
   Erase mRestore
   Erase mSubMenus
End Sub
   
Public Sub DestroyPopup(hMenu As Long, hMenuID As Long)
   Dim mIndex As Long
   
   On Error Resume Next
   
   mIndex = cItems(hMenuID & "." & hMenu)
   
   ReDim Preserve vItems(0 To UBound(vItems) - 1)
   
   mIndex = cPanels("p" & hMenu)
   gMenus.Remove gMenus.Count
End Sub
   
Private Sub CreateImageSidebar(mIndex As Long, pIndex As Long, wCaption As String)
   Dim bmpInfo As BITMAP
   Dim icoInfo As ICONINFO
   Dim hRect As RECT
   Dim imgSize As POINTAPI
   Dim imgAdj As POINTAPI
   Dim xyOffset As POINTAPI
   Dim tDC As Long
   Dim fDC As Long
   Dim wDC As Long
   Dim ImgType As Integer
   Dim lImgCopy As Long
   Dim bColor As Long
   Dim GColor As Long
   Dim sValue As String
   Dim iAlign As Integer
   Dim imgRatioX As Single
   Dim imgRatioY As Single
   Dim oldBMP(0 To 1) As Long
   Dim newBMP As Long
   Dim srcBMP As Long
   
   On Error Resume Next
   
   srcBMP = ValidateImage(wCaption, ImgType, imgSize.x, imgSize.y)
   
   If srcBMP = 0 Then
      GoTo CreateGenericSidebar
      Exit Sub
   End If
   
   fDC = GetDC(FormHwnd)
   tDC = CreateCompatibleDC(fDC)
   wDC = CreateCompatibleDC(fDC)
   
   With vItems(mIndex)
      ReturnComponentValue wCaption, "Width:", sValue
      
      If Val(sValue) < 16 Then sValue = 16
      .Dimension.x = Val(sValue)
      ReturnComponentValue wCaption, "ALIGN:", sValue
      
      Select Case sValue
         Case "TOP": iAlign = 1
         Case "BOT": iAlign = 2
         Case Else: iAlign = 0
      End Select
      
      If ImgType = 1 Then
         lImgCopy = CopyImage(srcBMP, 0, 0, 0, 0)
         oldBMP(1) = SelectObject(wDC, lImgCopy)
      End If
   
      If imgSize.x > .Dimension.x Or imgSize.y > .Dimension.y Then
         imgRatioX = (.Dimension.x / imgSize.x)
         imgRatioY = (.Dimension.y / imgSize.y)
         
         If imgRatioY < imgRatioX Then imgRatioX = imgRatioY
         imgAdj.x = imgRatioX * imgSize.x
         imgAdj.y = imgRatioX * imgSize.y
      Else
         imgAdj.x = imgSize.x
         imgAdj.y = imgSize.y
      End If
   
      Select Case iAlign
         Case 1: xyOffset.y = 0
         Case 2: xyOffset.y = .Dimension.y - imgAdj.y
         Case Else: xyOffset.y = (.Dimension.y - imgAdj.y) \ 2
      End Select
   
      xyOffset.x = (.Dimension.x - imgAdj.x) \ 2
      newBMP = CreateCompatibleBitmap(fDC, .Dimension.x, .Dimension.y)
      oldBMP(0) = SelectObject(tDC, newBMP)
      ReturnComponentValue wCaption, "BColor:", sValue
      bColor = Val(sValue)
      
      If Val(sValue) = vbNull Then
         bColor = GetSysColor(COLOR_MENU)
      Else
         If bColor <> -1 Then bColor = ExchangeVBcolor(sValue, -1)
      End If
   
      ReturnComponentValue wCaption, "GColor:", sValue
      GColor = ExchangeVBcolor(sValue, vbNull)
      
      If GColor = vbNull Or GColor = bColor Or bColor < 0 Then
         If bColor = -1 Then bColor = GetPixel(wDC, 0, 0)
         DrawRect tDC, 0, 0, .Dimension.x, .Dimension.y, bColor
      Else
         hRect.Right = .Dimension.x
         hRect.Bottom = .Dimension.y
         DrawGradient bColor, GColor, False, tDC, hRect
      End If
   
      If ImgType Then
         ReturnComponentValue wCaption, "Transparent", sValue
         
         If Len(sValue) = 0 Then
            StretchBlt tDC, xyOffset.x, xyOffset.y, imgAdj.x, imgAdj.y, wDC, 0, 0, imgSize.x, imgSize.y, vbSrcCopy
            SelectObject wDC, oldBMP(1)
         Else
            Dim bmpRect As RECT
            
            SelectObject wDC, oldBMP(1)
            hRect.Left = xyOffset.x
            hRect.Right = imgAdj.x + hRect.Left
            hRect.Top = xyOffset.y
            hRect.Bottom = imgAdj.y + hRect.Top
            DrawTransparentBitmap tDC, hRect, lImgCopy, bmpRect, , imgAdj.x, imgAdj.y
         End If
         
         DeleteObject lImgCopy
      Else
         DrawIconEx tDC, xyOffset.x, xyOffset.y, srcBMP, imgAdj.x, imgAdj.y, 0, 0, &H3
      End If
   
      SelectObject tDC, oldBMP(0)
      
      If vPanels(pIndex).PanelIcon Then DeleteObject vPanels(pIndex).PanelIcon
      vPanels(pIndex).PanelIcon = newBMP
   End With

Cleanup:
   DeleteDC tDC
   DeleteDC wDC
   ReleaseDC FormHwnd, fDC
   
   Erase oldBMP
   Exit Sub
   
CreateGenericSidebar:
   CreateTextSidebar mIndex, pIndex, "{Sidebar|Text:La<br0>Volpe Menus|FColor:8388608|BColor:12615680|GColor:16776960|Font:Times New Roman|FSize:14|MinFSize:9|Width:32|Align:Bot}"
End Sub
   
Private Sub CreateTextSidebar(mIndex As Long, pIndex As Long, wCaption As String)
   Dim sMultiLine As String
   Dim sWords() As String
   Dim sCaptionSeg As String
   Dim iFontSize As Single
   Dim MinFontSize As Single
   Dim FontPrefer As Single
   Dim i As Integer
   Dim iLineHT As Integer
   Dim lFontM As LOGFONT
   Dim tRect As RECT
   Dim hRect As RECT
   Dim hPrevFont As Long
   Dim lFont As Long
   Dim maxHT As Long
   Dim fColor As Long
   Dim bColor As Long
   Dim GColor As Long
   Dim sValue As String
   Dim sideBarX As Long
   Dim tDC As Long
   Dim fDC As Long
   Dim oldBMP As Long
   Dim newBMP As Long
   
   fDC = GetDC(FormHwnd)
   tDC = CreateCompatibleDC(fDC)
   ReturnComponentValue wCaption, "Text:", sMultiLine
   
   If Len(sMultiLine) = 0 Then sMultiLine = "La Volpe<br>Menus"
   wCaption = Replace$(wCaption, sMultiLine, "")
   ReturnComponentValue wCaption, "FONT:", sValue
   
   If Len(sValue) = 0 Then sValue = "Arial"
   lFontM.lfCharSet = 1
   lFontM.lfFaceName = sValue & Chr$(0)
   ReturnComponentValue wCaption, "BOLD", sValue
   
   If Len(sValue) Then lFontM.lfWeight = 800 Else lFontM.lfWeight = 400
   ReturnComponentValue wCaption, "ITALIC", sValue
   
   If Len(sValue) Then lFontM.lfItalic = 1
   ReturnComponentValue wCaption, "UNDERLINE", sValue
   
   If Len(sValue) Then lFontM.lfUnderline = 1
   ReturnComponentValue wCaption, "FSize:", sValue
   
   If Len(sValue) Then
      iFontSize = Val(sValue)
      
      If iFontSize < 7 Then iFontSize = 7
      If iFontSize > 24 Then iFontSize = 24
   Else
      iFontSize = 11
   End If
   
   FontPrefer = iFontSize
   ReturnComponentValue wCaption, "MINFSIZE:", sValue
   
   If Len(sValue) Then MinFontSize = Val(sValue)
   If MinFontSize < 5 Then MinFontSize = 5
   If MinFontSize > iFontSize Then MinFontSize = iFontSize
   ReturnComponentValue wCaption, "Width:", sValue
   
   If Val(sValue) < 16 Then sValue = 16
   sideBarX = Val(sValue)
   maxHT = vItems(mIndex).Dimension.y
   ReturnComponentValue wCaption, "FColor:", sValue
   fColor = ExchangeVBcolor(sValue, 0)
   ReturnComponentValue wCaption, "BColor:", sValue
   
   If Val(sValue) = vbNull Then
      bColor = GetSysColor(COLOR_MENU)
   Else
      bColor = ExchangeVBcolor(sValue, GetSysColor(COLOR_MENU))
   End If
   
   ReturnComponentValue wCaption, "GColor:", sValue
   GColor = ExchangeVBcolor(sValue, vbNull)
   
   ReDim sWords(0)
   
   sWords(0) = sMultiLine
   GoSub SetTestFont
   sMultiLine = LineBreaks(sWords(), True, tDC, maxHT)
   
   Do
      DrawText tDC, sMultiLine, Len(sMultiLine), tRect, DT_CALCRECT Or DT_LEFT Or DT_HIDEPREFIX Or DT_MULTILINE Or DT_NOPREFIX
      
      Do While tRect.Right + 6 > maxHT Or tRect.Bottom > sideBarX
         If iFontSize > MinFontSize + 1 Then
            iFontSize = iFontSize - 1
            GoSub SetTestFont
            DrawText tDC, sMultiLine, Len(sMultiLine), tRect, DT_CALCRECT Or DT_LEFT Or DT_HIDEPREFIX Or DT_MULTILINE Or DT_NOPREFIX
         Else
            Exit Do
         End If
      Loop
      
      If tRect.Right + 6 <= maxHT Then Exit Do
      
      If tRect.Bottom < sideBarX Then
         iFontSize = FontPrefer
         GoSub SetTestFont
      End If
      
      sMultiLine = LineBreaks(sWords(), False, tDC, maxHT)
      
      If sMultiLine = "" Then Exit Do
   Loop
   
   lFontM.lfEscapement = 900
   lFontM.lfOrientation = 900
   GoSub SetTestFont
   
   If tRect.Bottom + 6 > sideBarX Then sideBarX = tRect.Bottom + 6
   newBMP = CreateCompatibleBitmap(fDC, sideBarX, maxHT)
   oldBMP = SelectObject(tDC, newBMP)
   vItems(mIndex).Dimension.x = sideBarX
   
   If GColor = vbNull Or GColor = bColor Then
      DrawRect tDC, 0, 0, sideBarX, maxHT, bColor
   Else
      hRect.Right = sideBarX
      hRect.Bottom = maxHT
      DrawGradient bColor, GColor, False, tDC, hRect
   End If
   
   SetBkMode tDC, NEWTRANSPARENT
   SetTextColor tDC, fColor
   DrawText tDC, "X", 1, hRect, DT_CALCRECT Or DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE
   iLineHT = hRect.Right + (hRect.Right \ 2)
   tRect.Left = (sideBarX - tRect.Bottom) \ 2 - 2
   
   For i = 0 To UBound(sWords)
      sCaptionSeg = Replace$(sWords(i), "<br>", " ")
      sCaptionSeg = Replace$(sCaptionSeg, "<br0>", " ")
      hRect.Left = 0: hRect.Top = 0
      DrawText tDC, sCaptionSeg, Len(sCaptionSeg), hRect, DT_CALCRECT Or DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE Or DT_HIDEPREFIX
      ReturnComponentValue wCaption, "ALIGN:", sValue
      
      Select Case sValue
         Case "Bot": tRect.Top = maxHT - 3
         Case "Top": tRect.Top = maxHT - (maxHT - hRect.Right) + 3
         Case Else: tRect.Top = (maxHT - hRect.Right) \ 2 + hRect.Right
      End Select
   
      DrawText tDC, sCaptionSeg, Len(sCaptionSeg), tRect, DT_HIDEPREFIX Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
      tRect.Left = tRect.Left + iLineHT
   Next
   
   SelectObject tDC, hPrevFont
   DeleteObject lFont
   SelectObject tDC, oldBMP
   
   If vPanels(pIndex).PanelIcon Then DeleteObject vPanels(pIndex).PanelIcon
   vPanels(pIndex).PanelIcon = newBMP
   DeleteDC tDC
   ReleaseDC FormHwnd, fDC
   
   Erase sWords
   Exit Sub
   
SetTestFont:
   If hPrevFont Then
      SelectObject tDC, hPrevFont
      DeleteObject lFont
   End If
   
   lFontM.lfHeight = (iFontSize * -20) / Screen.TwipsPerPixelY
   lFont = CreateFontIndirect(lFontM)
   hPrevFont = SelectObject(tDC, lFont)
   Return
End Sub
   
Private Function LineBreaks(Words() As String, bInitial As Boolean, tDC As Long, MaxWidth As Long) As String
   Dim i As Integer
   Dim J As Integer
   Dim K As Integer
   Dim lb As Integer
   Dim sRtn As String
   Dim tRect As RECT
   Dim sSegment As String
   
   i = InStr(Words(0), "<br!>")
   
   Do While i
      ReDim Preserve Words(UBound(Words) + 1)
      
      Words(UBound(Words)) = Mid$(Words(UBound(Words) - 1), i + 5)
      Words(UBound(Words) - 1) = Left$(Words(UBound(Words) - 1), i - 1)
      i = InStr(Words(UBound(Words)), "<br!>")
   Loop
   
   If Not bInitial Then
      For J = 1 To 2
         For i = 0 To UBound(Words)
            lb = InStr(Words(i), Choose(J, "<br>", " "))
            
            If lb Then
               If J = 2 Then
                  DrawText tDC, Words(i), Len(Words(i)), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
                  
                  If tRect.Right + 6 <= MaxWidth Then
                     lb = 0
                  Else
                     K = InStr(Len(Words(i)) \ 2, Words(i), " ")
                     
                     If K = 0 Then K = Len(Words(i))
                     lb = InStrRev(Words(i), " ", Len(Words(i)) \ 2)
                     
                     If lb = 0 Then lb = Len(Words(i))
                     If Abs(Len(Words(i)) \ 2 - lb) > Abs(K - Len(Words(i)) \ 2) Then lb = K
                  End If
               End If
               
               If lb Then
                  ReDim Preserve Words(0 To UBound(Words) + 1)
                 
                  For K = UBound(Words) To i + 2 Step -1
                     Words(K) = Words(K - 1)
                  Next
                 
                  Words(i + 1) = Mid(Words(i), lb + Choose(J, 4, 1))
                  Words(i) = Left(Words(i), lb - 1)
                  Exit For
               End If
            End If
         Next
      
         If lb Then Exit For
      Next
   End If
   
   If lb Or bInitial Then
      For J = 0 To UBound(Words) - 1
         sRtn = sRtn & Words(J) & vbNewLine
      Next
      
      sRtn = sRtn & Words(UBound(Words))
      sRtn = Replace$(sRtn, "<br>", " ")
      LineBreaks = Replace$(sRtn, "<br0>", " ")
      Exit Function
   End If
End Function
   
Private Function ValidateImage(wCaption As String, Optional ImgType As Integer, Optional imgWidth As Long, Optional imgHeight As Long) As Long
   Dim tIcon As Long
   Dim sValue As String
   Dim bmpInfo As BITMAP
   Dim icoInfo As ICONINFO
   Dim frmObject As Control
   
   ReturnComponentValue wCaption, "IMG:", sValue
   
   If Len(sValue) = 0 Then Exit Function
   
   If IsNumeric(sValue) Then
      tIcon = Val(sValue)
   Else
      If IsNumeric(Mid$(sValue, 2)) And Left$(sValue, 1) = "I" Then
         If imgList Is Nothing Then
            tIcon = 0
         Else
            tIcon = Val(Mid$(sValue, 2))
         
            If tIcon <= imgList.ListImages.Count And tIcon > 0 Then
               tIcon = imgList.ListImages(tIcon).Picture.Handle
            Else
               tIcon = 0
            End If
         End If
      Else
         If IsNumeric(Mid$(sValue, 3)) And Left$(sValue, 2) = "&H" Then
            tIcon = Val(sValue)
         Else
            On Error Resume Next
            
            Set frmObject = SplitControlIndex(sValue)
            
            If frmObject Is Nothing Then
               tIcon = 0
            Else
               tIcon = frmObject.Picture.Handle
               
               Set frmObject = Nothing
            End If
         End If
      End If
   End If
   
   If tIcon <> 0 Then
      GetObject tIcon, Len(bmpInfo), bmpInfo
      
      If bmpInfo.bmBits Then
         imgWidth = bmpInfo.bmWidth
         imgHeight = bmpInfo.bmHeight
         ValidateImage = tIcon
         ImgType = 1
      Else
         If GetIconInfo(tIcon, icoInfo) Then
            If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
         
            If icoInfo.hbmColor <> 0 Then
               GetObject icoInfo.hbmColor, Len(bmpInfo), bmpInfo
               imgWidth = bmpInfo.bmWidth
               imgHeight = bmpInfo.bmHeight
               DeleteObject icoInfo.hbmColor
               ValidateImage = tIcon
            End If
         End If
      End If
   End If
End Function
   
Public Function GetShortFileName(ByVal strLongPath As String) As String
   Dim strShortPath As String
   Dim lResult As Long
   Dim nPos As Long
   
   Const cchBuffer = 300
   
   On Error GoTo Function_GetShortPathName_General_ErrTrap_by_LaVolpe
   
   strShortPath = String$(cchBuffer, 0)
   lResult = GetShortPathName(strLongPath, strShortPath, cchBuffer)
   
   If lResult = 0 Then
      GetShortFileName = strLongPath
   Else
      nPos = InStr(strShortPath, vbNullChar)
   
      If nPos > 0 Then
         GetShortFileName = Left$(strShortPath, nPos - 1)
      Else
         GetShortFileName = strShortPath
      End If
   End If
   
   Exit Function
   
Function_GetShortPathName_General_ErrTrap_by_LaVolpe:
   If MsgBox("Error " & Err.Number & " - Procedure [Function GetShortPathName]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function
   
Public Function GetFileDescription(ByVal sFile As String) As String
   Dim lVerSize As Long
   Dim lTemp As Long
   Dim lRet As Long
   Dim bInfo() As Byte
   Dim lpBuffer As Long
   Dim sDesc As String
   Dim sKey As String
   
   Const sEXE As String = "\FileDescription"
   
   lVerSize = GetFileVersionInfoSize(sFile, lTemp)
   
   ReDim bInfo(lVerSize)
   
   If lVerSize > 0 Then
      lRet = GetFileVersionInfo(sFile, lTemp, lVerSize, VarPtr(bInfo(0)))
      
      If lRet <> 0 Then
         sKey = GetNLSKey(bInfo)
         lRet = VerQueryValue(VarPtr(bInfo(0)), sKey & sEXE, lpBuffer, lVerSize)
         
         If lRet <> 0 Then
            sDesc = Space$(lVerSize)
            lstrcpyn sDesc, lpBuffer, lVerSize
            GetFileDescription = sDesc
         End If
      End If
   End If
End Function
   
Private Function GetNLSKey(byteVerData() As Byte) As String
   Const strTRANSLATION$ = "\VarFileInfo\Translation"
   Const strSTRINGFILEINFO$ = "\StringFileInfo\"
   Const strDEFAULTNLSKEY$ = "040904E4"
   Const LOCALE_IDEFAULTLANGUAGE& = &H9&
   Const LOCALE_IDEFAULTCODEPAGE& = &HB&
   
   Static strLANGCP As String
   
   Dim lpBufPtr As Long
   Dim strNLSKey As String
   Dim fGotNLSKey As Integer
   Dim intOffset As Integer
   Dim lVerSize As Long
   Dim lTmp As Long
   Dim lBufLen As Long
   Dim lLCID As Long
   Dim strTmp As String
   
   On Error GoTo GNLSKCleanup
   
   If VerQueryValue(VarPtr(byteVerData(0)), strTRANSLATION, lpBufPtr, lVerSize) <> 0 Then
      If Len(strLANGCP) = 0 Then
         lLCID = GetUserDefaultLCID()
      
         If lLCID > 0 Then
            strTmp = Space$(8)
            GetLocaleInfoA lLCID, LOCALE_IDEFAULTCODEPAGE, strTmp, 8
            strLANGCP = StringFromBuffer(strTmp)
            
            Do While Len(strLANGCP) < 4
               strLANGCP = "0" & strLANGCP
            Loop
            
            GetLocaleInfoA lLCID, LOCALE_IDEFAULTLANGUAGE, strTmp, 8
            strLANGCP = StringFromBuffer(strTmp) & strLANGCP
            
            Do While Len(strLANGCP) < 8
               strLANGCP = "0" & strLANGCP
            Loop
         End If
      End If
   
      If VerQueryValue(VarPtr(byteVerData(0)), strLANGCP, lTmp, lBufLen) <> 0 Then
         strNLSKey = strLANGCP
      Else
         For intOffset = 0 To lVerSize - 1 Step 4
            CopyMemory lTmp, ByVal lpBufPtr + intOffset, 4
            strTmp = Hex$(lTmp)
            
            Do While Len(strTmp) < 8
               strTmp = "0" & strTmp
            Loop
      
            strNLSKey = strSTRINGFILEINFO & Right$(strTmp, 4) & Left$(strTmp, 4)
      
            If VerQueryValue(VarPtr(byteVerData(0)), strNLSKey, lTmp, lBufLen) <> 0 Then
               fGotNLSKey = True
               Exit For
            End If
         Next
   
         If Not fGotNLSKey Then
            strNLSKey = strSTRINGFILEINFO & strDEFAULTNLSKEY
            
            If VerQueryValue(VarPtr(byteVerData(0)), strNLSKey, lTmp, lBufLen) <> 0 Then
               fGotNLSKey = True
            End If
         End If
      End If
   End If
   
GNLSKCleanup:
   If fGotNLSKey Then
      GetNLSKey = strNLSKey
   End If
End Function
   
Private Function GetUNCPath(strDriveLetter As String) As String
   Dim lpszLocalName As String
   Dim lpszRemoteName As String
   Dim cbRemoteName As Long
   
   lpszLocalName = strDriveLetter
   
   If Right$(lpszLocalName, 1) = "\" Then lpszLocalName = Left$(lpszLocalName, Len(lpszLocalName) - 1)
   lpszRemoteName = String$(255, Chr$(32))
   cbRemoteName = Len(lpszRemoteName)
   
   If WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName) Then
      GetUNCPath = "Unknown"
   Else
      Dim sPath As String
      Dim sName As String
      
      sPath = StringFromBuffer(lpszRemoteName)
      
      If Len(sPath) Then
         sName = StripFile(sPath, "N")
         sPath = StripFile(sPath, "P")
         
         If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
         sPath = Replace$(sPath, "\\", "")
         GetUNCPath = " " & sName & " on '" & sPath & "'"
      End If
   End If
End Function
   
Private Function DriveType(strDriveName As String) As String
   Dim lngRet As Long
   Dim strDrive As String
   
   Const DRIVE_UNKNOWN = 0
   Const DRIVE_ABSENT = 1
   Const DRIVE_REMOVABLE = 2
   Const DRIVE_FIXED = 3
   Const DRIVE_REMOTE = 4
   Const DRIVE_CDROM = 5
   Const DRIVE_RAMDISK = 6
   
   lngRet = GetDriveType(strDriveName)
   
   Select Case lngRet
      Case DRIVE_ABSENT: strDrive = "Drive does not exist"
      Case DRIVE_REMOVABLE: strDrive = "Removable Media"
      Case DRIVE_FIXED: strDrive = "Fixed Drive"
      Case DRIVE_REMOTE: strDrive = "Network Drive"
      Case DRIVE_CDROM: strDrive = "CD Rom"
      Case DRIVE_RAMDISK: strDrive = "Ram Disk"
      Case Else: strDrive = "Unknown Drive Type"
   End Select
   
   DriveType = strDrive
End Function
   
Private Sub ListAllDrives(vArray As Variant)
   Dim strAllDrives As String
   Dim strTmp As String
   Dim strRtn As String
   Dim lngRet As Long
   
   ReDim vArray(0 To 0)
   
   strTmp = String$(255, 0)
   lngRet = GetLogicalDriveStrings(255, strTmp)
   strAllDrives = Left(strTmp, lngRet)
   
   If strAllDrives <> "" Then
      Do
         strTmp = StringFromBuffer(strAllDrives)
         strAllDrives = Mid$(strAllDrives, Len(strTmp) + 2)
         strRtn = strTmp
         strTmp = DriveType(strTmp)
         
         If strTmp = "Network Drive" Then
            strTmp = GetUNCPath(strRtn)
            strRtn = strRtn & strTmp & "|Network/Mapped Drive"
         Else
            strRtn = strRtn & "|" & strTmp
         End If
         
         ReDim Preserve vArray(0 To UBound(vArray) + 1)
         
         vArray(UBound(vArray)) = strRtn
         
      Loop While strAllDrives <> ""
   End If
End Sub
