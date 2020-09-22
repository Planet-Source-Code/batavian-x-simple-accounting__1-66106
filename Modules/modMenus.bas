Attribute VB_Name = "modMenus"
Option Explicit
Option Compare Text

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long
Private Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long

Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal byPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As Any) As Long

Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   ItemID As Long
   itemAction As Long
   itemState As Long
   hWndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Public Type ICONINFO
   fIcon As Long
   xHotSpot As Long
   yHotSpot As Long
   hbmMask As Long
   hbmColor As Long
End Type

Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * 32
End Type

Private Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type

Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    ItemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Private Type NONCLIENTMETRICS
   cbSize As Long
   iBorderWidth As Long
   iScrollWidth As Long
   iScrollHeight As Long
   iCaptionWidth As Long
   iCaptionHeight As Long
   lfCaptionFont As LOGFONT
   iSMCaptionWidth As Long
   iSMCaptionHeight As Long
   lfSMCaptionFont As LOGFONT
   iMenuWidth As Long
   iMenuHeight As Long
   lfMenuFont As LOGFONT
   lfStatusFont As LOGFONT
   lfMessageFont As LOGFONT
End Type

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 255
   szTypeName As String * 80
End Type

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type MINMAXINFO
   ptReserved As POINTAPI
   ptMaxSize As POINTAPI
   ptMaxPosition As POINTAPI
   ptMinTrackSize As POINTAPI
   ptMaxTrackSize As POINTAPI
End Type

Public Type MenuComponentData
   Caption As String
   Display As String
   Cached As String
   HotKey As String
   Tip As String
   Dimension As POINTAPI
   OffsetCx As Integer
   ID As Long
   Index As Integer
   Icon As String
   ShowBKG As Boolean
   ControlType As Byte
   hControl As Long
   gControl As Long
   Status As Integer
End Type

Public Type PanelData
   HasIcons As Boolean
   IsSystem As Boolean
   PanelIcon As Long
   SidebarMenuItem As Long
   SubmenuID As Long
   Hourglass As Boolean
   Accelerators As String
End Type

Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNLIGHT = 22
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8
Private Const WHITENESS = &HFF0062

Public Const COLOR_MENU = 4
Public Const NEWTRANSPARENT = 3

Public Const vbMaroon = 128
Public Const vbOlive = 32896
Public Const vbNavy = 8388608
Public Const vbPurple = 8388736
Public Const vbTeal = 8421376
Public Const vbGray = 8421504
Public Const vbSilver = 12632256
Public Const vbViolet = 9445584
Public Const vbOrange = 42495
Public Const vbGold = 43724
Public Const vbIvory = 15794175
Public Const vbPeach = 12180223
Public Const vbTurquoise = 13749760
Public Const vbTan = 9221330
Public Const vbBrown = 17510

Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_HIDEPREFIX As Long = &H100000
Public Const DT_LEFT = &H0
Public Const DT_MULTILINE = (&H1)
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER As Long = &H4
Public Const DT_WORDBREAK = &H10

Private Const WM_USER As Long = &H400
Private Const GW_CHILD = 5
Private Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = -20
Private Const TBN_FIRST = (-700&)
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TBSTYLE_CUSTOMERASE = &H2000
Private Const TB_GETRECT = (WM_USER + 51)
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B
Private Const WM_ENTERIDLE = &H121
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_GETMINMAXINFO As Long = &H24&
Private Const WA_INACTIVE As Long = 0
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIDESTROY As Long = &H221
Private Const WM_MDIGETACTIVE As Long = &H229
Private Const WM_MDIMAXIMIZE As Long = &H225
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUCHAR = &H120
Private Const WM_MENUCOMMAND As Long = &H126
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_SETFOCUS As Long = &H7
Private Const WS_EX_MDICHILD = &H40&
Private Const WM_KEYUP As Long = &H101
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105

Public Const GWL_STYLE = -16
Public Const GWL_ID = -12
Public Const WM_COMMAND As Long = &H111

Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_HILITE = &H80&
Private Const MF_MOUSESELECT = &H8000&
Private Const MNC_EXECUTE = 2
Private Const MNC_IGNORE = 0
Private Const MNC_SELECT = 3
Private Const ODA_DRAWENTIRE As Long = &H1
Private Const ODT_MENU = 1
Private Const ODS_SELECTED = &H1
Private Const TPM_RETURNCMD As Long = &H100&
Private Const TPM_NONOTIFY As Long = &H80&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_VERTICAL = &H40&
Private Const TPM_LEFTBUTTON = &H0&

Public Const MF_CHANGE As Long = &H80&
Public Const MF_CHECKED As Long = &H8&
Public Const MF_DEFAULT As Long = &H1000&
Public Const MF_DISABLED As Long = &H2&
Public Const MF_GRAYED As Long = &H1&
Public Const MF_MENUBARBREAK As Long = &H20&
Public Const MF_MENUBREAK As Long = &H40&
Public Const MF_OWNERDRAW = &H100
Public Const MF_POPUP As Long = &H10&
Public Const MF_SEPARATOR = &H800
Public Const MIIM_DATA = &H20
Public Const MIIM_ID As Long = &H2
Public Const MIIM_STATE As Long = &H1
Public Const MIIM_SUBMENU As Long = &H4
Public Const MIIM_TYPE = &H10

Public Const SC_CLOSE = &HF060
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030
Public Const SC_RESTORE = &HF120

Private Const SPI_GETWORKAREA = 48
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_MENU As Long = &H12

Public Const RASTER_FONTTYPE As Long = &H1
Public Const TRUETYPE_FONTTYPE As Long = &H4

Public Const lv_mSep As Integer = 2
Public Const lv_mDisabled As Integer = 4
Public Const lv_mChk As Integer = 8
Public Const lv_mDefault As Integer = 16
Public Const lv_mSepRaised As Integer = 32
Public Const lv_mSBar As Integer = 64
Public Const lv_mSBarHidden As Integer = 128
Public Const lv_mSubmenu As Integer = 512
Public Const lv_mCustom As Integer = 1024
Public Const lv_mColor As Integer = 2048
Public Const lv_mFont As Integer = 4096

Public Enum MenuImageType
   lv_ImgListIndex = 0
   lv_ImgHandle = 1
   lv_ImgControl = 2
End Enum

Public Enum MenuCtrlType
   lv_ListBox = 1
   lv_ComboBox = 2
End Enum

Public Enum MenuCaptionProps
   lv_Caption = 0
   lv_ImgID = 1
   lv_Bold = 2
   lv_Tip = 3
   lv_ListBoxID = 4
   lv_ComboxID = 5
   lv_ShowIconBkg = 6
   lv_HotKey = 7
   lv_FilesPath = 8
End Enum

Public Enum SidebarTextProps
   lv_txtText = 1
   lv_txtForeColor = 2
   lv_txtBackColor = 3
   lv_txtGradientColor = 4
   lv_txtFontName = 5
   lv_txtFontSize = 6
   lv_txtMinFontSize = 7
   lv_txtWidth = 8
   lv_txtAlignment = 9
   lv_txtTip = 10
   lv_txtBold = 11
   lv_txtItalic = 12
   lv_txtUnderline = 13
   lv_txtNoScroll = 14
   lv_txtDisabled = 15
End Enum

Public Enum SidebarImgProps
   lv_imgImgID = 1
   lv_imgBackColor = 2
   lv_imgGradientColor = 3
   lv_imgWidth = 4
   lv_imgAlignment = 5
   lv_imgTip = 6
   lv_imgNoScroll = 7
   lv_imgTransparent = 8
   lv_imgDisabled = 9
End Enum

Public Enum MenuSepProps
   lv_sCaption = 0
   lv_sRaisedEffect = 1
End Enum

Public Enum FontTypeEnum
   lv_fAllFonts = 0
   lv_fTrueType = 1
   lv_fNonTrueType = 2
End Enum

Public Enum AlignmentEnum
   lv_TopOfMenu = 1
   lv_BottomOfMenu = 2
   lv_CenterOfMenu = 0
End Enum

Public Enum CstmMonth
   lv_cDefault = 0
   lv_cCalendarQuarter = 1
   lv_cFiscalQuarter = 2
End Enum

Public Enum ColorObjects
   cObj_Brush = 0
   cObj_Pen = 1
   cObj_Text = 2
End Enum

Public Enum SubClassContainers
   lv_NonMDIform = 0
   lv_MDIform_ChildrenHaveMenus = 0
   lv_MDIform_ChildrenMenuless = 5
   lv_VB_Toolbar = 1
   lv_MDIchildForm_NoMenus = 4
   lv_MDIchildForm_WithMenus = 3
End Enum

Private ExtraOffsetX As Integer
Private m_Font(0 To 7)
Private colMenuItems As Collection
Private OpenMenus As Collection
Private vFonts() As String
Private hWndRedirect As String
Private tempRedirect As Long
Private bKeyBoardSelect As Boolean

Public AmInIDE As Boolean
Public bAmDebugging As Boolean
Public DefaultIcon As Long
Public XferMenuData As MenuComponentData
Public XferPanelData As PanelData

Private bHiLiteDisabled As Boolean
Private bItalicSelected As Boolean
Private bGradientSelect As Boolean
Private mFontName As String
Private mFontSize As Single
Private vMenuListBox As Long
Private bReturnMDIkeystrokes As Boolean
Private bRaisedIcons As Boolean
Private bXPcheckmarks As Boolean

Private bModuleInitialized As Boolean
Private lSelectBColor As Long
Private TextColorNormal As Long
Private TextColorSelected As Long
Private TextColorDisabledDark As Long
Private TextColorDisabledLight As Long
Private TextColorSeparatorBar As Long
Private SeparatorBarColorDark As Long
Private SeparatorBarColorLight As Long
Private CheckedIconBColor As Long

Private FloppyIcon As Long
Private tbarClass() As String
Private bUseHourglass As Boolean

Public lFlag As Long

Public GradCol1 As Long
Public GradCol2 As Long
Public GradBackCol As Long

Public SelFrameCol As Long

Public SelGradCol As Long
Public SelGradBackCol As Long

Public DisSelGradCol As Long
Public DisSelGradBackCol As Long

Public SepCol As Long

Public Property Let MenuFontName(sFontName As String)
   If Len(sFontName) > 0 And sFontName <> mFontName Then
      mFontName = sFontName
      CreateDestroyMenuFont False, False
      CreateDestroyMenuFont True, False
      
      If bItalicSelected Then CreateDestroyMenuFont True, True
      
      If Not colMenuItems Is Nothing Then
         Dim I As Integer
         
         For I = colMenuItems.Count To 1 Step -1
            colMenuItems(I).RestoreMenus
         Next
      End If
   End If
End Property

Public Property Get MenuFontName() As String
   If mFontName = "" Then
       Dim ncm As NONCLIENTMETRICS
       Dim I As Integer

       ncm.cbSize = Len(ncm)
       SystemParametersInfo 41, 0, ncm, 0
       I = InStr(ncm.lfMenuFont.lfFaceName, Chr$(0))
       If I = 0 Then I = Len(ncm.lfMenuFont.lfFaceName) + 1
       mFontName = Left$(ncm.lfMenuFont.lfFaceName, I - 1)
       If mFontSize = 0 Then mFontSize = Abs(ncm.lfMenuFont.lfHeight) * 0.72
   End If

   MenuFontName = mFontName
End Property

Public Property Let MenuFontSize(NewSize As Single)
   If NewSize <> mFontSize And NewSize > 0 Then
      mFontSize = NewSize
      CreateDestroyMenuFont False, False
      CreateDestroyMenuFont True, False
      If bItalicSelected Then CreateDestroyMenuFont True, True

      If Not colMenuItems Is Nothing Then
         Dim I As Integer
         For I = colMenuItems.Count To 1 Step -1
            colMenuItems(I).RestoreMenus
         Next
      End If
   End If
End Property

Public Property Get MenuFontSize() As Single
   MenuFontSize = mFontSize
End Property

Public Property Let MenuCaptionListBox(hwnd As Long)
   If vMenuListBox Then
      If IsWindow(vMenuListBox) = 0 Then vMenuListBox = 0
   End If

   If Not vMenuListBox Then vMenuListBox = hwnd
End Property

Public Property Get MenuCaptionListBox() As Long
   MenuCaptionListBox = vMenuListBox
End Property

Public Property Let HighlightGradient(bGradient As Boolean)
   bGradientSelect = bGradient
End Property

Public Property Get HighlightGradient() As Boolean
   HighlightGradient = bGradientSelect
End Property

Public Property Get RaisedIconOnSelect() As Boolean
   RaisedIconOnSelect = bRaisedIcons
End Property

Public Property Let RaisedIconOnSelect(bYesNo As Boolean)
   bRaisedIcons = bYesNo
End Property

Public Property Get CheckMarksXPstyle() As Boolean
   CheckMarksXPstyle = bXPcheckmarks
End Property

Public Property Let CheckMarksXPstyle(bYesNo As Boolean)
   bXPcheckmarks = bYesNo
End Property

Public Property Get SelectedItemBackColor() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   SelectedItemBackColor = lSelectBColor
End Property

Public Property Let SelectedItemBackColor(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   lSelectBColor = ConvertColor(lColor)
End Property

Public Property Get SelectedItemTextColor() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   SelectedItemTextColor = TextColorSelected
End Property

Public Property Let SelectedItemTextColor(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   TextColorSelected = ConvertColor(lColor)
End Property

Public Property Get MenuItemTextColor() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   MenuItemTextColor = TextColorNormal
End Property

Public Property Let MenuItemTextColor(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   TextColorNormal = ConvertColor(lColor)
End Property

Public Property Get DisabledTextColor_Dark() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   DisabledTextColor_Dark = TextColorDisabledDark
End Property

Public Property Let DisabledTextColor_Dark(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   TextColorDisabledDark = ConvertColor(lColor)
End Property

Public Property Get DisabledTextColor_Light() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   DisabledTextColor_Light = TextColorDisabledLight
End Property

Public Property Let DisabledTextColor_Light(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   TextColorDisabledLight = ConvertColor(lColor)
End Property

Public Property Get SeparatorBarTextColor() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   SeparatorBarTextColor = TextColorSeparatorBar
End Property

Public Property Let SeparatorBarTextColor(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   TextColorSeparatorBar = ConvertColor(lColor)
End Property

Public Property Get SeparatorBarColor_Dark() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   SeparatorBarColor_Dark = SeparatorBarColorDark
End Property

Public Property Let SeparatorBarColor_Dark(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   SeparatorBarColorDark = ConvertColor(lColor)
End Property

Public Property Get SeparatorBarColor_Light() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   SeparatorBarColor_Light = SeparatorBarColorLight
End Property

Public Property Let SeparatorBarColor_Light(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   SeparatorBarColorLight = ConvertColor(lColor)
End Property

Public Property Get CheckedIconBackColor() As Long
   If Not bModuleInitialized Then LoadDefaultColors
   CheckedIconBackColor = CheckedIconBColor
End Property

Public Property Let CheckedIconBackColor(lColor As Long)
   If Not bModuleInitialized Then LoadDefaultColors
   CheckedIconBColor = ConvertColor(lColor)
End Property

Public Property Let ReturnMDIkeystrokes(bYesNo As Boolean)
   bReturnMDIkeystrokes = bYesNo
End Property

Public Property Get ReturnMDIkeystrokes() As Boolean
   ReturnMDIkeystrokes = bReturnMDIkeystrokes
End Property

Public Property Let HighlightDisabledMenuItems(bHiLite As Boolean)
   bHiLiteDisabled = bHiLite
End Property

Public Property Get HighlightDisabledMenuItems() As Boolean
   HighlightDisabledMenuItems = bHiLiteDisabled
End Property

Public Property Let ItalicizeSelectedItems(bItalics As Boolean)
   If bItalics = bItalicSelected Then Exit Property
   bItalicSelected = bItalics
   If bItalics Then CreateDestroyMenuFont True, True
End Property

Public Property Get ItalicizeSelectedItems() As Boolean
   ItalicizeSelectedItems = bItalicSelected
End Property

Public Property Get Win98MEoffset() As Integer
   Win98MEoffset = ExtraOffsetX
End Property

Public Function CreateTextSidebar(Caption As String, FontName As String, FontSize As Single, _
   Optional MinFontSize As Single = 9, Optional Bold As Boolean = False, Optional Underline As Boolean = False, _
   Optional Italic As Boolean = False, Optional ForeColor As Long, Optional BackColor As Long = -1, Optional Gradient2ndColor As Long = vbNull, _
   Optional Width As Integer = 32, Optional Alignment As AlignmentEnum = lv_BottomOfMenu, _
   Optional NoShowIfScrolls As Boolean = False, Optional Tip As String, Optional AlwaysDisabled As Boolean) As String

   Dim wCaption As String
   Dim sValue As String

   If Caption = "" Then Caption = " "
   If FontName = "" Then FontName = "Arial"

   Caption = "{Sidebar|Text:" & Caption & "|Font:" & FontName

   If FontSize < 9 Then wCaption = "|FSize:9" Else wCaption = "|FSize:" & FontSize
   If MinFontSize Then wCaption = wCaption & "|MinFSize:" & MinFontSize
   If Bold Then wCaption = wCaption & "|Bold"
   If Underline Then wCaption = wCaption & "|Underline"
   If Italic Then wCaption = wCaption & "|Italic"
   If AlwaysDisabled Then wCaption = wCaption & "|SBDisabled"

   wCaption = wCaption & "|FColor:" & ForeColor
   wCaption = wCaption & "|BColor:" & BackColor

   If Gradient2ndColor <> vbNull Then wCaption = wCaption & "|GColor:" & Gradient2ndColor
   If Width < 16 Then Width = 16

   wCaption = wCaption & "|Width:" & Width

   Select Case Alignment
      Case lv_TopOfMenu: wCaption = wCaption & "|Align:Top"
      Case lv_BottomOfMenu: wCaption = wCaption & "|Align:Bot"
      Case lv_CenterOfMenu: wCaption = wCaption & "|Align:Ctr"
   End Select

   If NoShowIfScrolls Then wCaption = wCaption & "|NoScroll"
   If Len(Tip) Then Caption = Caption & "|Tip:" & Tip

   CreateTextSidebar = Caption & wCaption & "}"
End Function

Public Function ChangeTextSidebar(CaptionNow As String, Property As SidebarTextProps, Optional newValue As Variant) As String
   Dim sNewCaption As String
   Dim sValue As String
   Dim wCaption As String
   Dim sProp As String

   ChangeTextSidebar = CaptionNow
   If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)

   Dim sParts(1 To 15) As String
   Dim sTarget As String
   Dim I As Integer

   For I = 1 To UBound(sParts)
      sTarget = Choose(I, "Text:", "FColor:", "BColor:", "GColor:", "Font:", "FSize:", "MinFSize:", "Width:", "Align:", "Tip", "Bold", "Italic", "Underline", "NoScroll", "SBDisabled")
      ReturnComponentValue CaptionNow, sTarget, sValue

      If Len(sValue) Then
         sParts(I) = sTarget
         If I < 11 Then sParts(I) = sParts(I) & sValue
         sParts(I) = sParts(I) & "|"
      Else
         sParts(I) = ""
      End If
   Next

   Select Case Property
      Case lv_txtText
         sTarget = "Text:"
         If Len(sProp) = 0 Then sProp = " "

      Case lv_txtForeColor
         sTarget = "FColor:"
         sProp = CStr(Val(sProp))

      Case lv_txtBackColor
         If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
         sTarget = "BColor:"

      Case lv_txtGradientColor
         If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
         sTarget = "GColor:"

      Case lv_txtFontName
         If Len(sProp) = 0 Then sProp = "Tahoma"
         sTarget = "Font:"

      Case lv_txtFontSize
         If Len(sProp) = 0 Then sProp = 9 Else sProp = CStr(Val(sProp))
         sTarget = "FSize:"

      Case lv_txtMinFontSize
         If Len(sProp) = 0 Then
            sTarget = ""
            sParts(Property) = ""
         Else
            sProp = CStr(Val(sProp))
            sTarget = "MinFSize:"
         End If

      Case lv_txtBold, lv_txtItalic, lv_txtUnderline, lv_txtNoScroll, lv_txtDisabled
         If Len(sProp) = 0 Then sProp = "False"

         If CBool(sProp) = False Then
            sParts(Property) = ""
         Else
            sParts(Property) = Choose(Property - 10, "Bold|", "Italic|", "Underline|", "NoScroll|", "SBDisabled|")
         End If

         sTarget = ""

      Case lv_txtTip
         sTarget = "Tip:"
         If Len(sProp) = 0 Then sTarget = "": sParts(Property) = ""

      Case lv_txtWidth
         If Val(sProp) < 35 Then sProp = 35
         sTarget = "Width:"

      Case lv_txtAlignment
         sTarget = "Align:"

         Select Case sProp
            Case "1", "Top": sProp = "Top"
            Case "2", "Bot": sProp = "Bot"
            Case Else:
               sTarget = ""
               sParts(Property) = ""
         End Select

      Case Else
         ChangeTextSidebar = CaptionNow
         Exit Function

   End Select

   If Len(sTarget) Then
      If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
   End If

   wCaption = ""

   For I = 1 To UBound(sParts)
      wCaption = wCaption & sParts(I)
   Next

   If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"

   ChangeTextSidebar = "{Sidebar|" & wCaption
   Erase sParts
End Function

Public Function CreateImageSidebar(ImgType As MenuImageType, ImgID As String, Optional TRANSPARENT As Boolean = False, _
   Optional BackColor As Long = -1, Optional Gradient2ndColor As Long = vbNull, Optional Width As Integer = 32, _
   Optional Alignment As AlignmentEnum, Optional NoShowIfScrolls As Boolean = False, Optional Tip As String) As String

   Dim Caption As String
   Dim sValue As String

   If Len(ImgID) = 0 Then Exit Function

   If (ImgType > lv_ImgListIndex - 1 And ImgType < lv_ImgControl + 1) Then
      If ImgType = lv_ImgListIndex Then sValue = "IMG:i" & ImgID Else sValue = "IMG:" & ImgID
      Caption = sValue
   Else
      Exit Function
   End If

   If TRANSPARENT Then Caption = Caption & "|Transparent"

   Caption = Caption & "|BColor:" & BackColor

   If Gradient2ndColor <> vbNull Then Caption = Caption & "|GColor:" & Gradient2ndColor
   If Width < 32 Then Width = 32

   Caption = Caption & "|Width:" & Width

   Select Case Alignment
      Case 1: Caption = Caption & "|Align:Top"
      Case 2: Caption = Caption & "|Align:Bot"
   End Select

   If NoShowIfScrolls Then Caption = Caption & "|NoScroll"
   If Len(Tip) Then Caption = Caption & "|Tip:" & Tip

   CreateImageSidebar = "{Sidebar|" & Caption & "}"
End Function

Public Function ChangeImageSidebar(CaptionNow As String, Property As SidebarImgProps, Optional newValue As Variant) As String
   Dim sNewCaption As String
   Dim sValue As String
   Dim wCaption As String
   Dim sProp As String

   ChangeImageSidebar = CaptionNow

   If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)

   Dim sParts(0 To 9) As String
   Dim sTarget As String
   Dim I As Integer

   For I = 1 To UBound(sParts)
      sTarget = Choose(I, "IMG:", "BColor", "GColor", "Width:", "Align:", "Tip:", "NoScroll", "Transparent", "SBDisabled")
      ReturnComponentValue CaptionNow, sTarget, sValue

      If Len(sValue) Then
         sParts(I) = sTarget
         If I < 7 Then sParts(I) = sParts(I) & sValue
         sParts(I) = sParts(I) & "|"
      Else
         sParts(I) = ""
      End If
   Next

   Select Case Property
      Case lv_imgImgID
         sTarget = "IMG:"

         If Len(sProp) = 0 Then
            ChangeImageSidebar = ""
            Exit Function
         End If

      Case lv_imgBackColor
         If Len(sProp) = 0 Then sProp = vbButtonFace Else sProp = CStr(Val(sProp))
         sTarget = "BColor:"

      Case lv_imgGradientColor
         If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
         sTarget = "GColor:"

      Case lv_imgAlignment
         sTarget = "Align:"

         Select Case sProp
            Case "1", "Top": sProp = "Top"
            Case "2", "Bot": sProp = "Bot"
            Case Else:
                sTarget = ""
                sParts(Property) = ""
         End Select

      Case lv_imgTip
         sTarget = "Tip:"
         If Len(sProp) = 0 Then sParts(Property) = "": sTarget = ""

      Case lv_imgWidth
         If Val(sProp) < 32 Then sProp = 32 Else sProp = CStr(Val(sProp))
         sTarget = "Width:"

      Case lv_imgNoScroll, lv_imgTransparent, lv_imgDisabled
         If Len(sProp) = 0 Then sProp = "False"

         If CBool(sProp) Then
            sParts(Property) = Choose(Property - 6, "NoScroll|", "Transparent|")
         Else
            sParts(Property) = ""
         End If

         sTarget = ""

   End Select

   If Len(sTarget) Then
      If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
   End If

   wCaption = ""

   For I = 1 To UBound(sParts)
      wCaption = wCaption & sParts(I)
   Next

   If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"

   ChangeImageSidebar = "{Sidebar|" & sNewCaption & wCaption
   Erase sParts
End Function

Public Function CreateSepartorBar(Optional Caption As String, Optional RaisedEffect As Boolean) As String
   Dim newCaption As String

   If Len(Caption) = 0 Then newCaption = "-" Else newCaption = Caption
   If Left(newCaption, 1) <> "-" Then newCaption = "-" & newCaption
   If RaisedEffect Then newCaption = newCaption & "{Raised}"

   CreateSepartorBar = newCaption
End Function

Public Function ChangeSepartorBar(CaptionNow As String, Property As MenuSepProps, Optional newValue As Variant)
   Dim sNewCaption As String
   Dim sValue As String
   Dim wCaption As String
   Dim sProp As String

   SeparateCaption CaptionNow, sNewCaption, wCaption

   If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)

   Select Case Property
      Case lv_sCaption
         If Len(sProp) = 0 Then sProp = "-"
         If Left(sProp, 1) <> "-" Then sProp = "-" & sProp
         ReturnComponentValue wCaption, "Raised", sValue
         If Len(sValue) Then sValue = "{Raised}"
         sNewCaption = sProp & sValue
      Case lv_sRaisedEffect
         If sProp = "" Then sProp = "False"
         If CBool(sProp) Then sProp = "{Raised}" Else sProp = ""
         sNewCaption = sNewCaption & sProp
   End Select

   ChangeSepartorBar = sNewCaption
End Function

Public Function CreateMenuCaption(Caption As String, Optional ImgType As MenuImageType, _
   Optional ImgID As String, Optional NoTransparency As Boolean, Optional HotKey As String, _
   Optional BoldText As Boolean = False, Optional Tip As String, _
   Optional ListComboType As MenuCtrlType, Optional ListComboID As String, _
   Optional ListsFiles As Boolean, Optional FilesPath As String) As String

   Dim wCaption As String
   Dim sValue As String

   If (ImgType > lv_ImgListIndex - 1 And ImgType < lv_ImgControl + 1) And Len(ImgID) > 0 Then
      If ImgType = lv_ImgListIndex Then sValue = "IMG:i" & ImgID Else sValue = "IMG:" & ImgID
      wCaption = sValue
   End If

   If NoTransparency Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      wCaption = wCaption & sValue & "ImgBkg"
   End If

   If Len(HotKey) Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      wCaption = wCaption & sValue & "HotKey:" & HotKey
   End If

   If BoldText Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      wCaption = wCaption & sValue & "Default"
   End If

   If ListsFiles Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      wCaption = wCaption & "|Files:"
      If Len(FilesPath) Then wCaption = wCaption & FilesPath Else wCaption = wCaption & "-1"
   End If

   If Len(Tip) Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      wCaption = wCaption & sValue & "Tip:" & Tip
   End If

   If (ListComboType = lv_ComboBox Or ListComboType = lv_ListBox) And Len(ListComboID) > 0 Then
      If Len(wCaption) Then sValue = "|" Else sValue = ""
      If ListComboType = lv_ComboBox Then sValue = sValue & "CB:" Else sValue = sValue & "LB:"

      wCaption = wCaption & sValue & ListComboID
   End If

   If Len(wCaption) Then wCaption = "{" & wCaption & "}"
   CreateMenuCaption = Caption & wCaption
End Function

Public Function ChangeMenuCaption(CaptionNow As String, Property As MenuCaptionProps, Optional newValue As Variant) As String
   Dim sNewCaption As String
   Dim sValue As String
   Dim wCaption As String
   Dim sProp As String
   Dim sParts(0 To 8) As String
   Dim sTarget As String
   Dim I As Integer

   If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)

   ChangeMenuCaption = CaptionNow
   sNewCaption = CaptionNow
   SeparateCaption CaptionNow, sNewCaption, wCaption

   For I = 1 To UBound(sParts)
      sTarget = Choose(I, "IMG:", "Default", "Tip:", "LB:", "CB:", "IMGBKG", "HotKey:", "Files:")
      ReturnComponentValue wCaption, sTarget, sParts(I)

      If Len(sParts(I)) Then sParts(I) = sTarget & sParts(I) & "|"
   Next

   Select Case Property
      Case lv_Caption
         sTarget = ""
         sNewCaption = newValue

      Case lv_ImgID: sTarget = "IMG:"

      Case lv_Bold
         sTarget = ""
         If Len(sProp) = 0 Then sProp = "False"
         If CBool(sProp) Then sParts(Property) = "Default|" Else sParts(Property) = ""

      Case lv_Tip: sTarget = "Tip:"

      Case lv_ListBoxID
         sParts(Property + 1) = ""
         sTarget = "LB:"

      Case lv_ComboxID
         sParts(Property - 1) = ""
         sTarget = "CB:"

      Case lv_ShowIconBkg
         sTarget = ""

         If Len(sProp) = 0 Then sProp = "False"
         If CBool(sProp) Then sParts(Property) = "ImgBkg|" Else sParts(Property) = ""

      Case lv_HotKey: sTarget = "HotKey:"

      Case lv_FilesPath
         If sProp = "" Then
            sParts(Property) = ""
            sTarget = ""
         Else
            sParts(Property) = sProp
            sTarget = "Files:"
         End If

      Case Else
         If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"
         ChangeMenuCaption = sNewCaption & wCaption
         Exit Function

   End Select

   If Len(sTarget) Then
      If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
   End If

   wCaption = ""

   For I = 1 To UBound(sParts)
      wCaption = wCaption & sParts(I)
   Next

   If Len(wCaption) Then wCaption = "{" & Left$(wCaption, Len(wCaption) - 1) & "}"

   ChangeMenuCaption = sNewCaption & wCaption
   Erase sParts
End Function

Public Function CreateLvColors(CaptionNow As String, Optional UserID As String, Optional CheckedColor As Long = -1) As String
   CreateLvColors = BuildSimpleCustomMenu(CaptionNow, "lvColors:", UserID, CStr(CheckedColor))
End Function

Public Function CreateLvDrives(CaptionNow As String, Optional UserID As String, Optional CheckedDrive As String = "-1") As String
   CreateLvDrives = BuildSimpleCustomMenu(CaptionNow, "lvDrives:", UserID, CStr(CheckedDrive))
End Function

Public Function CreateLvDaysOfWeek(CaptionNow As String, Optional UserID As String, Optional CheckedDay As Integer = -1) As String
   CreateLvDaysOfWeek = BuildSimpleCustomMenu(CaptionNow, "lvDays:", UserID, CStr(CheckedDay))
End Function

Public Function CreateLvStates(CaptionNow As String, Optional UserID As String, Optional CheckedState As String = "-1") As String
   CreateLvStates = BuildSimpleCustomMenu(CaptionNow, "lvStates:", UserID, CheckedState)
End Function

Public Function CreateLvMonths(CaptionNow As String, Optional UserID As String, Optional CheckedMonth As Integer = -1, Optional Grouping As CstmMonth) As String
   Dim newCaption As String
   Dim wCaption As String
   Dim sValue As String
   Dim sCode As String
   Dim ArrayIndex As Integer
   Dim I As Integer

   sCode = "lvMonths:" & CheckedMonth

   Select Case Grouping
      Case lv_cCalendarQuarter: sCode = sCode & ":Group:CYQtr"
      Case lv_cFiscalQuarter: sCode = sCode & ":Group:FYQtr"
      Case Else: sCode = sCode & ":Group:Default"
   End Select

   If Len(UserID) Then sCode = sCode & ":ID:" & UserID

   SeparateCaption CaptionNow, newCaption, wCaption
   ReturnComponentValue wCaption, "lvMonths:", sValue

   If Len(sValue) Then
      wCaption = Trim(Replace$(wCaption, "lvMonths:" & sValue & "|", ""))
      wCaption = Trim(Replace$(wCaption, "lvMonths:" & sValue, ""))
   End If

   If Len(wCaption) Then
      wCaption = Replace$(wCaption, "{", "")
      wCaption = Replace$(wCaption, "}", "")

      If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption

      wCaption = "{" & sCode & wCaption & "}"
   Else
      wCaption = "{" & sCode & "}"
   End If

   CreateLvMonths = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Function CreateLvDaysOfMonth(CaptionNow As String, Optional UserID As String, _
   Optional Year As Integer = 0, Optional Month As Integer = 0, _
   Optional CheckedDate As Integer = -1) As String

   Dim newCaption As String
   Dim wCaption As String
   Dim sValue As String
   Dim sCode As String

   sCode = "lvMonth:" & Month
   sCode = sCode & ":Year:" & Year
   sCode = sCode & ":Day:" & CheckedDate

   If Len(UserID) Then sCode = sCode & ":ID:" & UserID

   SeparateCaption CaptionNow, newCaption, wCaption
   ReturnComponentValue wCaption, "lvMonth:", sValue

   If Len(sValue) Then
      wCaption = Trim(Replace$(wCaption, "lvMonth:" & sValue & "|", ""))
      wCaption = Trim(Replace$(wCaption, "lvMonth:" & sValue, ""))
   End If

   If Len(wCaption) Then
      wCaption = Replace$(wCaption, "{", "")
      wCaption = Replace$(wCaption, "}", "")
      If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
      wCaption = "{" & sCode & wCaption & "}"
   Else
      wCaption = "{" & sCode & "}"
   End If

   CreateLvDaysOfMonth = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Function CreateLvFonts(CaptionNow As String, Optional UserID As String, _
   Optional CheckedFont As String = "-1", Optional FontType As FontTypeEnum, _
   Optional FilterLetterA As String, Optional FilterLetterZ As String) As String

   Dim newCaption As String
   Dim wCaption As String
   Dim sValue As String
   Dim sCode As String

   sCode = "lvFonts:" & CheckedFont
   sCode = sCode & ":Type:" & Choose(FontType + 1, "All", "TrueType", "System")

   If FilterLetterA <> "" And FilterLetterZ <> "" Then
      sCode = sCode & ":Group:" & FilterLetterA & "-" & FilterLetterZ
   End If

   If Len(UserID) Then sCode = sCode & ":ID:" & UserID

   SeparateCaption CaptionNow, newCaption, wCaption
   ReturnComponentValue wCaption, "lvFonts:", sValue

   If Len(sValue) Then
      wCaption = Trim(Replace$(wCaption, "lvFonts:" & sValue & "|", ""))
      wCaption = Trim(Replace$(wCaption, "lvFonts:" & sValue, ""))
   End If

   If Len(wCaption) Then
      wCaption = Replace$(wCaption, "{", "")
      wCaption = Replace$(wCaption, "}", "")
      If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
      wCaption = "{" & sCode & wCaption & "}"
   Else
      wCaption = "{" & sCode & "}"
   End If

   CreateLvFonts = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Function ChangeCustomMenu(CaptionNow As String, Optional NewUserID As String, _
   Optional NewCheckedItem As Variant, Optional NewGrouping As String, Optional NewFontType As FontTypeEnum = -1) As String

   Dim sValue As String
   Dim wCaption As String
   Dim bRecognized As Boolean
   Dim oldCaption As String
   Dim oldWcaption As String
   Dim chkType As String

   SeparateCaption CaptionNow, oldCaption, wCaption

   Dim sType As String
   Dim I As Integer
   Dim J As Integer

   For I = 1 To 7
      sType = Choose(I, "lvFonts:", "lvColors:", "lvMonths:", "lvMonth:", "lvDrives:", "lvDays:", "lvStates:")

      If InStr(wCaption, sType) Then
         I = InStr(wCaption, sType)
         J = InStr(I, wCaption, "|")
         oldWcaption = Replace$(wCaption, Mid$(wCaption, I, J - I), "")
         wCaption = Mid$(wCaption, I, J - 1)
         oldWcaption = Replace$(oldWcaption, "{", "")

         If Right$(oldWcaption, 1) = "|" Then oldWcaption = Left$(oldWcaption, Len(oldWcaption) - 1)
         If Right$(wCaption, 1) <> "|" Then wCaption = wCaption & "|"

         wCaption = "{" & wCaption
         bRecognized = True
         Exit For
      End If
   Next

   If Not bRecognized Then Exit Function

   If Not IsMissing(NewCheckedItem) Then
      If Len(CStr(NewCheckedItem)) Then
         If sType = "lvMonth:" Then chkType = "Day:" Else chkType = sType
         ReturnComponentValue wCaption, chkType, sValue
         J = InStr(sValue, ":")

         If J Then
            If Mid$(sValue, J + 1, 1) = "\" And sType = "lvDrives:" Then
               J = InStr(J + 1, sValue, ":")
            End If
         End If

         If J = 0 Then
            J = InStr(sValue, "|")
            If J = 0 Then J = Len(sValue) + 1
         End If

         wCaption = Replace$(wCaption, chkType & Left$(sValue, J - 1), chkType & CStr(NewCheckedItem))
      End If
   End If

   I = InStr(wCaption, sType)
   I = InStr(I, wCaption, "|")

   If Len(NewUserID) Then
      ReturnComponentValue wCaption, "ID:", sValue

      If sValue = "" Then
         wCaption = Left$(wCaption, Len(wCaption) - 1) & ":ID:" & NewUserID & "|"
      Else
         wCaption = Replace$(wCaption, ":ID:" & sValue, ":ID:" & NewUserID)
      End If
   End If

   If Len(NewGrouping) Then
      ReturnComponentValue wCaption, "Group:", sValue

      If sValue = "" Then
         wCaption = Left$(wCaption, Len(wCaption) - 1) & ":Group:" & NewGrouping & "|"
      Else
         wCaption = Replace$(wCaption, ":Group:" & sValue, ":Group:" & NewGrouping)
      End If
   End If

   If NewFontType > lv_fAllFonts - 1 And NewFontType < lv_fNonTrueType + 1 Then
      ReturnComponentValue wCaption, "Type:", sValue

      If sValue = "" Then
         wCaption = Left$(wCaption, Len(wCaption) - 1) & ":Type:" & Choose(NewFontType + 1, "ALL", "TrueType", "System") & "|"
      Else
         wCaption = Replace$(wCaption, ":Type:" & sValue, ":Type:" & Choose(NewFontType + 1, "ALL", "TrueType", "System"))
      End If
   End If

   ChangeCustomMenu = Replace$(oldCaption & wCaption & oldWcaption, "||", "|") & "}"
End Function

Private Function BuildSimpleCustomMenu(CaptionNow As String, lvType As String, Optional UserID As String, Optional CheckedItem As String) As String
   Dim newCaption As String
   Dim wCaption As String
   Dim sValue As String
   Dim sCode As String

   sCode = lvType & CheckedItem
   If Len(UserID) Then sCode = sCode & ":ID:" & UserID
   SeparateCaption CaptionNow, newCaption, wCaption
   ReturnComponentValue wCaption, lvType, sValue

   If Len(sValue) Then
      wCaption = Trim(Replace$(wCaption, lvType & sValue & "|", ""))
      wCaption = Trim(Replace$(wCaption, lvType & sValue, ""))
   End If

   If Len(wCaption) Then
      wCaption = Replace$(wCaption, "{", "")
      wCaption = Replace$(wCaption, "}", "")

      If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption

      wCaption = "{" & sCode & wCaption & "}"
   Else
      wCaption = "{" & sCode & "}"
   End If

   BuildSimpleCustomMenu = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Sub SetMenu(hwnd As Long, Optional ImageList As Control = Nothing, Optional TipClass As clsTips = Nothing, Optional ContainerType As SubClassContainers = 0)
   If bAmDebugging Then Exit Sub

   If colMenuItems Is Nothing Then
      If Not bModuleInitialized Then LoadDefaultColors
      Set colMenuItems = New Collection

      DetermineOS
      CreateDestroyMenuFont True, False
   End If

   bUseHourglass = False

   On Error Resume Next

   Dim cMenu As clsMenuItems
   Dim cHwnd As Long
   Dim sHwnd As Long
   Dim pTips As Long
   Dim targetHwnd As Long
   Dim lFlags As Long

   If Abs(CLng(ContainerType)) = lv_VB_Toolbar Then
      targetHwnd = GetToolTipWindow(hwnd)
   Else
      targetHwnd = hwnd
   End If

   If targetHwnd = 0 Then Exit Sub

   Set cMenu = colMenuItems("h" & targetHwnd)

   If Err = 0 Then
      Set cMenu = Nothing

      Select Case Abs(ContainerType)
         Case lv_MDIchildForm_NoMenus: colMenuItems("h" & targetHwnd).IsMenuLess = True

         Case lv_MDIform_ChildrenHaveMenus, lv_MDIform_ChildrenMenuless
            If FindWindowEx(targetHwnd, 0, "MDIClient", "") Then
               colMenuItems("h" & targetHwnd).IsMenuLess = (ContainerType = lv_MDIform_ChildrenMenuless)
            End If

         Case lv_MDIchildForm_WithMenus: colMenuItems("h" & targetHwnd).IsMenuLess = False

      End Select

      hWndRedirect = "h0"

      Exit Sub
   End If

   On Error GoTo 0

   If Not TipClass Is Nothing Then pTips = ObjPtr(TipClass)

   Set cMenu = New clsMenuItems

   cMenu.IsMenuLess = (ContainerType = lv_MDIchildForm_NoMenus Or ContainerType = lv_MDIform_ChildrenMenuless)
   colMenuItems.Add cMenu, "h" & targetHwnd

   Set cMenu = Nothing

   colMenuItems(colMenuItems.Count).InitializeSubMenu targetHwnd, ImageList, pTips, False
   colMenuItems(colMenuItems.Count).hPrevProc = SetWindowLong(targetHwnd, GWL_WNDPROC, AddressOf MenuMessages)
   cHwnd = FindWindowEx(targetHwnd, 0, "MDIClient", "")

   If cHwnd Then
      Set cMenu = New clsMenuItems

      colMenuItems.Add cMenu, "h" & cHwnd

      Set cMenu = Nothing

      colMenuItems(colMenuItems.Count).InitializeSubMenu cHwnd, , , True
      colMenuItems(colMenuItems.Count).hPrevProc = SetWindowLong(cHwnd, GWL_WNDPROC, AddressOf MenuMessages)
      colMenuItems("h" & hwnd).MDIClient = cHwnd
      colMenuItems("h" & cHwnd).IsMDIclient = True
   End If

   If hWndRedirect = "MDIchildToSubclass" And cHwnd = 0 Then
      cHwnd = GetParent(GetParent(hwnd))
      colMenuItems("h" & hwnd).InitializeSubMenu hwnd, colMenuItems("h" & cHwnd).ImageListObject, colMenuItems("h" & cHwnd).ShowTips

      If ContainerType < lv_MDIchildForm_WithMenus Then
         colMenuItems("h" & hwnd).IsMenuLess = colMenuItems("h" & cHwnd).IsMenuLess
      End If
   End If

   hWndRedirect = "h0"
End Sub

Public Sub SetPopupParentForm(hwnd As Long)
   tempRedirect = hwnd
End Sub

Public Sub PopupMenuCustom(MenuFormsHwnd As Long, CustomCaption As String, _
   Optional Flags As Long, Optional x As Long = -1, Optional y As Long = -1, _
   Optional TipsReRoute As clsTips = Nothing)

   Dim newMenu As Long
   Dim lReturn As Long
   Dim lPopup As Long
   Dim menuPT As POINTAPI
   Dim pRect As RECT
   Dim sPopupCaption As String

   If x < 0 Or y < 0 Then
      GetCursorPos menuPT
   Else
      menuPT.x = x: menuPT.y = y
   End If

   newMenu = CreatePopupMenu()

   If newMenu = 0 Then Exit Sub
   If Not TipsReRoute Is Nothing Then RerouteTips MenuFormsHwnd, TipsReRoute

   SetPopupParentForm MenuFormsHwnd
   sPopupCaption = "PopMenu" & CustomCaption
   AppendMenu newMenu, MF_DISABLED, 32500, sPopupCaption
   colMenuItems("h" & MenuFormsHwnd).IsWindowList newMenu, False

   If ((Flags And TPM_NONOTIFY) = TPM_NONOTIFY) Then Flags = Flags And Not TPM_NONOTIFY

   Flags = Flags Or TPM_RETURNCMD
   lPopup = GetSubMenu(newMenu, 0)
   lReturn = TrackPopupMenu(lPopup, Flags, menuPT.x, menuPT.y, 0&, MenuFormsHwnd, pRect)

   If lReturn Then
      colMenuItems("h" & MenuFormsHwnd).MenuSelected lReturn, lPopup, 0
   End If

   colMenuItems("h" & MenuFormsHwnd).DestroyPopup newMenu, 32500
   DestroyMenu newMenu
   DoTips 0, 0, 0

   If Not TipsReRoute Is Nothing Then RerouteTips MenuFormsHwnd
End Sub

Public Sub RerouteTips(hwnd As Long, Optional TipsClass As clsTips = Nothing)
   Dim lStyle As Long

   On Error Resume Next

   lStyle = colMenuItems("h" & hwnd).ShowTips

   If Err Then
      lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)

      If ((lStyle And WS_EX_MDICHILD) = WS_EX_MDICHILD) Then
         On Error Resume Next

         hWndRedirect = "MDIchildToSubclass"
         SetMenu hwnd

         If Not TipsClass Is Nothing Then colMenuItems("h" & hwnd).InitializeSubMenu hwnd, , ObjPtr(TipsClass)
      Else
         SetMenu hwnd, , TipsClass
      End If

      Err.Clear
   Else
      If TipsClass Is Nothing Then lStyle = 0 Else lStyle = ObjPtr(TipsClass)

      colMenuItems("h" & hwnd).InitializeSubMenu hwnd, , lStyle, True
   End If
End Sub

Public Sub CleanClass(hwnd As Long)
   Dim sysHwnd As Long

   Set OpenMenus = Nothing

   SetWindowLong hwnd, GWL_WNDPROC, colMenuItems("h" & hwnd).hPrevProc
   sysHwnd = colMenuItems("h" & hwnd).SystemMenu

   If AmInIDE Then colMenuItems("h" & hwnd).RestoreMenus

   If colMenuItems.Count = 1 Then
      CreateDestroyMenuFont False, False

      Set colMenuItems = Nothing

      If FloppyIcon Then DestroyIcon FloppyIcon

      FloppyIcon = 0
      Erase tbarClass
   Else
      colMenuItems.Remove "h" & hwnd
   End If

   On Error Resume Next

   If sysHwnd Then GetSystemMenu sysHwnd, 1
End Sub

Public Sub CreateDestroyMenuFont(bCreate As Boolean, ItalicFonts As Boolean, Optional FontSample As String)
   If bCreate = True Then
      Dim ncm As NONCLIENTMETRICS, newFont As LOGFONT, oldWT As Long

      ncm.cbSize = Len(ncm)
      SystemParametersInfo 41, 0, ncm, 0
      newFont = ncm.lfMenuFont
      newFont.lfCharSet = 1
      newFont.lfFaceName = MenuFontName & Chr$(0)
      newFont.lfHeight = (mFontSize * -20) / Screen.TwipsPerPixelY

      If ItalicFonts Then
         If m_Font(4) = 0 Then
            newFont.lfItalic = 1
            m_Font(4) = CreateFontIndirect(newFont)
            newFont.lfWeight = 800
            m_Font(5) = CreateFontIndirect(newFont)
         End If
      Else
         If Len(FontSample) Then
            If m_Font(6) Then DeleteObject m_Font(6)
            If m_Font(7) Then DeleteObject m_Font(7)
            newFont.lfWeight = 400
            newFont.lfFaceName = FontSample & Chr$(0)
            m_Font(6) = CreateFontIndirect(newFont)
            newFont.lfItalic = 1
            m_Font(7) = CreateFontIndirect(newFont)
         Else
            m_Font(1) = CreateFontIndirect(newFont)
            oldWT = newFont.lfWeight
            newFont.lfWeight = 800
            m_Font(2) = CreateFontIndirect(newFont)
            newFont.lfWeight = oldWT

            If (mFontSize * 0.8) < 7.5 Then
                newFont.lfHeight = ((mFontSize - 0.5) * -20) / Screen.TwipsPerPixelY
            Else
                newFont.lfHeight = ((mFontSize * 0.8) * -20) / Screen.TwipsPerPixelY
            End If

            m_Font(3) = CreateFontIndirect(newFont)     ' mini-font
         End If
      End If
   Else
      On Error Resume Next

      DeleteObject m_Font(1)
      DeleteObject m_Font(2)
      DeleteObject m_Font(3)

      If m_Font(4) Then DeleteObject m_Font(4)
      If m_Font(5) Then DeleteObject m_Font(5)
      If m_Font(6) Then DeleteObject m_Font(6)
      If m_Font(7) Then DeleteObject m_Font(7)

      Erase m_Font
   End If
End Sub

Public Sub ApplyMenuFont(FontID As Integer, hdc As Long)
   If hdc = 0 Then Exit Sub

   If FontID Then
      m_Font(0) = SelectObject(hdc, m_Font(FontID))
   Else
      SelectObject hdc, m_Font(0)
   End If
End Sub

Private Sub DoTips(wParam As Long, lParam As Long, hMenu As Long)
   Dim hWord As Integer
   Dim tipProc As Long
   Dim lWord As Long
   Dim sTip As String
   Dim bMenuClosed As Boolean

   tipProc = colMenuItems(hWndRedirect).ShowTips

   On Error Resume Next

   If hMenu Then
      If wParam <> 0 Or lParam <> 0 Then
         lWord = CLng(LoWord(wParam))
         hWord = HiWord(wParam)

         If ((hWord And MF_POPUP) = MF_POPUP) Then lWord = GetSubMenu(lParam, lWord)

         sTip = colMenuItems(hWndRedirect).Tips(lWord, hMenu)
      End If
   Else
      If wParam = 0 And lParam = 0 Then bMenuClosed = True
   End If

   If tipProc = 0 Then Exit Sub

   Dim oTipClass As clsTips

   CopyMemory oTipClass, tipProc, 4&
   oTipClass.SendTip sTip

   If bMenuClosed Then oTipClass.SendCustomSelection "", "MenusClosed", 0&

   CopyMemory oTipClass, 0&, 4&

   Set oTipClass = Nothing
End Sub

Private Function DoMeasureItem(lParam As Long) As Boolean
   Dim MeasureInfo As MEASUREITEMSTRUCT

   Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo))

   If MeasureInfo.CtlType <> ODT_MENU Then Exit Function

   colMenuItems(hWndRedirect).GetMenuItem MeasureInfo.ItemID, MeasureInfo.ItemData
   MeasureInfo.ItemHeight = XferMenuData.Dimension.y
   MeasureInfo.ItemWidth = XferMenuData.Dimension.x + XferMenuData.OffsetCx
   Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
   DoMeasureItem = True
End Function

Private Function DoDrawItem(lParam As Long) As Boolean
   Dim DrawInfo As DRAWITEMSTRUCT
   Dim IsSep As Boolean
   Dim IsDisabled As Boolean
   Dim sysIcon As Boolean
   Dim bSelectDisabled As Boolean
   Dim Xoffset As Integer
   Dim yIconOffset As Integer
   Dim isSelected As Boolean
   Dim IsChecked As Boolean
   Dim bGradientFill As Boolean
   Dim sFont As String
   Dim iFont As Integer
   Dim lTextColor As Long
   Dim cBack As Long
   Dim tRect As RECT
   Dim mItem As RECT

   On Error GoTo ShowErrors

   Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo))

   If DrawInfo.CtlType <> ODT_MENU Then Exit Function
   DoDrawItem = True
   colMenuItems(hWndRedirect).GetMenuItem DrawInfo.ItemID, DrawInfo.ItemData
   colMenuItems(hWndRedirect).GetPanelItem DrawInfo.hWndItem

   With DrawInfo
      IsSep = ((XferMenuData.Status And lv_mSep) = lv_mSep)
      IsDisabled = ((XferMenuData.Status And lv_mDisabled) = lv_mDisabled)

      If (.itemAction And Not ODA_DRAWENTIRE) And IsSep Then Exit Function
      IsChecked = ((XferMenuData.Status And lv_mChk) = lv_mChk)
      isSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED)
      bSelectDisabled = (bHiLiteDisabled) Or (XferPanelData.IsSystem) Or (bKeyBoardSelect)
      cBack = vbWhite 'GetSysColor(COLOR_MENU)

      Select Case LoWord(.ItemID)
         Case SC_CLOSE, SC_MAXIMIZE, SC_MINIMIZE, SC_RESTORE: sysIcon = True
         Case Else: sysIcon = False
      End Select

      If ((XferMenuData.Status And lv_mSBar) = lv_mSBar) Then
         If XferPanelData.PanelIcon <> 0 Then
            If (.itemAction And ODA_DRAWENTIRE) Or Not IsDisabled Then
               Dim tDC As Long
               Dim sDC As Long
               Dim oldBMP As Long

               sDC = GetDC(CLng(Mid$(hWndRedirect, 2)))
               tDC = CreateCompatibleDC(sDC)
               oldBMP = SelectObject(tDC, XferPanelData.PanelIcon)

               mItem = .rcItem
               mItem.Right = mItem.Right + 2
               
               DrawGradient cBack, vbRed, False, .hdc, mItem
               BitBlt .hdc, 1, 1, .rcItem.Right, .rcItem.Bottom - 2, tDC, 0, 0, vbSrcCopy

               SelectObject tDC, oldBMP
               DeleteDC tDC
               ReleaseDC CLng(Mid$(hWndRedirect, 2)), sDC
               DeleteObject oldBMP
            End If
         End If

         Exit Function
      End If

      mItem = .rcItem
      mItem.Right = mItem.Left + 25
      DrawGradient GradCol1, GradCol2, True, .hdc, mItem

      mItem = .rcItem
      SetBkMode .hdc, NEWTRANSPARENT

      If ((Len(XferMenuData.Icon) = 0 And IsChecked = False) Or _
         XferPanelData.HasIcons = False) And sysIcon = False Then mItem.Left = .rcItem.Left

      If ((XferMenuData.Status And lv_mColor) = lv_mColor) Then
         If InStr(XferMenuData.Caption, "LColor:-1") = 0 Then mItem.Right = mItem.Right - 38
      End If

      If (Not IsSep And Not IsDisabled) Or (IsDisabled And bSelectDisabled) Then
         If Not bGradientSelect Or (IsSep Or Not isSelected) Then
            'DrawRect .hDC, mItem.Left + 25, mItem.Top, mItem.Right, mItem.Bottom, cBack
            mItem.Left = mItem.Left + 25
            DrawGradient GradCol2, GradBackCol, True, .hdc, mItem
            mItem.Left = mItem.Left - 25
         Else
'            DrawGradient GradCol2, cBack, True, .hDC, mItem
            DrawRect .hdc, mItem.Left, mItem.Top, mItem.Right, mItem.Bottom, SelFrameCol

            mItem.Left = mItem.Left + 1
            mItem.Top = mItem.Top + 1
            mItem.Right = mItem.Right - 1
            mItem.Bottom = mItem.Top + Int((mItem.Bottom - mItem.Top) / 2)

            If IsDisabled Then
               DrawGradient DisSelGradBackCol, DisSelGradCol, False, .hdc, mItem
               mItem.Top = mItem.Bottom
               mItem.Bottom = .rcItem.Bottom - 1
               DrawGradient DisSelGradCol, DisSelGradBackCol, False, .hdc, mItem
           Else
               DrawGradient SelGradBackCol, SelGradCol, False, .hdc, mItem
               mItem.Top = mItem.Bottom
               mItem.Bottom = .rcItem.Bottom - 1
               DrawGradient SelGradCol, SelGradBackCol, False, .hdc, mItem
            End If
            
            bGradientFill = True
         End If
      End If

      If IsSep Then
         mItem = .rcItem

         If Len(XferMenuData.Display) Then
            tRect = .rcItem
            OffsetRect tRect, 0, 1
            ApplyMenuFont 3, .hdc
            SetMenuColor True, .hdc, cObj_Text, TextColorSeparatorBar
            DrawText .hdc, XferMenuData.Display, Len(XferMenuData.Display), tRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_CENTER Or DT_VCENTER
            OffsetRect tRect, (mItem.Right - tRect.Right) \ 2, 0
            DrawText .hdc, XferMenuData.Display, Len(XferMenuData.Display), tRect, DT_SINGLELINE Or DT_NOCLIP Or DT_VCENTER
            ThreeDbox .hdc, mItem.Left + 2, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, tRect.Left - 0, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, True, ((XferMenuData.Status And lv_mSepRaised) = lv_mSepRaised), True, 1, 1, RGB(31, 31, 31), RGB(191, 191, 191)
            ThreeDbox .hdc, tRect.Right + 3, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, mItem.Right, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, True, ((XferMenuData.Status And lv_mSepRaised) = lv_mSepRaised), True, 1, 1, RGB(31, 31, 31), RGB(191, 191, 191)
         Else
            mItem.Left = mItem.Left + 25
            mItem.Right = mItem.Left + Int((mItem.Right - mItem.Left) / 2)
            mItem.Top = mItem.Top + 2
            mItem.Bottom = mItem.Bottom - 2
            DrawGradient GradCol2, SepCol, True, .hdc, mItem

            mItem.Left = mItem.Right
            mItem.Right = .rcItem.Right
            DrawGradient SepCol, GradBackCol, True, .hdc, mItem
            
            mItem.Left = .rcItem.Left + 25
            mItem.Top = mItem.Top - 1
            mItem.Bottom = mItem.Top + 2

            DrawGradient GradCol2, GradBackCol, True, .hdc, mItem
            'DrawRect .hDC, .rcItem.Left + 25, mItem.Top - 1, mItem.Right, mItem.Top + 1, cBack
         End If

         ApplyMenuFont 0, .hdc
      Else
         mItem = .rcItem

         If Len(XferMenuData.Icon) > 0 Or (IsChecked And XferPanelData.HasIcons) Then
            If IsChecked And Not IsDisabled And (.itemAction And ODA_DRAWENTIRE Or bRaisedIcons) Then
               yIconOffset = -1
            End If
         End If

         If IsChecked And (Len(XferMenuData.Icon) = 0 And Not sysIcon) Then
            mItem = .rcItem

            If XferPanelData.HasIcons Or XferPanelData.IsSystem Then
               If IsDisabled Then
                  OffsetRect mItem, 1, 1
                  DrawCheckMark False, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledLight, TextColorDisabledDark
               Else
                  DrawCheckMark False, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorNormal
               End If
            Else
               If IsDisabled Then
                  If bSelectDisabled And isSelected Then
                     DrawCheckMark False, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledDark
                  Else
                     OffsetRect mItem, 1, 1
                     DrawCheckMark False, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledLight, TextColorDisabledDark
                  End If
               Else
                  If isSelected Then
                     DrawCheckMark True, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorSelected
                  Else
                     DrawCheckMark False, .hdc, mItem, .rcItem.Left, bXPcheckmarks, TextColorNormal
                  End If
               End If
            End If
         End If

         mItem = .rcItem
         mItem.Left = .rcItem.Left + 29
         iFont = 1 + Abs(CInt((XferMenuData.Status And lv_mDefault) = lv_mDefault))
         ApplyMenuFont iFont, .hdc
         SetMenuColor True, .hdc, cObj_Text, vbBlack
         DrawText .hdc, XferMenuData.Display, Len(XferMenuData.Display), mItem, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER

         If Len(XferMenuData.HotKey) Then
            mItem.Right = .rcItem.Right - 15

            If ((XferMenuData.Status And lv_mFont) = lv_mFont) Then
               ReturnComponentValue XferMenuData.Caption, "LFont:", sFont
               ApplyMenuFont 0, .hdc
               CreateDestroyMenuFont True, False, sFont
               ApplyMenuFont (Abs(CInt(bItalicSelected And isSelected))) + 6, .hdc
            End If

            DrawText .hdc, XferMenuData.HotKey, Len(XferMenuData.HotKey), mItem, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
         End If

         If (IsDisabled And bSelectDisabled) Or (IsDisabled And Not isSelected) Then
            mItem.Right = .rcItem.Right
            SetMenuColor True, .hdc, cObj_Text, RGB(195, 195, 195)
            DrawText .hdc, XferMenuData.Display, Len(XferMenuData.Display), mItem, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER

            If Len(XferMenuData.HotKey) Then
               mItem.Right = .rcItem.Right - 16

               If ((XferMenuData.Status And lv_mFont) = lv_mFont) Then
                  ApplyMenuFont 0, .hdc
                  ApplyMenuFont (Abs(CInt(bItalicSelected And isSelected))) + 6, .hdc
               End If

               DrawText .hdc, XferMenuData.HotKey, Len(XferMenuData.HotKey), mItem, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
            Else

            End If
         End If

         ApplyMenuFont 0, .hdc

         If Len(XferMenuData.Icon) > 0 Or sysIcon Then
            mItem = .rcItem

            If Len(XferMenuData.Icon) Then
               mItem.Left = mItem.Left + 4
               mItem.Top = mItem.Top + ((mItem.Bottom - mItem.Top) - 16) \ 2

               If Not IsDisabled And isSelected Then
                  DrawMenuIcon .hdc, XferMenuData.Icon, 0, mItem, False, , Abs(CInt(XferMenuData.ShowBKG)), 0, 0
                  DrawMenuIcon .hdc, XferMenuData.Icon, 0, mItem, False, , Abs(CInt(XferMenuData.ShowBKG)), -1, -1
               Else
                  If (.itemAction = ODA_DRAWENTIRE) Or bRaisedIcons Then
                     DrawMenuIcon .hdc, XferMenuData.Icon, 0, mItem, IsDisabled, , Abs(CInt(XferMenuData.ShowBKG)), -1, yIconOffset
                  End If
               End If
            Else
               DrawSystemIcon .hdc, mItem, .ItemID, IsDisabled, isSelected
            End If
         End If

         If (.itemAction And ODA_DRAWENTIRE) Then
            If ((XferMenuData.Status And lv_mColor) = lv_mColor) Then
               Dim sColor As String

               Xoffset = InStr(XferMenuData.Caption, "{")
               ReturnComponentValue Mid$(XferMenuData.Caption, Xoffset), "LColor:", sColor

               If Val(sColor) <> -1 Then DrawRect .hdc, .rcItem.Right - 35, .rcItem.Top + 2, .rcItem.Right - 5, .rcItem.Bottom - 2, Val(sColor), vbBlack
            End If
         End If
      End If
   End With

ShowErrors:
End Function

Public Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function

Public Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
   MakeLong = CLng(LoWord)
   Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
End Function

Private Function MenuMessages(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim sMsg As String
   Static LastHmenu As Long
   Static CursorType As Integer

   On Error GoTo AllowMsgThru

   Select Case uMsg
      Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
         MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)

         If OpenMenus Is Nothing And bReturnMDIkeystrokes = True Then
            ProcessKeyStroke hwnd, wParam, ((uMsg = WM_KEYUP) Or (uMsg = WM_SYSKEYUP))
            Exit Function
         End If

   Case WM_GETMINMAXINFO
       If colMenuItems("h" & hwnd).RestrictSize(lParam, False) = True Then Exit Function

   Case WM_ENTERMENULOOP
      If Screen.MousePointer <> vbHourglass Then CursorType = Screen.MousePointer + 1

      sMsg = "WM_Entermenuloop"

      Set OpenMenus = New Collection    ' create new index of opened menus

   Case WM_MDIACTIVATE
      sMsg = "WM_MdiActivate"

      If hwnd = GetParent(wParam) Then
         MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
         hWndRedirect = "MDIchildToSubclass"
         SetMenu wParam
         Exit Function
      End If

   Case WM_MDIMAXIMIZE
      MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
      hWndRedirect = "MDIchildToSubclass"
      SetMenu wParam
      Exit Function

   Case WM_MENUCHAR
      MenuMessages = IdentifyAccelerator(LoWord(wParam), lParam)
      Exit Function

   Case WM_INITMENUPOPUP
      sMsg = "WM_INitmenupopup"
      MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
      IDCurrentWindow hwnd, wParam, (HiWord(lParam) <> 0)
      colMenuItems(hWndRedirect).GetPanelItem wParam

      If XferPanelData.Hourglass = True And CursorType - 1 <> vbHourglass Then
         Screen.MousePointer = vbHourglass
         bUseHourglass = True
      Else
         bUseHourglass = False
      End If

      On Error Resume Next

      wParam = OpenMenus("m" & wParam)    ' if menu is already open, no need to mess with it

      If Err = 0 Then Exit Function

      On Error GoTo AllowMsgThru

      If colMenuItems(hWndRedirect).IsWindowList(wParam, False) = False Then OpenMenus.Add wParam, "m" & wParam

      Exit Function

   Case WM_MEASUREITEM
      If DoMeasureItem(lParam) = True Then Exit Function

   Case WM_DRAWITEM
      If DoDrawItem(lParam) = True Then Exit Function

   Case WM_MENUSELECT
      sMsg = "WM_Menuselect"
      If lParam Then LastHmenu = lParam
      DoTips wParam, lParam, LastHmenu
      bKeyBoardSelect = ((HiWord(wParam) And MF_MOUSESELECT) <> MF_MOUSESELECT)

   Case WM_COMMAND
      sMsg = "WM_Command"

      If HiWord(wParam) = 0 And lParam = 0 Then
         If colMenuItems(hWndRedirect).MenuSelected(CLng(LoWord(wParam)), LastHmenu, 0) Then Exit Function  ' LoWord is menu ID
      End If

   Case WM_MENUCOMMAND
      sMsg = "WM_menucommand"
      colMenuItems(hWndRedirect).MenuSelected GetMenuItemID(lParam, LoWord(wParam)), lParam, 0

   Case WM_EXITMENULOOP
      sMsg = "WM_exitmenuloop"

      If bUseHourglass Then
         Screen.MousePointer = CursorType - 1
         bUseHourglass = False
      End If

      DoTips 0, 0, 0
      tempRedirect = 0

      Dim lMenus As Integer

      For lMenus = 1 To OpenMenus.Count
         colMenuItems(hWndRedirect).UpdateMenuItems OpenMenus.Item(lMenus)
      Next

      Set OpenMenus = Nothing       ' clear collection of opened menus

   Case WM_DESTROY
      sMsg = "WM_Destroy"

      On Error Resume Next

      MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
      CleanClass hwnd
      Exit Function

   Case WM_MDIDESTROY
      sMsg = "WM_MDIdestroy"

      On Error Resume Next

      MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
      CleanClass wParam
      Exit Function

   Case WM_ENTERIDLE
      If bUseHourglass Then
         Screen.MousePointer = CursorType - 1
         bUseHourglass = False
      End If

   End Select

AllowMsgThru:
   On Error Resume Next

   MenuMessages = CallWindowProc(colMenuItems("h" & hwnd).hPrevProc, hwnd, uMsg, wParam, lParam)
End Function

Private Function DetermineOS(Optional SetGraphicsModeDC As Long = 0) As Integer
   Const os_Win95 = "1.4.0"
   Const os_Win98 = "1.4.10"
   Const os_WinNT4 = "2.4.0"
   Const os_WinNT351 = "2.3.51"
   Const os_Win2K = "2.5.0"
   Const os_WinME = "1.4.90"
   Const os_WinXP = "2.5.1"

   Dim verinfo As OSVERSIONINFO
   Dim sVersion As String

   verinfo.dwOSVersionInfoSize = Len(verinfo)

   If (GetVersionEx(verinfo)) = 0 Then Exit Function

   With verinfo
      sVersion = .dwPlatformId & "." & .dwMajorVersion & "." & .dwMinorVersion
   End With

   Select Case sVersion
      Case os_Win98: ExtraOffsetX = 35
      Case os_Win2K: ExtraOffsetX = 0
      Case os_WinNT4: ExtraOffsetX = 0
      Case os_WinNT351: SetGraphicsMode SetGraphicsModeDC, 2
      Case os_Win95
      Case os_WinXP: ExtraOffsetX = 0
      Case os_WinME: ExtraOffsetX = 35
   End Select
End Function

Public Sub ReturnComponentValue(sSource As String, sTarget As String, sRtnVal As String)
   sRtnVal = ""

   If Len(sSource) < 3 Then Exit Sub

   Dim cI As Integer
   Dim cJ As Integer
   Dim sSourceNoTip As String

   sSourceNoTip = sSource
   cI = InStr(sSource, "|Tip:")

   If cI = 0 Then cI = InStr(sSource, "{Tip:")

   If cI Then
      cI = cI + 1
      cJ = InStr(cI + 1, sSource, "|")

      If cJ = 0 Then cJ = InStr(cI + 1, sSource, "}")

      If cJ Then
         If sTarget <> "Tip:" Then sSourceNoTip = Replace$(sSource, Mid$(sSource, cI, cJ - cI), "")
      End If
   End If

   If sTarget <> "Tip:" Then
      cI = InStr(sSourceNoTip, "|" & sTarget)

      If cI = 0 Then cI = InStr(sSourceNoTip, "{" & sTarget)
      If cI = 0 Then cI = InStr(sSourceNoTip, ":" & sTarget)
      If cI = 0 Then Exit Sub

      cI = cI + 1
      cJ = InStr(cI + 1, sSourceNoTip, "|")

      If cJ = 0 Then cJ = InStr(cI + 1, sSourceNoTip, "}")
   End If

   If cJ Then
      sRtnVal = Mid$(sSourceNoTip, cI, cJ - cI)
      sRtnVal = Mid$(sRtnVal, Len(sTarget) + 1)

      If Len(sRtnVal) = 0 Then
         Select Case sRtnVal
            Case "vbAlignBottom": sRtnVal = "Bot"
            Case "vbAlignTop": sRtnVal = "Top"
            Case "vbNull": sRtnVal = "1"
            Case Else
                Select Case sTarget
                  Case "Raised", "Bold", "Italic", "Underline", "NoScroll", "Transparent", "ImgBkg", "Default", "Sidebar", "SBDisabled"
                     sRtnVal = "True"
                  Case Else: sRtnVal = "0"
               End Select
         End Select
      End If
   End If
End Sub

Public Sub SeparateCaption(sSource, sNewCaption, sCodeCaption, Optional sHotKey)
   Dim IdX As Integer
   Dim tSource As String

   tSource = sSource
   sHotKey = ""
   sNewCaption = ""
   sCodeCaption = ""
   IdX = InStr(tSource, vbTab)

   If IdX Then
      sHotKey = Mid$(tSource, IdX + 1)
      If IdX > 1 Then tSource = Left$(tSource, IdX - 1) Else tSource = ""
   End If

   sNewCaption = Trim$(tSource)

   If Len(tSource) Then
      IdX = InStr(tSource, "{")

      If IdX Then
         If InStr(IdX, tSource, "}") = 0 Then Exit Sub

         sCodeCaption = Mid$(tSource, IdX, InStr(IdX, tSource, "}") - IdX + 1)
         sNewCaption = Trim$(Replace$(tSource, sCodeCaption, ""))
         sCodeCaption = Left$(sCodeCaption, Len(sCodeCaption) - 1) & "|"

         If AmInIDE Then
            Dim Idy As Integer

            IdX = InStr(sCodeCaption, "IMG:")

            If IdX Then Idy = InStr(IdX, sCodeCaption, "|")

            If Idy Then
               sCodeCaption = Replace$(sCodeCaption, Mid$(sCodeCaption, IdX, Idy - IdX), "IMG:" & DefaultIcon)
            End If
         End If
      End If
   End If
End Sub

Private Function DrawMenuIcon(m_hDC As Long, sImgID As String, imageType As Long, _
   rt As RECT, bDisabled As Boolean, Optional bDisableColored As Boolean = True, _
   Optional bNoTransparency As Long = 0, Optional iOffset As Integer = 0, _
   Optional Yoffset As Integer, Optional imgWidth As Integer = 16, _
   Optional imgHeight As Integer = 16, Optional lMask As Long = -1) As Boolean

   Dim tDC As Long
   Dim lPrevImage As Long
   Dim lImageType As Long
   Dim lImgCopy As Long
   Dim lImageHdl As Long
   Dim sImgHandle As String
   Dim bmpInfo As BITMAP
   Dim icoInfo As ICONINFO
   Dim rcImage As RECT
   Dim dRect As RECT
   Dim shInfo As SHFILEINFO
   Dim exeIcon As Long

   Const CI_BITMAP = &H0
   Const CI_ICON = &H1

   sImgHandle = sImgID

   If IsNumeric(sImgHandle) Or Len(sImgHandle) = 0 Then
      lImageHdl = Val(sImgHandle)
   Else
      SHGetFileInfo sImgHandle, 0, shInfo, Len(shInfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES
      lImageHdl = shInfo.hIcon
   End If

   If lImageHdl = 0 Then Exit Function

   GetObject lImageHdl, Len(bmpInfo), bmpInfo

   If bmpInfo.bmBits Then
      lImageType = CI_BITMAP
      lImgCopy = CopyImage(lImageHdl, CI_BITMAP, imgWidth, imgHeight, 0)
   Else
      GetIconInfo lImageHdl, icoInfo

      If icoInfo.hbmColor <> 0 Then
         lImageType = CI_ICON
         DeleteObject icoInfo.hbmColor

         If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
      Else
         Exit Function
      End If
   End If

   dRect = rt
   OffsetRect dRect, iOffset, Yoffset
   DrawMenuIcon = True

   If Not bDisabled Then
      If lImageType = CI_ICON Then
         exeIcon = CopyImage(lImageHdl, CI_ICON, imgWidth, imgHeight, 0)
         DrawIconEx m_hDC, dRect.Left, dRect.Top, exeIcon, 0, 0, 0, 0, &H3
         DestroyIcon exeIcon

         If shInfo.hIcon Then DestroyIcon shInfo.hIcon
      Else
         If bNoTransparency = 1 Then
            tDC = CreateCompatibleDC(m_hDC)
            lPrevImage = SelectObject(tDC, lImgCopy)
            StretchBlt m_hDC, dRect.Left, dRect.Top, imgWidth, imgHeight, tDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, vbSrcCopy
         Else
            DrawTransparentBitmap m_hDC, dRect, lImgCopy, rcImage, , CLng(imgWidth), CLng(imgHeight)
         End If

         DeleteObject lImgCopy
      End If

      If IsNumeric(sImgID) = False Then DestroyIcon lImageHdl
   Else
      Const MAGICROP = &HB8074A

      Dim hBitmap As Long
      Dim hOldBitmap As Long
      Dim hMemDC As Long
      Dim hOldBrush As Long
      Dim hOldBackColor As Long
      Dim hbrShadow As Long
      Dim hbrHilite As Long

      hMemDC = CreateCompatibleDC(m_hDC)
      hBitmap = CreateCompatibleBitmap(m_hDC, imgWidth, imgHeight)
      hOldBitmap = SelectObject(hMemDC, hBitmap)
      PatBlt hMemDC, 0, 0, imgWidth, imgHeight, &HFFFF62   ' WHITENESS

      dRect = rt

      If lImageType = CI_ICON Then
         exeIcon = CopyImage(lImageHdl, CI_ICON, imgWidth, imgHeight, 0)
         DrawIconEx hMemDC, 0, 0, exeIcon, 0, 0, 0, 0, &H3
         DestroyIcon exeIcon

         If shInfo.hIcon Then DestroyIcon shInfo.hIcon
      Else
         If bNoTransparency = 1 Then
            tDC = CreateCompatibleDC(m_hDC)
            lPrevImage = SelectObject(tDC, lImgCopy)
            StretchBlt hMemDC, 0, 0, imgWidth, imgHeight, tDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, vbSrcCopy
         Else
            OffsetRect dRect, rt.Left * -1, rt.Top * -1
            DrawTransparentBitmap hMemDC, dRect, lImgCopy, rcImage, , CLng(imgWidth), CLng(imgHeight)
            dRect = rt
         End If

         DeleteObject lImgCopy
      End If

      If IsNumeric(sImgID) = False Then DestroyIcon lImageHdl

      hOldBackColor = SetBkColor(m_hDC, vbWhite)
      hbrShadow = CreateSolidBrush(ConvertColor(GetSysColor(COLOR_BTNSHADOW)))

      If bDisableColored Then
         hbrHilite = CreateSolidBrush(ConvertColor(GetSysColor(COLOR_BTNHIGHLIGHT)))
         hOldBrush = SelectObject(m_hDC, hbrHilite)
         BitBlt m_hDC, dRect.Left, dRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
         SelectObject m_hDC, hbrShadow
         BitBlt m_hDC, dRect.Left, dRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
      Else
         SelectObject m_hDC, hbrShadow
         BitBlt m_hDC, dRect.Left + 1, dRect.Top + 1, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
         hbrHilite = CreateSolidBrush(hOldBackColor)
         SelectObject m_hDC, hbrHilite
         BitBlt m_hDC, dRect.Left, dRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
      End If

      SelectObject m_hDC, hOldBrush
      SetBkColor m_hDC, hOldBackColor
      SelectObject hMemDC, hOldBitmap
      DeleteObject hbrHilite
      DeleteObject hbrShadow
      DeleteObject hBitmap
      DeleteDC hMemDC
   End If

   If lPrevImage Then SelectObject tDC, lPrevImage
   If tDC Then DeleteDC tDC
End Function

Public Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, lBMPsource As Long, bmpRect As RECT, _
   Optional lMaskColor As Long = -1, Optional lNewBmpCx As Long, Optional lNewBmpCy As Long, Optional lBkgHDC As Long, _
   Optional bkgX As Long, Optional bkgY As Long, Optional FlipHorz As Boolean = False, _
   Optional FlipVert As Boolean = False, Optional srcDC As Long)

   Const DSna = &H220326

   Dim udtBitMap As BITMAP
   Dim lMask2Use As Long
   Dim lBmMask As Long
   Dim lBmAndMem As Long
   Dim lBmColor As Long
   Dim lBmObjectOld As Long
   Dim lBmMemOld As Long
   Dim lBmColorOld As Long
   Dim lHDCMem As Long
   Dim lHDCscreen As Long
   Dim lHDCsrc As Long
   Dim lHDCMask As Long
   Dim lHDCcolor As Long
   Dim OrientX As Long
   Dim OrientY As Long
   Dim x As Long
   Dim y As Long
   Dim srcX As Long
   Dim srcY As Long
   Dim lRatio(0 To 1) As Single
   Dim hPalOld As Long
   Dim hPalMem As Long

   lHDCscreen = GetDC(0&)
   lHDCsrc = CreateCompatibleDC(lHDCscreen)
   SelectObject lHDCsrc, lBMPsource
   GetObject lBMPsource, Len(udtBitMap), udtBitMap

   srcX = udtBitMap.bmWidth
   srcY = udtBitMap.bmHeight

   If lNewBmpCx = 0 Then
      If bmpRect.Right > 0 Then lNewBmpCx = bmpRect.Right - bmpRect.Left Else lNewBmpCx = srcX
   End If

   If lNewBmpCy = 0 Then
      If bmpRect.Bottom > 0 Then lNewBmpCy = bmpRect.Bottom - bmpRect.Top Else lNewBmpCy = srcY
   End If

   If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
   If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
   If (destRect.Right) = 0 Then x = lNewBmpCx Else x = (destRect.Right - destRect.Left)
   If (destRect.Bottom) = 0 Then y = lNewBmpCy Else y = (destRect.Bottom - destRect.Top)

   If lNewBmpCx > x Or lNewBmpCy > y Then
      lRatio(0) = (x / lNewBmpCx)
      lRatio(1) = (y / lNewBmpCy)

      If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)

      lNewBmpCx = lRatio(0) * lNewBmpCx
      lNewBmpCy = lRatio(0) * lNewBmpCy
      Erase lRatio
   End If

   lMask2Use = lMaskColor

   If lMask2Use < 0 Then lMask2Use = GetPixel(lHDCsrc, 0, 0)

   lMask2Use = ConvertColor(lMask2Use)
   lHDCMask = CreateCompatibleDC(lHDCscreen)
   lHDCMem = CreateCompatibleDC(lHDCscreen)
   lHDCcolor = CreateCompatibleDC(lHDCscreen)
   lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
   lBmAndMem = CreateCompatibleBitmap(lHDCscreen, x, y)
   lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
   lBmColorOld = SelectObject(lHDCcolor, lBmColor)
   lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
   lBmObjectOld = SelectObject(lHDCMask, lBmMask)

   ReleaseDC 0&, lHDCscreen

   SetMapMode lHDCMem, GetMapMode(lHDCdest)
   hPalMem = SelectPalette(lHDCMem, 0, True)
   RealizePalette lHDCMem

   If (lBkgHDC <> 0) Then
      BitBlt lHDCMem, 0, 0, x, y, lBkgHDC, bkgX, bkgY, vbSrcCopy
   Else
      BitBlt lHDCMem, 0&, 0&, x, y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
   End If

   hPalOld = SelectPalette(lHDCcolor, 0, True)
   RealizePalette lHDCcolor
   SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
   SetTextColor lHDCcolor, GetTextColor(lHDCsrc)

   BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy

   If FlipHorz Then StretchBlt lHDCcolor, srcX, 0, -srcX, srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
   If FlipVert Then StretchBlt lHDCcolor, 0, srcY, srcX, -srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy

   SetBkColor lHDCcolor, lMask2Use
   SetTextColor lHDCcolor, vbWhite

   BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy

   SetTextColor lHDCcolor, vbBlack
   SetBkColor lHDCcolor, vbWhite

   BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna
   StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
   StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
   BitBlt lHDCdest, destRect.Left, destRect.Top, x, y, lHDCMem, 0&, 0&, vbSrcCopy

   DeleteObject SelectObject(lHDCcolor, lBmColorOld)
   DeleteObject SelectObject(lHDCMask, lBmObjectOld)
   DeleteObject SelectObject(lHDCMem, lBmMemOld)

   DeleteDC lHDCMem
   DeleteDC lHDCMask
   DeleteDC lHDCcolor

   If srcDC = 0 Then DeleteDC lHDCsrc
End Sub

Public Sub DrawRect(m_hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, tColor As Long, Optional pColor As Long = -1)
   If pColor <> -1 Then SetMenuColor True, m_hDC, cObj_Pen, pColor
   SetMenuColor True, m_hDC, cObj_Brush, tColor, (pColor = -1)
   Call Rectangle(m_hDC, X1, Y1, X2, Y2)
   SetMenuColor False, m_hDC, cObj_Brush, 0
End Sub

Private Sub ThreeDbox(tHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
   ByVal X2 As Long, ByVal Y2 As Long, bSelected As Boolean, _
   Optional Sunken As Boolean = False, Optional bSeparatorBar As Boolean, _
   Optional PenWidthLt As Long = 1, Optional PenWidthDk As Long = 1, _
   Optional ColorLt As Long = -1, Optional ColorDk As Long = -1)

   If tHdc = 0 Then Exit Sub

   Dim dM As POINTAPI
   Dim iOffset As Integer

   If ColorLt = -1 Then ColorLt = SeparatorBarColorLight
   If ColorDk = -1 Then ColorDk = SeparatorBarColorDark

   iOffset = Abs(CInt(bSelected)) + 1

   If Sunken = False Then
      SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorLt), , PenWidthLt
   Else
      SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorDk), , PenWidthDk
   End If

   If bSeparatorBar Then
      MoveToEx tHdc, X1, Y1 - (CInt(Sunken) + 1), dM
      LineTo tHdc, X2 - 1, Y1 - (CInt(Sunken) + 1)
   Else
      MoveToEx tHdc, X1, Y2, dM
      LineTo tHdc, X1, Y1
      LineTo tHdc, X2, Y1
   End If

   SetMenuColor False, tHdc, cObj_Pen, 0

   If Sunken = False Then
      SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorDk), , PenWidthDk
   Else
      SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorLt), , PenWidthLt
   End If

   If bSeparatorBar = True Then
      MoveToEx tHdc, X1, Y2 - 1 - (CInt(Sunken) + 1), dM
      LineTo tHdc, X2 - 1, Y2 - 1 - (CInt(Sunken) + 1)
   Else
      LineTo tHdc, X2, Y2
      LineTo tHdc, X1, Y2
   End If

   SetMenuColor False, tHdc, cObj_Pen, 0
End Sub

Private Sub DrawSystemIcon(m_hDC As Long, mItem As RECT, ItemID As Long, IsDisabled As Boolean, isSelected As Boolean)
   Dim x As Long
   Dim y As Long
   Dim Xoffset As Long
   Dim Yoffset As Long
   Dim tPt As POINTAPI
   Dim Looper As Integer
   Dim PenType As Integer

   Xoffset = mItem.Left + 6
   Yoffset = mItem.Top + ((mItem.Bottom - mItem.Top) - 8) \ 2 - 1

   Select Case LoWord(ItemID)
      Case SC_CLOSE
         If IsDisabled Then
            PenType = 1
            GoSub GetPen

            MoveToEx m_hDC, Xoffset + 2, Yoffset + 0, tPt
            LineTo m_hDC, Xoffset + 5, Yoffset + 3
            MoveToEx m_hDC, Xoffset + 9, Yoffset + 0, tPt

            For Looper = 1 To 4
               LineTo m_hDC, Xoffset + Choose(Looper, 9, 6, 9, 9), Yoffset + Choose(Looper, 1, 4, 7, 8)
            Next

            MoveToEx m_hDC, Xoffset + 1, Yoffset + 9, tPt
            LineTo m_hDC, Xoffset + 5, Yoffset + 5

            PenType = 0
            GoSub GetPen
         End If

         PenType = 2
         GoSub GetPen

         MoveToEx m_hDC, Xoffset + 0, Yoffset + 0, tPt
         LineTo m_hDC, Xoffset + 9, Yoffset + 9
         MoveToEx m_hDC, Xoffset + 8, Yoffset + 0, tPt
         LineTo m_hDC, Xoffset + -1, Yoffset + 9
         MoveToEx m_hDC, Xoffset + 0, Yoffset + 1, tPt
         LineTo m_hDC, Xoffset + 8, Yoffset + 9
         MoveToEx m_hDC, Xoffset + 1, Yoffset + 0, tPt
         LineTo m_hDC, Xoffset + 9, Yoffset + 8
         MoveToEx m_hDC, Xoffset + 7, Yoffset + 0, tPt
         LineTo m_hDC, Xoffset + -1, Yoffset + 8
         MoveToEx m_hDC, Xoffset + 8, Yoffset + 1, tPt
         LineTo m_hDC, Xoffset + 0, Yoffset + 9

         PenType = 0
         GoSub GetPen

      Case SC_MAXIMIZE
         If IsDisabled Then
            PenType = 1
            GoSub GetPen

            MoveToEx m_hDC, Xoffset + 10, Yoffset + 1, tPt

            For Looper = 1 To 4
               LineTo m_hDC, Xoffset + Choose(Looper, 10, 1, 1, 10), Yoffset + Choose(Looper, 10, 10, 2, 2)
            Next

            PenType = 0
            GoSub GetPen
         End If

         PenType = 2
         GoSub GetPen

         MoveToEx m_hDC, Xoffset + 0, Yoffset + 0, tPt

         For Looper = 1 To 5
            LineTo m_hDC, Xoffset + Choose(Looper, 9, 9, 0, 0, 9), Yoffset + Choose(Looper, 0, 9, 9, 1, 1)
         Next

         PenType = 0
         GoSub GetPen

      Case SC_MINIMIZE
         If IsDisabled Then
            PenType = 1
            GoSub GetPen

            MoveToEx m_hDC, Xoffset + 9, Yoffset + 8, tPt
            LineTo m_hDC, Xoffset + 9, Yoffset + 9
            LineTo m_hDC, Xoffset + 1, Yoffset + 9
            PenType = 0
            GoSub GetPen
         End If

         PenType = 2
         GoSub GetPen

         MoveToEx m_hDC, Xoffset + 1, Yoffset + 8, tPt
         LineTo m_hDC, Xoffset + 9, Yoffset + 8
         MoveToEx m_hDC, Xoffset + 1, Yoffset + 7, tPt
         LineTo m_hDC, Xoffset + 9, Yoffset + 7

         PenType = 0
         GoSub GetPen

      Case SC_RESTORE
         If IsDisabled Then
            PenType = 1
            GoSub GetPen

            MoveToEx m_hDC, Xoffset + 3, Yoffset + 2, tPt
            LineTo m_hDC, Xoffset + 8, Yoffset + 2
            MoveToEx m_hDC, Xoffset + 8, Yoffset + 4, tPt

            For Looper = 1 To 7
               LineTo m_hDC, Xoffset + Choose(Looper, 8, 1, 1, 8, 8, 10, 10), Yoffset + Choose(Looper, 5, 5, 9, 9, 6, 6, 0)
            Next

            PenType = 0
            GoSub GetPen
         End If

         PenType = 2
         GoSub GetPen

         MoveToEx m_hDC, Xoffset + 2, Yoffset + 2, tPt

         For Looper = 1 To 10
            LineTo m_hDC, Xoffset + Choose(Looper, 2, 9, 9, 7, 7, 0, 0, 7, 7, 0), Yoffset + Choose(Looper, 0, 0, 5, 5, 3, 3, 8, 8, 4, 4)
         Next

         MoveToEx m_hDC, Xoffset + 3, Yoffset + 1, tPt
         LineTo m_hDC, Xoffset + 9, Yoffset + 1

         PenType = 0
         GoSub GetPen

      Case Else
         Exit Sub

   End Select

   Exit Sub

GetPen:
   Select Case PenType
      Case 0: SetMenuColor False, m_hDC, cObj_Pen, 0
      Case 1: SetMenuColor True, m_hDC, cObj_Pen, TextColorDisabledLight
      Case 2
         If Not IsDisabled Then
            SetMenuColor True, m_hDC, cObj_Pen, vbBlack 'lSelectBColor
         Else
            SetMenuColor True, m_hDC, cObj_Pen, TextColorDisabledDark
         End If
   End Select

   Return
End Sub

Private Sub DrawCheckMark(isSelected As Boolean, tDC As Long, tRect As RECT, CXoffset As Long, bXPstyle As Boolean, Color1 As Long, Optional Color2 As Long = -1)
   Dim dM As POINTAPI
   Dim Yoffset As Integer
   Dim Xoffset As Integer
   Dim X1 As Integer
   Dim X2 As Integer
   Dim Y1 As Integer
   Dim Y2 As Integer
   Dim Looper As Integer
   Dim Loops As Integer

   If bXPstyle Then
      If isSelected Then
         DrawRect tDC, tRect.Left + 3, tRect.Top + 4, tRect.Left + 16, tRect.Top + 18, GradCol1
         DrawRect tDC, tRect.Left + 4, tRect.Top + 5, tRect.Left + 15, tRect.Top + 17, GradCol2
      Else
         DrawRect tDC, tRect.Left + 3, tRect.Top + 4, tRect.Left + 16, tRect.Top + 18, GradCol2
         DrawRect tDC, tRect.Left + 4, tRect.Top + 5, tRect.Left + 15, tRect.Top + 17, GradCol1
      End If
   End If
   
   Xoffset = CXoffset
   Xoffset = 5 + Xoffset
   Yoffset = ((tRect.Bottom - tRect.Top) - 8) \ 2 + tRect.Top

   If Color2 <> -1 Then Loops = 1

   For Looper = 0 To Loops
      If Looper Then
         SetMenuColor False, tDC, cObj_Pen, 0
         SetMenuColor True, tDC, cObj_Pen, Color2
         Yoffset = Yoffset - 1: Xoffset = Xoffset - 1
      Else
         SetMenuColor True, tDC, cObj_Pen, Color1
      End If

      If bXPstyle Then
         Yoffset = Yoffset + 1
         MoveToEx tDC, Xoffset + 1, Yoffset + 2, dM
         LineTo tDC, Xoffset + 3, Yoffset + 4
         LineTo tDC, Xoffset + 8, Yoffset - 1
         MoveToEx tDC, Xoffset + 1, Yoffset + 3, dM
         LineTo tDC, Xoffset + 3, Yoffset + 5
         LineTo tDC, Xoffset + 8, Yoffset
         Yoffset = Yoffset - 1
      Else
         MoveToEx tDC, 1 + Xoffset, 4 + Yoffset, dM
         LineTo tDC, 2 + Xoffset, 4 + Yoffset
         LineTo tDC, 2 + Xoffset, 5 + Yoffset
         LineTo tDC, 3 + Xoffset, 5 + Yoffset
         LineTo tDC, 3 + Xoffset, 6 + Yoffset
         LineTo tDC, 4 + Xoffset, 6 + Yoffset
         LineTo tDC, 4 + Xoffset, 4 + Yoffset
         LineTo tDC, 5 + Xoffset, 4 + Yoffset
         LineTo tDC, 5 + Xoffset, 2 + Yoffset
         LineTo tDC, 6 + Xoffset, 2 + Yoffset
         LineTo tDC, 6 + Xoffset, 1 + Yoffset
         LineTo tDC, 7 + Xoffset, 1 + Yoffset
         LineTo tDC, 7 + Xoffset, 0 + Yoffset
      End If
   Next

   SetMenuColor False, tDC, cObj_Pen, 0
End Sub

Private Sub SetMenuColor(bSet As Boolean, m_hDC As Long, TypeObject As ColorObjects, lColor As Long, Optional bSamePenColor As Boolean = True, Optional PenWidth As Long = 1)
   Static bObject As Long
   Static pObject As Long
   Static bBMP As Long

   Dim tBrush As Long
   Dim tPen As Long

   If bSet Then
      Select Case TypeObject
         Case cObj_Brush
            tBrush = CreateSolidBrush(ConvertColor(lColor))
            bObject = SelectObject(m_hDC, tBrush)

            If bSamePenColor Then
               tPen = CreatePen(0, PenWidth, ConvertColor(lColor))
               pObject = SelectObject(m_hDC, tPen)
            End If

         Case cObj_Pen
            tPen = CreatePen(0, PenWidth, ConvertColor(lColor))
            pObject = SelectObject(m_hDC, tPen)

         Case cObj_Text
            SetTextColor m_hDC, ConvertColor(lColor)

      End Select
   Else
      Select Case TypeObject
         Case cObj_Brush
           tBrush = SelectObject(m_hDC, bObject)
           DeleteObject tBrush
           bBMP = 0

           If pObject Then
              tPen = SelectObject(m_hDC, pObject)
              DeleteObject tPen
              pObject = 0
           End If
      Case cObj_Pen        ' return original pen & delete one created
          tPen = SelectObject(m_hDC, pObject)
          DeleteObject tPen
          pObject = 0

      End Select
   End If
End Sub

Public Function ConvertColor(tColor As Long) As Long
   If tColor < 0 Then
      ConvertColor = GetSysColor(tColor And &HFF&)
   Else
      ConvertColor = tColor
   End If
End Function

Public Function DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long, HorizontalGrade As Boolean, destDC As Long, dRect As RECT) As Long
   Dim mRect As RECT
   Dim I As Long
   Dim rctOffset As Integer
   Dim DestWidth As Long
   Dim DestHeight As Long
   Dim PixelStep As Long
   Dim XBorder As Long
   Dim Colors() As Long

   On Error Resume Next

   DestWidth = dRect.Right - dRect.Left
   DestHeight = dRect.Bottom - dRect.Top
   mRect = dRect
   rctOffset = 1

   If HorizontalGrade Then
      If (Screen.Width \ Screen.TwipsPerPixelX) \ DestWidth < 5 Then
         PixelStep = Round(DestWidth / 2)
         rctOffset = 2
      Else
         PixelStep = DestWidth
      End If

      ReDim Colors(PixelStep)

      LoadColors Colors(), Color1, Color2
      mRect.Right = rctOffset + dRect.Left

      For I = 0 To PixelStep - 1
         DrawRect destDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(I)
         OffsetRect mRect, rctOffset, 0
      Next
   Else
      If (Screen.Height \ Screen.TwipsPerPixelY) \ DestHeight < 5 Then
         PixelStep = Round(DestHeight / 2)
         rctOffset = 2
      Else
         PixelStep = DestHeight
      End If

      ReDim Colors(PixelStep) As Long

      LoadColors Colors(), Color2, Color1
      mRect.Bottom = rctOffset + dRect.Top

      For I = 0 To PixelStep - 1
         DrawRect destDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(I)
         OffsetRect mRect, 0, rctOffset
      Next
   End If
End Function

Public Sub LoadColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
   Dim I As Long
   Dim BaseR As Single
   Dim BaseG As Single
   Dim BaseB As Single
   Dim PlusR As Single
   Dim PlusG As Single
   Dim PlusB As Single
   Dim MinusR As Single
   Dim MinusG As Single
   Dim MinusB As Single

   BaseR = CSng(Color1 And &HFF)
   BaseG = CSng(Color1 And &HFF00&) / 255
   BaseB = CSng(Color1 And &HFF0000) / &HFF00&
   
   MinusR = CSng(Color2 And &HFF&)
   MinusG = CSng(Color2 And &HFF00&) / 255
   MinusB = CSng(Color2 And &HFF0000) / &HFF00&
   
   PlusR = (MinusR - BaseR) / UBound(Colors)
   PlusG = (MinusG - BaseG) / UBound(Colors)
   PlusB = (MinusB - BaseB) / UBound(Colors)

   For I = 0 To UBound(Colors)
      BaseR = BaseR + PlusR
      BaseG = BaseG + PlusG
      BaseB = BaseB + PlusB
      
      If BaseR > 255 Then BaseR = 255
      If BaseG > 255 Then BaseG = 255
      If BaseB > 255 Then BaseB = 255
      
      If BaseR < 0 Then BaseR = 0
      If BaseG < 0 Then BaseG = 0
      If BaseG < 0 Then BaseB = 0
      
      Colors(I) = RGB(BaseR, BaseG, BaseB)
   Next
End Sub

Public Function ExchangeVBcolor(vValue As Variant, DefaultColor As Long) As Long
   If IsNumeric(vValue) Then
      ExchangeVBcolor = ConvertColor(CLng(vValue))
   Else
      Select Case CStr(vValue)
         Case "vbWhite": ExchangeVBcolor = vbWhite
         Case "vbBlack": ExchangeVBcolor = vbBlack
         Case "vbBlue": ExchangeVBcolor = vbBlue
         Case "vbGreen": ExchangeVBcolor = vbGreen
         Case "vbRed": ExchangeVBcolor = vbRed
         Case "vbMagenta": ExchangeVBcolor = vbMagenta
         Case "vbYellow": ExchangeVBcolor = vbYellow
         Case "vbCyan": ExchangeVBcolor = vbCyan
         Case "vbMaroon": ExchangeVBcolor = vbMaroon
         Case "vbOlive": ExchangeVBcolor = vbOlive
         Case "vbNavy": ExchangeVBcolor = vbNavy
         Case "vbPurple": ExchangeVBcolor = vbPurple
         Case "vbTeal": ExchangeVBcolor = vbTeal
         Case "vbGray": ExchangeVBcolor = vbGray
         Case "vbSilver": ExchangeVBcolor = vbSilver
         Case "vbViolet": ExchangeVBcolor = vbViolet
         Case "vbOrange": ExchangeVBcolor = vbOrange
         Case "vbGold": ExchangeVBcolor = vbGold
         Case "vbIvory": ExchangeVBcolor = vbIvory
         Case "vbPeach": ExchangeVBcolor = vbPeach
         Case "vbTurquoise": ExchangeVBcolor = vbTurquoise
         Case "vbTan": ExchangeVBcolor = vbTan
         Case "vbBrown": ExchangeVBcolor = vbBrown
         Case "vbScrollBars": ExchangeVBcolor = ConvertColor(vbScrollBars)
         Case "vbDesktop": ExchangeVBcolor = ConvertColor(vbDesktop)
         Case "vbActiveTitleBar": ExchangeVBcolor = ConvertColor(vbActiveTitleBar)
         Case "vbInactiveTitleBar": ExchangeVBcolor = ConvertColor(vbInactiveTitleBar)
         Case "vbMenuBar": ExchangeVBcolor = ConvertColor(vbMenuBar)
         Case "vbWindowBackground": ExchangeVBcolor = ConvertColor(vbWindowBackground)
         Case "vbWindowFrame": ExchangeVBcolor = ConvertColor(vbWindowFrame)
         Case "vbMenuText": ExchangeVBcolor = ConvertColor(vbMenuText)
         Case "vbWindowText": ExchangeVBcolor = ConvertColor(vbWindowText)
         Case "vbTitleBarText": ExchangeVBcolor = ConvertColor(vbTitleBarText)
         Case "vbActiveBorder": ExchangeVBcolor = ConvertColor(vbActiveBorder)
         Case "vbInactiveBorder": ExchangeVBcolor = ConvertColor(vbInactiveBorder)
         Case "vbApplicationWorkspace": ExchangeVBcolor = ConvertColor(vbApplicationWorkspace)
         Case "vbHighlight": ExchangeVBcolor = ConvertColor(vbHighlight)
         Case "vbHighlightText": ExchangeVBcolor = ConvertColor(vbHighlightText)
         Case "vbButtonFace": ExchangeVBcolor = ConvertColor(vbButtonFace)
         Case "vbButtonShadow": ExchangeVBcolor = ConvertColor(vbButtonShadow)
         Case "vbGrayText": ExchangeVBcolor = ConvertColor(vbGrayText)
         Case "vbButtonText": ExchangeVBcolor = ConvertColor(vbButtonText)
         Case "vbInactiveCaptionText": ExchangeVBcolor = ConvertColor(vbInactiveCaptionText)
         Case "vb3DHighlight": ExchangeVBcolor = ConvertColor(vb3DHighlight)
         Case "vb3DDKShadow": ExchangeVBcolor = ConvertColor(vb3DDKShadow)
         Case "vb3DLight": ExchangeVBcolor = ConvertColor(vb3DLight)
         Case "vb3DFace": ExchangeVBcolor = ConvertColor(vb3DFace)
         Case "vb3Dshadow": ExchangeVBcolor = ConvertColor(vb3DShadow)
         Case "vbInfoText": ExchangeVBcolor = ConvertColor(vbInfoText)
         Case "vbInfoBackground": ExchangeVBcolor = ConvertColor(vbInfoBackground)
         Case Else
            If Left$(vValue, 2) = "&H" Then
                ExchangeVBcolor = ConvertColor(Val(vValue))
            Else
                ExchangeVBcolor = DefaultColor
            End If

      End Select
   End If
End Function

Public Sub LoadFontMenu(vFontArray As Variant, Optional FontType As Long)
   ReDim vFonts(0 To 0)

   Dim hdc As Long

   hdc = GetDC(CLng(Mid(hWndRedirect, 2)))
   EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, FontType
   ReleaseDC CLng(Mid(hWndRedirect, 2)), hdc

   If UBound(vFonts) Then ShellSort vFonts

   vFontArray = vFonts
   Erase vFonts
End Sub

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As Long) As Long
   Dim FaceName As String
   Dim bInclude As Boolean

   EnumFontFamProc = 1

   Select Case lParam
      Case RASTER_FONTTYPE
         bInclude = ((FontType = RASTER_FONTTYPE) Or (FontType = 0)) ' = RASTER_FONTTYPE Or FontType = 0)

      Case TRUETYPE_FONTTYPE
         bInclude = (FontType = TRUETYPE_FONTTYPE) ' = TRUETYPE_FONTTYPE)

      Case Else
         bInclude = True

   End Select

   If bInclude Then
      FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)

      ReDim Preserve vFonts(0 To UBound(vFonts) + 1)

      vFonts(UBound(vFonts)) = StringFromBuffer(FaceName)
   End If
End Function

Private Sub ShellSort(vArray As Variant)
   Dim lLoop1 As Long
   Dim lHold As Long
   Dim lHValue As Long
   Dim lTemp As Variant

   lHValue = LBound(vArray)

   Do
      lHValue = 3 * lHValue + 1
   Loop Until lHValue > UBound(vArray)

   Do
      lHValue = lHValue / 3

      For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
         lTemp = vArray(lLoop1)
         lHold = lLoop1

         Do While vArray(lHold - lHValue) > lTemp
            vArray(lHold) = vArray(lHold - lHValue)
            lHold = lHold - lHValue
            If lHold < lHValue Then Exit Do
         Loop

         vArray(lHold) = lTemp
      Next lLoop1
   Loop Until lHValue = LBound(vArray)
End Sub

Public Function StripFile(Pathname As String, DPNEm As String) As String
   Dim ChrsIn As String
   Dim Chrs As Integer

   On Error GoTo StripFile_General_ErrTrap

   If Pathname = "" Then Exit Function

   ChrsIn = Pathname

   Select Case InStr("DPNEm", DPNEm)
      Case 1
         Chrs = InStr(ChrsIn, ":")

         If Chrs Then
            StripFile = Left(ChrsIn, Chrs) & "\"
         Else
            Chrs = InStr(ChrsIn, "\\")

            If Chrs = 1 Then
               Chrs = InStr(Chrs + 2, ChrsIn, "\")

               If Chrs Then StripFile = Left$(ChrsIn, Chrs) Else StripFile = ChrsIn & "\"
            End If
         End If

      Case 2
         Chrs = InStrRev(ChrsIn, "\")

         If Chrs = 0 Then Chrs = InStr(ChrsIn, ":") Else Chrs = Chrs - 1
         If Chrs Then StripFile = Left$(ChrsIn, Chrs) & "\"

      Case 3
         Chrs = InStrRev(ChrsIn, "\")

         If Chrs = 0 Then Chrs = InStr(ChrsIn, ":")
         If Chrs Then StripFile = Mid$(ChrsIn, Chrs + 1)

      Case 4
         Chrs = InStrRev(ChrsIn, ".")

         If Chrs Then
            If InStr(Chrs, "\") = 0 Then StripFile = Mid$(ChrsIn, Chrs + 1)
         End If

      Case 5
         Chrs = InStrRev(ChrsIn, "\")

         If Chrs = 0 Then Chrs = InStr(ChrsIn, ":")
         If Chrs Then ChrsIn = Mid$(ChrsIn, Chrs + 1)

         Chrs = InStrRev(ChrsIn, ".")

         If Chrs > 1 Then ChrsIn = Left$(ChrsIn, Chrs - 1) Else Chrs = 0
         If Chrs Then StripFile = ChrsIn

   End Select

   Exit Function

StripFile_General_ErrTrap:
   MsgBox "Err: " & Err.Number & " - Procedure: StripFile" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Function GetFloppyIcon(sDrive As String) As Long
   If FloppyIcon = 0 Then
      FloppyIcon = ExtractAssociatedIcon(App.hInstance, sDrive, 0)
   End If

   GetFloppyIcon = FloppyIcon
End Function

Public Function StringFromBuffer(Buffer As String) As String
   Dim nPos As Long

   nPos = InStr(Buffer, vbNullChar)

   If nPos > 0 Then
      StringFromBuffer = Left$(Buffer, nPos - 1)
   Else
      StringFromBuffer = Buffer
   End If
End Function

Private Function GetToolTipWindow(pHwnd As Long) As Long
   Dim I As Integer
   Dim tHwnd As Long

   On Error GoTo ReDimensionArray

   I = UBound(tbarClass)

   For I = 0 To UBound(tbarClass)
      tHwnd = FindWindowEx(pHwnd, 0, tbarClass(I), vbNullString)

      If tHwnd Then Exit For
   Next

   GetToolTipWindow = tHwnd
   Exit Function

ReDimensionArray:
   ReDim tbarClass(0)

   tbarClass(0) = "msvb_lib_toolbar"
   Resume
End Function

Public Sub AddToolbarClass(sClass As String)
   Dim I As Integer

   On Error GoTo ReDimensionArray

   I = UBound(tbarClass)

   For I = 0 To UBound(tbarClass)
      If tbarClass(I) = sClass Then Exit Sub
   Next

   ReDim Preserve tbarClass(0 To UBound(tbarClass))

   tbarClass(UBound(tbarClass)) = sClass
   Exit Sub

ReDimensionArray:
   ReDim tbarClass(0)

   tbarClass(0) = "msvb_lib_toolbar"
   Resume
End Sub

Private Sub IDCurrentWindow(hwnd As Long, MenuID As Long, bSystemMenu As Boolean)
   If bSystemMenu Then
      hWndRedirect = "h" & hwnd
      MenuID = colMenuItems("h" & hwnd).SystemMenu
      Exit Sub
   End If

   hWndRedirect = "h" & hwnd

   Dim hMDI As Long
   Dim hParent As Long

   If colMenuItems("h" & hwnd).IsMDIclient Then
      hMDI = SendMessage(hwnd, WM_MDIGETACTIVE, 0&, ByVal 0&)
      hParent = GetParent(hwnd)
   Else
      If colMenuItems("h" & hwnd).MDIClient Then
         hParent = hwnd
         hMDI = SendMessage(colMenuItems("h" & hwnd).MDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
      End If
   End If

   If hMDI Then
      If tempRedirect Then hWndRedirect = "h" & tempRedirect Else hWndRedirect = "h" & hMDI

      If colMenuItems("h" & hMDI).IsMenuLess Then
         If colMenuItems("h" & hMDI).SystemMenu <> MenuID Then
            If tempRedirect = 0 Then hWndRedirect = "h" & hParent
         End If
      End If
   Else
      If tempRedirect Then hWndRedirect = "h" & tempRedirect
   End If
End Sub

Private Function IdentifyAccelerator(KeyCode As Long, hMenu As Long) As Long
   Dim Index As Integer
   Dim I As Integer

   colMenuItems(hWndRedirect).GetPanelItem hMenu

   With XferPanelData
      Index = InStr(.Accelerators, UCase(Chr$(KeyCode)))

      If Index = 0 Then
         IdentifyAccelerator = MakeLong(0, MNC_IGNORE)
      Else
         If InStrRev(.Accelerators, UCase(Chr$(KeyCode))) = Index Then
            IdentifyAccelerator = MakeLong(Index - 1, MNC_EXECUTE)
         Else
            Dim vIndex() As Integer
            Dim MI() As Byte
            Dim MII As MENUITEMINFO

            ReDim vIndex(0)

            Do While Index > 0
               ReDim Preserve vIndex(0 To UBound(vIndex) + 1)

               vIndex(UBound(vIndex)) = Index
               Index = InStr(Index + 1, .Accelerators, UCase(Chr$(KeyCode)))
            Loop

            For I = 1 To UBound(vIndex)
               ReDim MI(0 To 1023)

               MII.cbSize = Len(MII)
               MII.fType = 0
               MII.fMask = MIIM_STATE
               MII.cch = UBound(MI)
               GetMenuItemInfo hMenu, vIndex(I) - 1, True, MII

               If ((MII.fState And MF_HILITE) = MF_HILITE) Then
                  Index = I
                  Exit For
               End If
            Next

            Erase MI

            If Index Then
               If Index = UBound(vIndex) Then Index = 1 Else Index = Index + 1
            Else
               Index = 1
            End If

            IdentifyAccelerator = MakeLong(vIndex(Index) - 1, MNC_SELECT)

            Erase vIndex
         End If
      End If
   End With
End Function

Private Sub ProcessKeyStroke(hwnd As Long, KeyStroke As Long, bKeyUp As Boolean)
   Dim hMDI As Long
   Dim hParent As Long
   Dim pClass As Long
   Dim ShiftStatus As Long

   On Error GoTo FailedKeyRepeater

   If colMenuItems("h" & hwnd).IsMDIclient Then
      hMDI = SendMessage(hwnd, WM_MDIGETACTIVE, 0&, ByVal 0&)
      hParent = GetParent(hwnd)
   Else
      If colMenuItems("h" & hwnd).MDIClient Then
         hParent = hwnd
         hMDI = SendMessage(colMenuItems("h" & hwnd).MDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
      End If
   End If

   If hParent <> 0 And hMDI = 0 Then
      pClass = colMenuItems("h" & hParent).ShowTips

      If pClass = 0 Then Exit Sub
      If (GetKeyState(VK_SHIFT) And &HF0000000) Then ShiftStatus = ShiftStatus Or 1
      If (GetKeyState(VK_CONTROL) And &HF0000000) Then ShiftStatus = ShiftStatus Or 2
      If (GetKeyState(VK_MENU) And &HF0000000) Then ShiftStatus = ShiftStatus Or 4

      Dim oTipClass As clsTips

      CopyMemory oTipClass, pClass, 4&
      oTipClass.SendMDIKeyPress KeyStroke, ShiftStatus, bKeyUp
      CopyMemory oTipClass, 0&, 4&

      Set oTipClass = Nothing
   End If

FailedKeyRepeater:
End Sub

Public Function SetMinMaxInfo(hwnd As Long, MaximizedW As Long, MaximizedH As Long, MaximizedLeft As Long, MaximizedTop As Long, _
   MaxDragSizeW As Long, MaxDragSizeH As Long, MinDragSizeW As Long, MinDragSizeH As Long)

   On Error Resume Next

   If colMenuItems("h" & hwnd).hPrevProc = 0 Then Exit Function

   Dim uMinMax As MINMAXINFO
   Dim sRect As RECT

   SystemParametersInfo SPI_GETWORKAREA, 0, sRect, 0

   If MaximizedW < 0 Then MaximizedW = sRect.Right - sRect.Left
   If MaximizedH < 0 Then MaximizedH = sRect.Bottom - sRect.Top + 87

   uMinMax.ptMaxSize.x = MaximizedW
   uMinMax.ptMaxSize.y = MaximizedH

   If MaximizedLeft < 0 Then MaximizedLeft = sRect.Left
   If MaximizedTop < 0 Then MaximizedTop = sRect.Top

   uMinMax.ptMaxPosition.x = MaximizedLeft
   uMinMax.ptMaxPosition.y = MaximizedTop

   If MaxDragSizeW < 0 Then MaxDragSizeW = sRect.Right - sRect.Left
   If MaxDragSizeH < 0 Then MaxDragSizeH = sRect.Bottom - sRect.Top + 87

   uMinMax.ptMaxTrackSize.x = MaxDragSizeW
   uMinMax.ptMaxTrackSize.y = MaxDragSizeH

   If MinDragSizeW < 0 Then MinDragSizeW = 0
   If MinDragSizeH < 0 Then MinDragSizeH = 0

   uMinMax.ptMinTrackSize.x = MinDragSizeW
   uMinMax.ptMinTrackSize.y = MinDragSizeH

   colMenuItems("h" & hwnd).RestrictSize VarPtr(uMinMax), True
End Function

Private Sub LoadDefaultColors()
   bModuleInitialized = True
   lSelectBColor = GetSysColor(COLOR_HIGHLIGHT)
   TextColorNormal = GetSysColor(COLOR_MENUTEXT)
   TextColorSelected = GetSysColor(COLOR_HIGHLIGHTTEXT)
   TextColorDisabledDark = GetSysColor(COLOR_GRAYTEXT)
   TextColorDisabledLight = TextColorSelected ' GetSysColor(COLOR_HIGHLIGHTTEXT)
   TextColorSeparatorBar = lSelectBColor
   SeparatorBarColorDark = GetSysColor(COLOR_BTNSHADOW)
   SeparatorBarColorLight = GetSysColor(COLOR_BTNHIGHLIGHT)
   CheckedIconBColor = GetSysColor(COLOR_BTNLIGHT)
End Sub
