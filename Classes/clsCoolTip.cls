VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const WM_USER = &H400

Private Const CW_USEDEFAULT = &H80000000

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2

Private Const HWND_TOPMOST = -1

Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOFADE As Long = &H20
Private Const TTS_NOANIMATE As Long = &H10
Private Const TTS_NOPREFIX = &H2

Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTM_HITTESTA As Long = (WM_USER + 10)
Private Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Private Const TTM_POPUP As Long = (WM_USER + 34)
Private Const TTM_POP As Long = (WM_USER + 28)

Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK As Long = &H20
Private Const TTF_ABSOLUTE As Long = &H80

Private Const WS_BORDER As Long = &H800000

Private Const TTDT_RESHOW As Long = 1
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TOOLINFO
   lSize As Long
   lFlags As Long
   lhWnd As Long
   lId As Long
   lpRect As RECT
   hInstance As Long
   lpStr As String
   lParam As Long
End Type

Public Enum IconMode
   [1-No Icon]
   [2-Information Icon]
   [3-Exclamation Icon]
   [4-Critical Icon]
End Enum

Public Enum TipStyle
   [Standard Tip]
   [Balloon Tip]
End Enum

Private m_BackColor As Long
Private m_TitleText As String
Private m_TipText As String
Private m_ForeColor As Long
Private m_ParenthWnd As Long
Private m_Icon As IconMode
Private m_Centered As Boolean
Private m_Style As TipStyle
Private m_ForeverTip As Boolean

Private lhWnd As Long
Private eTI As TOOLINFO
Private lpRect As RECT

Public Property Let ForeverTip(ByVal vData As Boolean)
   m_ForeverTip = vData
End Property

Public Property Get ForeverTip() As Boolean
   ForeverTip = m_ForeverTip
End Property

Public Property Let Style(ByVal vData As TipStyle)
   m_Style = vData
End Property

Public Property Get Style() As TipStyle
   Style = m_Style
End Property

Public Property Let Centered(ByVal vData As Boolean)
   m_Centered = vData
End Property

Public Property Get Centered() As Boolean
   Centered = m_Centered
End Property

Public Sub Destroy()
   If ForeverTip Then
      UnSubClassTip lhWnd
   End If
   
   m_ParenthWnd = 0
   m_TipText = vbNullString
   m_TitleText = vbNullString
   
   If lhWnd <> 0 Then
      DestroyWindow lhWnd
   End If
End Sub

Public Sub Create()
   Dim lWinStyle As Long
   
   If lhWnd <> 0 Then
      DestroyWindow lhWnd
   End If

   lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
   If m_Style = [Balloon Tip] Then lWinStyle = lWinStyle Or TTS_BALLOON
    
   If m_ParenthWnd <> 0 Then
      lhWnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, "Tooltip", lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, m_ParenthWnd, 0&, App.hInstance, 0&)
      GetClientRect m_ParenthWnd, lpRect
      
      With eTI
         If m_Centered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
         Else
            .lFlags = TTF_SUBCLASS
         End If
         
         .lhWnd = m_ParenthWnd
         .lId = 0
         .hInstance = App.hInstance
         .lpRect = lpRect
      End With
      
      If m_TitleText <> vbNullString Or m_Icon <> [1-No Icon] Then
         SendMessage lhWnd, TTM_SETTITLE, CLng(m_Icon), ByVal m_TitleText
      End If
      
      If m_ForeColor <> Empty Then
         SendMessage lhWnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, 0&
      End If
      
      If m_BackColor <> Empty Then
         SendMessage lhWnd, TTM_SETTIPBKCOLOR, m_BackColor, 0&
      End If
      
      If ForeverTip Then SubClassTip lhWnd
      SendMessage lhWnd, TTM_ADDTOOLA, 0&, eTI
   End If
End Sub

Public Property Let ParenthWnd(ByVal vData As Long)
   m_ParenthWnd = vData
End Property

Public Property Get ParenthWnd() As Long
   ParenthWnd = m_ParenthWnd
End Property

Public Property Let Icon(ByVal vData As IconMode)
   m_Icon = vData
   
   If lhWnd <> 0 And m_TitleText <> Empty And m_Icon <> [1-No Icon] Then
      SendMessage lhWnd, TTM_SETTITLE, CLng(m_Icon), ByVal m_TitleText
   End If
End Property

Public Property Get Icon() As IconMode
   Icon = m_Icon
End Property

Public Property Let ForeColor(ByVal vData As Long)
   m_ForeColor = vData
   If lhWnd <> 0 Then SendMessage lhWnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, 0&
End Property

Public Property Get ForeColor() As Long
   ForeColor = m_ForeColor
End Property

Public Property Let TitleText(ByVal vData As String)
   m_TitleText = vData
   
   If lhWnd <> 0 And m_TitleText <> Empty And m_Icon <> [1-No Icon] Then
      SendMessage lhWnd, TTM_SETTITLE, CLng(m_Icon), ByVal m_TitleText
   End If
End Property

Public Property Get TitleText() As String
   TitleText = eTI.lpStr
End Property

Public Property Let BackColor(ByVal vData As Long)
   m_BackColor = vData
   If lhWnd <> 0 Then SendMessage lhWnd, TTM_SETTIPBKCOLOR, m_BackColor, 0&
End Property

Public Property Get BackColor() As Long
   BackColor = m_BackColor
End Property

Public Property Let TipText(ByVal vData As String)
   eTI.lpStr = vData
   m_TipText = vData
   If lhWnd <> 0 Then SendMessage lhWnd, TTM_UPDATETIPTEXTA, 0&, eTI
End Property

Public Property Get TipText() As String
    TipText = m_TipText
End Property

Private Sub Class_Terminate()
   If lhWnd <> 0 Then DestroyWindow lhWnd
End Sub
