VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event DisplayTip(TipText As String)
Public Event CustomSelection(UserID As String, Category As String, Value As Variant)
Public Event MDIKeyDown(KeyCode As Integer, Shift As Integer)
Public Event MDIKeyUp(KeyCode As Integer, Shift As Integer)

Friend Sub SendTip(sTip As String)
   On Error Resume Next
   RaiseEvent DisplayTip(sTip)
   On Error GoTo 0
End Sub

Friend Sub SendCustomSelection(sID As String, sCat As String, NewVal As Variant)
   On Error Resume Next
   RaiseEvent CustomSelection(sID, sCat, NewVal)
   On Error GoTo 0
End Sub

Friend Sub SendMDIKeyPress(ByVal KeyCodex As Long, ByVal ShiftCodex As Long, ByVal KeyUp As Boolean)
   If KeyUp Then
      RaiseEvent MDIKeyUp(CInt(KeyCodex), CInt(ShiftCodex))
   Else
      RaiseEvent MDIKeyDown(CInt(KeyCodex), CInt(ShiftCodex))
   End If
End Sub

