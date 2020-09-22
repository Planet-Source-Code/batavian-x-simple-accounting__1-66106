VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicItemCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPicItem"
Attribute VB_Ext_KEY = "Member0" ,"clsPicItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Private mCol As Collection

Friend Function Add(qPicItem As clsPicItem, Optional Key As String) As Boolean
   If Len(Key) = 0 Then mCol.Add qPicItem Else mCol.Add qPicItem, Key
End Function

Public Property Get Item(vntIndexKey As Variant) As clsPicItem
   Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub

