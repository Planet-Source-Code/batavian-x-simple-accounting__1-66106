VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarX1 As Single 'local copy
Private mvarX2 As Single 'local copy
Private mvarY1 As Single 'local copy
Private mvarY2 As Single 'local copy
Private mvarLineWidth As Integer 'local copy
Private mvarPage As Integer 'local copy
Private mvarvPrinted As Boolean 'local copy

Friend Property Let vPrinted(ByVal vData As Boolean)
    mvarvPrinted = vData
End Property

Friend Property Get vPrinted() As Boolean
    vPrinted = mvarvPrinted
End Property

Friend Property Let LineWidth(ByVal vData As Integer)
    mvarLineWidth = vData
End Property

Friend Property Get LineWidth() As Integer
    LineWidth = mvarLineWidth
End Property

Friend Property Let Page(ByVal vData As Integer)
    mvarPage = vData
End Property

Friend Property Get Page() As Integer
    Page = mvarPage
End Property

Friend Property Let Y2(ByVal vData As Single)
    mvarY2 = vData
End Property

Friend Property Get Y2() As Single
    Y2 = mvarY2
End Property

Friend Property Let Y1(ByVal vData As Single)
    mvarY1 = vData
End Property

Friend Property Get Y1() As Single
    Y1 = mvarY1
End Property

Friend Property Let X2(ByVal vData As Single)
    mvarX2 = vData
End Property

Friend Property Get X2() As Single
    X2 = mvarX2
End Property

Friend Property Let X1(ByVal vData As Single)
    mvarX1 = vData
End Property

Friend Property Get X1() As Single
    X1 = mvarX1
End Property



