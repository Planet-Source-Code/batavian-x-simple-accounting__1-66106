VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarvLeft As Single 'local copy
Private mvarvPage As Integer 'local copy
Private mvarvTop As Single 'local copy
Private mvarvHeight As Single 'local copy
Private mvarvWidth As Single 'local copy

'local variable(s) to hold property value(s)
Private mvarvPicture As PictureBox 'local copy
Private mvarvScaleMode As Integer 'local copy
Private mvarvCenter As Boolean 'local copy
Private mvarvPrinted As Boolean 'local copy

Friend Property Let vCenter(ByVal vData As Boolean)
    mvarvCenter = vData
End Property

Friend Property Get vCenter() As Boolean
    vCenter = mvarvCenter
End Property

Friend Property Let vScaleMode(ByVal vData As qePrinterScale)
    mvarvScaleMode = vData
End Property

Friend Property Get vScaleMode() As qePrinterScale
    vScaleMode = mvarvScaleMode
End Property

Friend Property Set vPicture(ByVal vData As PictureBox)
    Set mvarvPicture = vData
End Property

Friend Property Get vPicture() As PictureBox
    Set vPicture = mvarvPicture
End Property

Friend Property Let vWidth(ByVal vData As Single)
    mvarvWidth = vData
End Property

Friend Property Get vWidth() As Single
    vWidth = mvarvWidth
End Property

Friend Property Let vHeight(ByVal vData As Single)
    mvarvHeight = vData
End Property

Friend Property Get vHeight() As Single
    vHeight = mvarvHeight
End Property

Friend Property Let vTop(ByVal vData As Single)
    mvarvTop = vData
End Property

Friend Property Get vTop() As Single
    vTop = mvarvTop
End Property

Friend Property Let vPage(ByVal vData As Integer)
    mvarvPage = vData
End Property

Friend Property Get vPage() As Integer
    vPage = mvarvPage
End Property

Friend Property Let vLeft(ByVal vData As Single)
    mvarvLeft = vData
End Property

Friend Property Get vLeft() As Single
    vLeft = mvarvLeft
End Property

Friend Property Let vPrinted(ByVal vData As Boolean)
   mvarvPrinted = vData
End Property

Friend Property Get vPrinted() As Boolean
   vPrinted = mvarvPrinted
End Property


