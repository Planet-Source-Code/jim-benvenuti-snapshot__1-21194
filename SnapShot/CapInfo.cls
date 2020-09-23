VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This class is used in frmShots. It is just a container to
' store information about a SnapShot.

Option Explicit

Private mCapWidth As Long
Private mCapHeight As Long
Private mCapX As Long
Private mCapY As Long
Private mCapFileName As String

Public Property Let CapFileName(ByVal vData As String)
    
    mCapFileName = vData

End Property

Public Property Get CapFileName() As String
    
    CapFileName = mCapFileName

End Property

Public Property Let CapY(ByVal vData As Long)
    
    mCapY = vData

End Property

Public Property Get CapY() As Long
    
    CapY = mCapY

End Property

Public Property Let CapX(ByVal vData As Long)
    
    mCapX = vData

End Property

Public Property Get CapX() As Long
    
    CapX = mCapX

End Property

Public Property Let CapHeight(ByVal vData As Long)
    
    mCapHeight = vData

End Property

Public Property Get CapHeight() As Long
    
    CapHeight = mCapHeight

End Property

Public Property Let CapWidth(ByVal vData As Long)
    
    mCapWidth = vData

End Property

Public Property Get CapWidth() As Long
    
    CapWidth = mCapWidth

End Property



