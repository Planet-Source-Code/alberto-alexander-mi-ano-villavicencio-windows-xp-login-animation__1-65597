VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_cMemDC As clsAlphaDibSection

Private m_bAutoSize As Boolean
Private m_Picture As Picture
Private m_lngTop As Long
Private m_lngLeft As Long
Private m_lHeight As Long
Private m_lWidth As Long

Private m_lngPosX As Long
Private m_lngPosY As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'
Public Property Get AutoSize() As Boolean
  AutoSize = m_bAutoSize
End Property
Public Property Let AutoSize(ByVal rhs As Boolean)
  m_bAutoSize = rhs
  Refresh
End Property

'
Public Property Set Picture(ByVal rhs As StdPicture)
  If Not rhs Is Nothing Then
    Set m_Picture = rhs
  
    Set m_cMemDC = New clsAlphaDibSection
    With m_cMemDC
      .CreateFromPicture m_Picture
  
      m_lWidth = .Width
      m_lHeight = .Height
    End With
  Else
    Set m_cMemDC = Nothing
    Set m_Picture = Nothing
  End If
End Property
Public Property Get Picture() As StdPicture
  Set Picture = m_Picture
End Property

'
Public Property Let Left(ByVal vData As Long)
  m_lngLeft = vData
End Property
Public Property Get Left() As Long
  Left = m_lngLeft
  Refresh
End Property

'
Public Property Let Top(ByVal vData As Long)
  m_lngTop = vData
End Property
Public Property Get Top() As Long
  Top = m_lngTop
  Refresh
End Property

'
Public Property Get Width() As Long
  Width = m_lWidth
End Property
Public Property Let Width(ByVal rhs As Long)
  If (m_bAutoSize) Then
    Err.Raise 425
  Else
    m_lWidth = rhs
    Refresh
  End If
End Property

'
Public Property Get Height() As Long
  Height = m_lHeight
End Property
Public Property Let Height(ByVal rhs As Long)
  m_lHeight = rhs
  Refresh
End Property

Private Sub Class_Initialize()
  m_bAutoSize = True
  m_lngLeft = 0
  m_lngTop = 0
  m_lWidth = 100
  m_lHeight = 20
End Sub

Public Function Refresh()
  ' Determine the width if we are autosizing.
  If m_bAutoSize Then
   
  Else
    
  End If
End Function

Public Sub Draw(ByVal lHDC As Long, Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight As Long = -1)
  If m_cMemDC Is Nothing Then Exit Sub
  m_cMemDC.PaintPicture lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, vbSrcCopy
End Sub

