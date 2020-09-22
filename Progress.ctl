VERSION 5.00
Begin VB.UserControl Progress 
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ToolboxBitmap   =   "Progress.ctx":0000
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------'
' Copyright © 2006 AMV® Solutions.  All rights reserved.                  '
'                                                                         '
' Programer by: Alberto A. Miñano Villavicencio.                          '
' Date Created: 04/02/2003                                                '
' Last Updated: 23/02/2005                                                '
'-------------------------------------------------------------------------'
Option Explicit

Public Enum AdvanceConst
  acLeftToRight
  acRightToLeft
  acBiDirectional
End Enum

'Default Property Values:
Const m_def_Advance = 0
Const m_def_Speed = 4

Private objProgressBar As clsProgress
Private WithEvents objTimer As clsTimerPlus
Attribute objTimer.VB_VarHelpID = -1

'Property Variables:
Private m_Advance As AdvanceConst
Private m_AutoSize As Boolean
Private m_Picture As Picture
Private m_Speed As Long


'
Public Property Get Advance() As AdvanceConst
  Advance = m_Advance
End Property

Public Property Let Advance(ByVal New_Advance As AdvanceConst)
  m_Advance = New_Advance
  PropertyChanged "Advance"
End Property

'
Public Property Get AutoSize() As Boolean
  AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal rhs As Boolean)
  m_AutoSize = rhs
  UserControl_Resize
End Property

'
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
  Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
  If Not New_Picture Is Nothing Then
    Set m_Picture = New_Picture
    Set objProgressBar.Picture = m_Picture
  
    With UserControl
      .Height = objProgressBar.Height * Screen.TwipsPerPixelY
      .Width = objProgressBar.Width * Screen.TwipsPerPixelX
    End With
  
    objProgressBar.Draw UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Else
    Set m_Picture = Nothing
  End If
  
  PropertyChanged "Picture"
End Property

'
Public Property Get Speed() As Long
  Speed = m_Speed
End Property
Public Property Let Speed(ByVal New_Speed As Long)
  m_Speed = New_Speed
  PropertyChanged "Speed"
End Property

'
Public Function StartAnimation() As Variant
  Set objTimer = New clsTimerPlus
  objTimer.Interval = 15
End Function

'
Public Function StopAnimation() As Variant
  If Not objTimer Is Nothing Then objTimer.Interval = 0
End Function

'
Public Function EndAnimation() As Variant
  If Not objTimer Is Nothing Then
    objTimer.Interval = 0
    Set objTimer = Nothing
  End If
  
  UserControl_Paint
End Function

Private Sub objTimer_Tick()
  Static lngPosX As Long
  Static lngPosY As Long
  Static blnShiftLeft As Boolean
  
  
  With UserControl
    Select Case m_Advance
      Case acLeftToRight: GoTo LeftToRight
      Case acRightToLeft: GoTo RightToLeft
      Case acBiDirectional: If Not blnShiftLeft Then GoTo LeftToRight Else GoTo RightToLeft
    End Select
    Exit Sub
    
    
LeftToRight:
  lngPosY = (.ScaleWidth - ((.ScaleWidth * 2) + 1)) + lngPosX
        
  objProgressBar.Draw .hdc, lngPosY, 0, .ScaleWidth, .ScaleHeight
  objProgressBar.Draw .hdc, lngPosX, 0, .ScaleWidth, .ScaleHeight
        
  If lngPosX >= .ScaleWidth Then
    lngPosX = 0
    blnShiftLeft = True
  Else
    lngPosX = lngPosX + 1 + m_Speed
  End If
  Exit Sub


RightToLeft:
  lngPosY = .ScaleWidth + lngPosX
        
  objProgressBar.Draw .hdc, lngPosX, 0, .ScaleWidth, .ScaleHeight
  objProgressBar.Draw .hdc, lngPosY, 0, .ScaleWidth, .ScaleHeight
        
  If Abs(lngPosX) >= .ScaleWidth Then
    lngPosX = 0
    blnShiftLeft = False
  Else
    lngPosX = lngPosX - m_Speed
  End If
  End With
End Sub

Private Sub UserControl_Initialize()
  Set objProgressBar = New clsProgress
End Sub

Private Sub UserControl_InitProperties()
  Set m_Picture = LoadPicture(vbNullString)
  m_AutoSize = True
  m_Speed = m_def_Speed
  m_Advance = m_def_Advance
End Sub

Private Sub UserControl_Paint()
  objProgressBar.Draw UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Set Picture = .ReadProperty("Picture", Nothing)
    m_AutoSize = .ReadProperty("AutoSize", True)
    m_Speed = .ReadProperty("Speed", m_def_Speed)
    m_Advance = .ReadProperty("Advance", m_def_Advance)
  End With
End Sub

Private Sub UserControl_Resize()
  If m_AutoSize Then
    With UserControl
      .Height = objProgressBar.Height * Screen.TwipsPerPixelY
      .Width = objProgressBar.Width * Screen.TwipsPerPixelX
    End With
  Else
  
  End If
End Sub

Private Sub UserControl_Terminate()
  If Not objTimer Is Nothing Then
    objTimer.Interval = 0
    Set objTimer = Nothing
  End If
  
  Set objProgressBar = Nothing
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Picture", m_Picture, Nothing
    .WriteProperty "AutoSize", m_AutoSize, True
    .WriteProperty "Speed", m_Speed, m_def_Speed
    .WriteProperty "Advance", m_Advance, m_def_Advance
  End With
End Sub


