VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimerPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Use these class if need to execute events at intervals greater than what
'the Visual Basic Timer control allows. The Timer control allows a maximum
'interval of slightly over a minute while these class allow you to set an
'interval up to 24.86 days.
'
'NOTE: This class needs for your operation of the module modSubtimer


Private m_lngInterval As Long
Private m_lngTimerID As Long
Public Tag As Variant       'User can attach any Variant data they want to the timer.
Attribute Tag.VB_VarDescription = "Gets/sets a piece of additional data to store with the timer."


'modSubtimer is independent of VBCore, so it hard codes error handling.
Public Enum EErrorTimer
  eeBaseTimer = 13650     'clsTimerPlus
  eeTooManyTimers         'No more than 1024 timers allowed per class
  eeCantCreateTimer       'Can't create system timer
End Enum

Public Event Tick()


Property Get Interval() As Long
Attribute Interval.VB_Description = "Can't just change interval--you must kill timer and start a new one. These propertie allow you to set an interval up to 24.86 days."
  Interval = m_lngInterval
End Property

'Can't just change interval--you must kill timer and start a new one
Property Let Interval(iIntervalA As Long)
  Dim f As Boolean
  
  If iIntervalA > 0 Then
    'Don't mess with it if interval is the same.
    If m_lngInterval = iIntervalA Then Exit Property
    
    'Must destroy any existing timer to change interval.
    If m_lngInterval Then
      f = TimerDestroy(Me)
      Debug.Assert f      'Shouldn't fail.
    End If
    'Create new timer with new interval
    m_lngInterval = iIntervalA
    If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer
  Else
    If (m_lngInterval > 0) Then
      m_lngInterval = 0
      f = TimerDestroy(Me)
      Debug.Assert f      'Shouldn't fail.
    End If
  End If
End Property

Friend Property Get TimerID() As Long
  TimerID = m_lngTimerID
End Property

Friend Property Let TimerID(New_TimerID As Long)
  m_lngTimerID = New_TimerID
End Property


Private Sub Class_Terminate()
  Interval = 0
End Sub


Friend Sub ErrRaise(e As Long)
  Dim sText As String
  Dim sSource As String
  
  If e > 1000 Then
    sSource = App.EXEName & ".WindowProc"
    Select Case e
      Case eeTooManyTimers: sText = "No more than 1024 timers allowed per class"
      Case eeCantCreateTimer: sText = "Can't create system timer"
    End Select
    Err.Raise e Or vbObjectError, sSource, sText
  Else
    Err.Raise e, sSource    'Raise standard Visual Basic error
  End If
End Sub


'Must be public so that Timer object can't terminate while client's Tick
'event is being processed--Friend wouldn't prevent this disaster.
Public Sub PulseTimer()
  RaiseEvent Tick
End Sub
