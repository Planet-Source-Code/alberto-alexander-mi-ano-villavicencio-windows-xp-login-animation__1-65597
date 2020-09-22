Attribute VB_Name = "modSubTimer"
Option Explicit

'This module shows how to create a timer event using the SetTimer and KillTimer
'Windows API functions.


'Declares:
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Const TIMER_MAX As Long = 1024

'Array of timers.
Public aTimers(1 To TIMER_MAX) As clsTimerPlus
Attribute aTimers.VB_VarDescription = "Array of timers."

'Added SPM to prevent excessive searching through aTimers array:
Private m_lngTimerCount As Long

'Create the timer.
Function TimerCreate(timer As clsTimerPlus) As Boolean
Attribute TimerCreate.VB_Description = "Create the timer."
  timer.TimerID = SetTimer(0&, 0&, timer.Interval, AddressOf TimerProc)
  
  If timer.TimerID Then
    TimerCreate = True
    
    'Find the timer para ver si no se a superado el máximo de timer permitidos.
    Dim i As Long
    
    For i = 1 To TIMER_MAX
      If aTimers(i) Is Nothing Then
        Set aTimers(i) = timer
        If (i > m_lngTimerCount) Then
          m_lngTimerCount = i
        End If
        TimerCreate = True
        Exit Function
      End If
    Next
    timer.ErrRaise eeTooManyTimers
  Else
    'TimerCreate = False
    timer.TimerID = 0
    timer.Interval = 0
  End If
End Function

'Destroy the timer.
Public Function TimerDestroy(timer As clsTimerPlus) As Long
Attribute TimerDestroy.VB_Description = "Destroy the timer."
  Dim i As Long
  Dim f As Boolean
  
  'Find and remove this timer
  'TimerDestroy = False
  
  'SPM - no need to count past the last timer set up in the
  'aTimer array:
  For i = 1 To m_lngTimerCount
    'Find timer in array.
    If Not aTimers(i) Is Nothing Then
      If timer.TimerID = aTimers(i).TimerID Then
        f = KillTimer(0, timer.TimerID)
        'Remove timer and set reference to nothing.
        Set aTimers(i) = Nothing
        TimerDestroy = True
        Exit Function
      End If
    'SPM: aTimers(1) could well be nothing before
    'aTimers(2) is.  This original [else] would leave
    'timer 2 still running when the class terminates -
    'not very nice!  Causes serious GPF in IE and VB design
    'mode...
    'Else
    '  TimerDestroy = True
    '  Exit Function
    End If
  Next i
End Function


Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Attribute TimerProc.VB_Description = "Address Of timer procedure."
  Dim i As Long
    
  'Find the timer with this ID.
  For i = 1 To m_lngTimerCount
    'SPM: Add a check to ensure aTimers(i) is not nothing!
    'This would occur if we had two timers declared from
    'the same thread and we terminated the first one before
    'the second!  Causes serious GPF if we don't do this...
    If Not (aTimers(i) Is Nothing) Then
      If idEvent = aTimers(i).TimerID Then
        aTimers(i).PulseTimer   'Generate the event.
        Exit Sub
      End If
    End If
  Next i
End Sub
