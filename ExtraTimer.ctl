VERSION 5.00
Begin VB.UserControl ExtraTimer 
   BackColor       =   &H00FF80FF&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ExtraTimer.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ExtraTimer.ctx":08CA
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ExtraTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default property values
Const m_def_Interval = 1
'Property Variables
Dim m_Interval As Long
Dim m_IntervalUnit As IntervalUnits ' A Class Module was created for it
                                    ' to make IntervalUnits a data type
                                    ' and set its possible values
'Other Variables
Dim StartTime As Long
'Event Declaration
Event Timer()

Private Sub Timer1_Timer()
' Keep record of the number of seconds that have elasped using the global
' variable starttime
StartTime = StartTime + 1
' Use select case to determine what the interval unit is, then use the
' If....Then Statement to check if the time elapsed if upto the set
' interval. If it is then raise the timer event for the control and set
' the global variable "StartTime" to be equal to zero.
Select Case m_IntervalUnit
Case 0
If StartTime >= m_Interval Then
RaiseEvent Timer
StartTime = 0
End If
Case 1
If StartTime >= (m_Interval * 60) Then
RaiseEvent Timer
StartTime = 0
End If
Case 2
If StartTime >= (m_Interval * 3600) Then
RaiseEvent Timer
StartTime = 0
End If
Case 3
If StartTime >= (m_Interval * 86400) Then
RaiseEvent Timer
StartTime = 0
End If
End Select
End Sub

Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_InitProperties()
' Set the initial values of your properties the first time the control is
' placed on a form
m_Interval = m_def_Interval
m_IntervalUnit = 0
StartTime = 0
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
    Timer1.Enabled = PropBag.ReadProperty("Enabled", False)
    m_IntervalUnit = PropBag.ReadProperty("IntervalUnit", 0)
End Sub

Private Sub UserControl_Resize()
' Set the size of the control so it is fixed
If UserControl.Height <> 300 Then UserControl.Height = 300
If UserControl.Width <> 300 Then UserControl.Width = 300
End Sub
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/Sets the number of (Seconds, Minutes, Hours or Days)\r\nbetween calls to the controls timer event."
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    m_Interval = New_Interval
    PropertyChanged "Interval"
End Property

Public Property Get IntervalUnit() As IntervalUnits
Attribute IntervalUnit.VB_Description = "Returns/Sets a value that determines the unit used in \r\nsetting the controls interval"
IntervalUnit = m_IntervalUnit
End Property

Public Property Let IntervalUnit(ByVal New_IntervalUnit As IntervalUnits)
m_IntervalUnit = New_IntervalUnit
PropertyChanged "IntervalUnit"
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/Sets a value that determines \r\nwhether the control can respond to user\r\ngenerated events."
Enabled = Timer1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Timer1.Enabled = New_Enabled
PropertyChanged "Enabled"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' The various Properties of the control, Interval sets the interval at
' which the control fires its timer event, IntervalUnits sets the Unit
' of the interval be it Seconds, Minutes, Hours or Days, while the Enabled
' property sets if the control is enabled or not.
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
    Call PropBag.WriteProperty("Enabled", Timer1.Enabled, False)
    Call PropBag.WriteProperty("IntervalUnit", m_IntervalUnit, 0)
End Sub
