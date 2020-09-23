VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElapsedTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides time in specified timebase between a start and stop call or a running count."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Function to get current clock count
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
  lpPerformanceCount As Currency) As Long
  
'Function to see how many clock ticks per second on this computer
Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
  lpPerformanceCount As Currency) As Long
  
'local variable(s) to hold property value(s)

'How many tic's per second on this computer
Private mvarSecFrequency As Currency

'Clock time when timer was started
Private mvarStartTime As Currency

'TRUE or FALSE to know if we are counting time
Private mvarIsRunning As Boolean

'How much time has elapsed since timer was started
Private mvarElapseTime As Currency

'When the STOP timer was issued, also used to re-adjust
'mvarStartTime when restarting the timer
Private mvarStopTime As Currency


Public Function StopTimer() As Double
    
    If mvarIsRunning Then QueryPerformanceCounter mvarStopTime
    StopTimer = (mvarStopTime - mvarStartTime) / mvarSecFrequency
    mvarIsRunning = False
    
End Function

Public Function Elapsed() As Double
    
    If mvarIsRunning Then QueryPerformanceCounter mvarElapseTime
    Elapsed = (mvarElapseTime - mvarStartTime) / mvarSecFrequency

End Function

Public Sub Start()
    
    If Not mvarIsRunning Then
        If (mvarStopTime > 0) Then
            mvarStopTime = mvarStartTime - mvarStopTime
        End If
        QueryPerformanceCounter mvarStartTime
        mvarStartTime = mvarStartTime + mvarStopTime
        mvarStopTime = 0
        mvarIsRunning = True
    End If
    
End Sub

Public Sub Reset()
    
    mvarStartTime = 0
    mvarElapseTime = 0
    mvarStopTime = 0
    QueryPerformanceFrequency mvarSecFrequency
    mvarIsRunning = False

End Sub
Public Function IsRunning() As Boolean
    IsRunning = mvarIsRunning
End Function
Private Sub Class_Initialize()
  
    Reset
  
End Sub

Public Sub kill()

    mvarIsRunning = False

End Sub
