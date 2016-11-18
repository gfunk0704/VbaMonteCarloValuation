VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private startTime As Double
Private stopTime As Double
Private elapsedTime_ As Double
Private timeRunning As Boolean

Public Sub startTimer()
If Not timeRunning Then
    timeRunning = True
    startTime = Timer
End If
End Sub

Public Sub stopTimer()
If timeRunning Then
    stopTime = Timer
    elapsedTime_ = elapsedTime_ + stopTime - startTime
    timeRunning = False
End If
End Sub

Public Function elapsedTime() As Double
Dim extraTime As Double
If timeRunning Then extraTime = Timer - startTime
elapsedTime = extraTime + elapsedTime_
End Function

Public Sub reset()
elapsedTime_ = 0
timeRunning = False
End Sub
