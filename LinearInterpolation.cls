VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearInterpolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements Interpolator

Private xData() As Double
Private yData() As Double
Private nData As Integer

Public Sub Interpolator_setData(xVec() As Double, yVec() As Double)

nData = UBound(xVec)

If nData <> UBound(yVec) Then
    Err.Raise Number:=16001, source:="setData", Description:="x, y 資料長度不同"
End If

xData = xVec
yData = yVec

End Sub

Public Function Interpolator_interpolate(x As Double) As Double

Dim nextLocation As Integer: nextLocation = 1
Dim preLocation As Integer

If x <= xData(1) Then
    Interpolator_interpolate = yData(1)
ElseIf x >= xData(nData) Then
    Interpolator_interpolate = yData(nData)
Else
    Do While x > xData(nextLocation)
        nextLocation = nextLocation + 1
    Loop
    
    preLocation = nextLocation - 1
    Interpolator_interpolate = yData(preLocation) + (yData(nextLocation) - yData(preLocation)) _
                                                        * (x - xData(preLocation)) / (xData(nextLocation) - xData(preLocation))
End If
End Function
