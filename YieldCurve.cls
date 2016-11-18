VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YieldCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private valueDate_ As Date
Private pointsInCurve() As Rate
Private nRate As Integer
Public dayCounts As DayCountConvention

Public Property Let valueDate(value As Date)
valueDate_ = value
sortPoints
End Property

Public Property Get valueDate() As Date
valueDate = valueDate_
End Property

Public Sub add(newRate As Rate)

If nRate > 0 Then
    Dim i As Integer
    Dim newMaturityDate As Date: newMaturityDate = newRate.maturity.toMaturityDate(valueDate)
    
    For i = 1 To nRate
        If pointsInCurve(i).maturity.toMaturityDate(valueDate) = newMaturityDate Then
            If pointsInCurve(i).value <> newRate.value Then
                Err.Raise Number:=15001, source:="add", Description:="同一個時間點只有單一利率"
            Else
                Exit Sub
            End If
        End If
    Next i
    
    nRate = nRate + 1
    ReDim Preserve pointsInCurve(nRate)
    pointsInCurve(nRate) = newRate
    sortPoints
Else
    ReDim pointsInCurve(1)
    pointsInCurve(1) = newRate
    nRate = 1
End If

End Sub


Private Sub sortPoints()

Dim temp As Rate
Dim i As Integer
Dim j As Integer
ReDim maturityDates(nRate) As Date

For i = 1 To nRate
    maturityDates(i) = pointsInCurve(i).maturity.toMaturityDate(valueDate)
Next i

For i = 1 To nRate - 1
    For j = i + 1 To nRate
        If maturityDates(i) > maturityDates(j) Then
            temp = pointsInCurve(i)
            pointsInCurve(i) = pointsInCurve(j)
            pointsInCurve(j) = temp
        End If
    Next j
Next i

End Sub


