Attribute VB_Name = "RateObj"
Option Explicit
Option Base 1

Public Type Rate
    value As Double
    maturity As Tenor
End Type

Public Function toRate(value As Double, tenorString As String) As Rate

Dim maturity As New Tenor
maturity.symbol = tenorString

Set toRate.maturity = maturity
toRate.value = value

End Function
