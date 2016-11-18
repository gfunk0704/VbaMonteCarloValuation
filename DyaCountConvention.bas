Attribute VB_Name = "DyaCountConvention"
Option Explicit
Option Base 1

Public Enum DayCountConvention
    ACT360 = 0
    ACT365F = 1
End Enum

Public Function getYearDays(dayCount As DayCountConvention) As Integer

Select Case dayCount
    Case ACT360
        getYearDays = 360
    Case ACT365F
        getYearDays = 365
End Select

End Function

