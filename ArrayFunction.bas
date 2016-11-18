Attribute VB_Name = "ArrayFunction"
Option Explicit
Option Base 1

Public Function getDimension(inputArray As Variant) As Integer

Dim temp As Integer
Dim dimension As Integer

On Error GoTo getResult
Do While True
    dimension = dimension + 1
    temp = UBound(inputArray, dimension)
Loop

getResult:
getDimension = dimension - 1
End Function

