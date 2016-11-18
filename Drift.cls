VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Drift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Property Let parameters(params As Collection): End Property
Public Property Get parameters() As Collection: End Property
Public Function nextIncrement(value As Double) As Double: End Function
Public Function clone() As Drift: End Function

