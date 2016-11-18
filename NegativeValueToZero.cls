VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NegativeValueToZero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements NegativeHandler

Public Function NegativeHandler_handleNegative(value As Double) As Double
If value > 0 Then NegativeHandler_handleNegative = value
End Function


