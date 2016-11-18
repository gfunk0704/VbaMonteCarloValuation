VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbsoluteValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements NegativeHandler

Public Function NegativeHandler_handleNegative(value As Double) As Double
NegativeHandler_handleNegative = Abs(value)
End Function


