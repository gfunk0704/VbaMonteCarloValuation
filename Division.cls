VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Division"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements ArithmeticOperator
Friend Function ArithmeticOperator_evaluate(elem1 As Double, elem2 As Double) As Double
ArithmeticOperator_evaluate = elem1 / elem2
End Function


