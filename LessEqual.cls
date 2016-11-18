VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LessEqual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements ComparisonOperator
Friend Function ComparisonOperator_compare(elem1 As Variant, elem2 As Variant) As Boolean
ComparisonOperator_compare = elem1 <= elem2
End Function

