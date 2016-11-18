VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NormalRandNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Property Let seed(ByVal value As Double): End Property
Public Property Get seed() As Double: End Property
Public Sub initialize(): End Sub
Public Function nextNormalRnd() As Double: End Function

