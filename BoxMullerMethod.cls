VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoxMullerMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements NormalRandNum

Private seed_ As Double
Private addRandNum As Double
Private hasAdditional As Boolean
Private Const pi As Double = 3.14159265358979

Friend Property Let NormalRandNum_seed(ByVal value As Double)
seed_ = Abs(value)
End Property

Friend Property Get NormalRandNum_seed() As Double
NormalRandNum_seed = seed_
End Property

Friend Sub NormalRandNum_initialize()
Rnd -1
Randomize seed_
hasAdditional = False
End Sub

Friend Function NormalRandNum_nextNormalRnd() As Double

Dim unifRnd1 As Double
Dim unifRnd2 As Double

If hasAdditional Then
    hasAdditional = False
    NormalRandNum_nextNormalRnd = addRandNum
Else
    hasAdditional = True
    unifRnd1 = Rnd
    unifRnd2 = Rnd
    addRandNum = Sqr(-2 * Log(unifRnd1)) * Sin(2 * pi * unifRnd2)
    NormalRandNum_nextNormalRnd = Sqr(-2 * Log(unifRnd1)) * Cos(2 * pi * unifRnd2)
End If

End Function



