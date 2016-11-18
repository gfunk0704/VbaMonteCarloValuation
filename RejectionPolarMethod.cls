VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RejectionPolarMethod"
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

Dim v1 As Double
Dim v2 As Double
Dim rSqr As Double
Dim timesScalar As Double

If hasAdditional Then
    hasAdditional = False
    NormalRandNum_nextNormalRnd = addRandNum
Else
    hasAdditional = True
    rSqr = 2
    
    Do While rSqr >= 1
        v1 = 2 * Rnd - 1
        v2 = 2 * Rnd - 1
        rSqr = v1 * v1 + v2 * v2
    Loop
    
    timesScalar = Sqr(-2 * Log(rSqr) / rSqr)
    addRandNum = v2 * rSqr
    NormalRandNum_nextNormalRnd = v1 * rSqr
End If

End Function




