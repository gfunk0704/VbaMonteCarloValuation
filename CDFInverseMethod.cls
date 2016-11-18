VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDFInverseMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements NormalRandNum

'°Ñ¦Ò Glasserman [2004], P.67
Private seed_ As Double

Friend Property Let NormalRandNum_seed(ByVal value As Double)
seed_ = Abs(value)
End Property

Friend Property Get NormalRandNum_seed() As Double
NormalRandNum_seed = seed_
End Property

Friend Sub NormalRandNum_initialize()
Rnd -1
Randomize seed_
End Sub

Friend Function NormalRandNum_nextNormalRnd() As Double
NormalRandNum_nextNormalRnd = normalInverse(Rnd)
End Function


Private Function normalInverse(u As Double) As Double

Static isInitialized As Boolean
Static a() As Double
Static b() As Double
Static c() As Double

If Not isInitialized Then
    isInitialized = True
    
    ReDim a(0 To 3)
    ReDim b(0 To 3)
    ReDim c(0 To 8)
    
    'a vector
    a(0) = 2.50662823884
    a(1) = -18.61500062529
    a(2) = 41.39119773534
    a(3) = -25.44106049637
    
    'b vector
    b(0) = -8.4735109309
    b(1) = 23.08336743743
    b(2) = -21.06224101826
    b(3) = 3.13082909833
    
    'c vector
    c(0) = 0.337475482272615
    c(1) = 0.976169019091719
    c(2) = 0.160797971491821
    c(3) = 2.76438810333863E-02
    c(4) = 3.8405729373609E-03
    c(5) = 3.951896511919E-04
    c(6) = 3.21767881768E-05
    c(7) = 2.888167364E-07
    c(8) = 3.960315187E-07
End If

Dim y As Double: y = u - 0.5
Dim phi As Double: phi = 1
Dim r As Double
Dim i As Integer

If Abs(y) < 0.42 Then
    r = y * y
    normalInverse = y * (((a(3) * r + a(2)) * r + a(1)) * r + a(0)) / ((((b(3) * r + b(2)) * r + b(1)) * r + b(0)) * r + 1)
Else
    r = u
    
    If y > 0 Then
        r = 1 - u
    End If
    
    r = Log(-Log(r))
    
    If y < 0 Then
        phi = -1
    End If
    
    normalInverse = c(8)
    
    For i = 7 To 0 Step -1
        normalInverse = normalInverse * r + c(i)
    Next i
    
    normalInverse = normalInverse * phi
End If
End Function

