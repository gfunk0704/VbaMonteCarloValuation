VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EulerDescretization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private driftTerm_ As Drift
Private duffusionTerm_ As Diffusion
Private initialValue_ As Double
Private preValue As Double
Public handler As NegativeHandler

Public Property Set driftTerm(value As Drift)
Set driftTerm_ = value.clone
End Property

Public Property Get driftTerm() As Drift
Set driftTerm = driftTerm_.clone
End Property

Public Property Set diffusionTerm(value As Diffusion)
Set duffusionTerm_ = value.clone
End Property

Public Property Get diffusionTerm() As Diffusion
Set diffusionTerm = diffusionTerm.clone
End Property

Public Property Let parameters(params As Collection)

If hasItem(params, "initialValue") Then
    initialValue_ = params("initialValue")
    preValue = initialValue_
End If

driftTerm_.parameters = params
duffusionTerm_.parameters = params
End Property

Public Function initialValue() As Double
initialValue = initialValue_
End Function

Sub reset()
preValue = initialValue
End Sub

Public Function nextStep(randNum As Double) As Double
Dim increment As Double
increment = driftTerm_.nextIncrement(preValue) + duffusionTerm_.nextIncrement(preValue, randNum)
preValue = handler.handleNegative(preValue + increment)
nextStep = preValue
End Function




