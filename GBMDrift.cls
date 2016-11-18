VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GBMDrift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Implements Drift

Private mu As Double
Private dt As Double

Friend Property Let Drift_parameters(params As Collection)

If hasItem(params, "mu") Then mu = params("mu")
If hasItem(params, "dt") Then dt = params("dt")

If dt < 0 Then
    Err.Raise Number:=12000, source:="Drift_parameter", Description:="dt 須為正數"
End If

End Property

Friend Property Get Drift_parameters() As Collection
Set Drift_parameters = getParams()
End Property

Friend Function Drift_nextIncrement(value As Double) As Double
Drift_nextIncrement = mu * value * dt
End Function

Friend Function Drift_clone() As Drift

Dim cloneDrift As New GBMDrift
cloneDrift.Drift_parameters = getParams()
Set Drift_clone = cloneDrift
End Function

Private Function getParams() As Collection
Dim params As New Collection
params.add Item:=mu, key:="mu"
params.add Item:=dt, key:="dt"
Set getParams = params
End Function

