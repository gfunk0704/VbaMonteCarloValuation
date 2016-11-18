VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrnsteinUhlenbeckDiffusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Implements Diffusion

Private sigma As Double
Private sqrtDt As Double

Public Property Let Diffusion_parameters(params As Collection)

If hasItem(params, "sigma") Then sigma = params("sigma")
If hasItem(params, "dt") Then sqrtDt = Sqr(params("dt"))

If (sqrtDt < 0) Or (sigma < 0) Then
    Err.Raise Number:=12001, source:="Diffusion_parameter", Description:="參數須為正數"
End If

End Property

Public Property Get Diffusion_parameters() As Collection
Set Diffusion_parameters = getParams()
End Property

Friend Function Diffusion_clone() As Diffusion

Dim cloneDiffusion As New OrnsteinUhlenbeckDiffusion
cloneDiffusion.Diffusion_parameters = getParams()
Set Diffusion_clone = cloneDiffusion
End Function

Public Function Diffusion_nextIncrement(value As Double, randNum As Double) As Double
Diffusion_nextIncrement = sigma * sqrtDt * randNum
End Function

Private Function getParams() As Collection
Dim params As New Collection

params.add Item:=sigma, key:="sigma"
params.add Item:=(sqrtDt * sqrtDt), key:="dt"

Set getParams = params
End Function

