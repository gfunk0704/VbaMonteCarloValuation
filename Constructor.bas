Attribute VB_Name = "Constructor"
Option Explicit
Option Base 1

Public Enum ModelType
    GBM = 0
    CEV = 1
    Vasicek = 2
    CIR = 3
End Enum

Public Enum RNGType
    CDFInverse = 0
    BoxMuller = 1
    RejectionPolar = 2
End Enum

Public Function modelFactory(model As ModelType, parameters As Collection) As EulerDescretization

Dim sdeDrift As Drift
Dim sdeDiffusion As Diffusion

Select Case model
    Case GBM
        Set sdeDrift = New GBMDrift
        Set sdeDiffusion = New GBMDiffusion
    Case CEV
        Set sdeDrift = New GBMDrift
        Set sdeDiffusion = New CEVDiffusion
    Case Vasicek
        Set sdeDrift = New OrnsteinUhlenbeckDrift
        Set sdeDiffusion = New OrnsteinUhlenbeckDiffusion
    Case CIR
        Set sdeDrift = New OrnsteinUhlenbeckDrift
        Set sdeDiffusion = New CIRDiffusion
End Select

Dim sde As New EulerDescretization
Set sde.driftTerm = sdeDrift
Set sde.diffusionTerm = sdeDiffusion
sde.parameters = parameters

Set modelFactory = sde
End Function

Public Function rngFactory(methodType As RNGType, seed As Variant) As NormalRandNum

Dim randNumGenerator As NormalRandNum

Select Case methodType
    Case CDFInverse
        Set randNumGenerator = New CDFInverseMethod
    Case BoxMuller
        Set randNumGenerator = New BoxMullerMethod
    Case RejectionPolar
        Set randNumGenerator = New RejectionPolarMethod
End Select

randNumGenerator.seed = CDbl(seed)
Set rngFactory = randNumGenerator
End Function
