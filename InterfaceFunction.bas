Attribute VB_Name = "InterfaceFunction"
Option Explicit
Option Base 1

Public Type SamplingSetting
    nStep As Integer
    nPath As Integer
    handler As NegativeHandler
End Type

Public Function simulationSetting(nStep As Integer, nPath As Integer, handler As NegativeHandler) As SamplingSetting
If (nStep <= 0) Or (nPath <= 0) Then
    Err.Raise Number:=13001, source:="simulationSetting", Description:="步數與路徑數須為正數"
End If
   
simulationSetting.nPath = nPath
simulationSetting.nStep = nStep
Set simulationSetting.handler = handler
End Function

Public Function simulate(model As EulerDescretization, rng As NormalRandNum, setting As SamplingSetting) As Matrix

Dim nStep As Integer: nStep = setting.nStep
Dim nPath As Integer: nPath = setting.nPath
Dim i As Integer
Dim j As Integer
ReDim samplingResult(nStep + 1, nPath) As Double

Set model.handler = setting.handler
rng.initialize

For j = 1 To nPath

    samplingResult(1, j) = model.initialValue
    
    For i = 1 To nStep
        samplingResult(i + 1, j) = model.nextStep(rng.nextNormalRnd)
    Next i
    
    model.reset
Next j

Dim output As New Matrix
output.data = samplingResult
Set simulate = output
End Function




    

