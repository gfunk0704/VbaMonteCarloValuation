Attribute VB_Name = "UserTest"
Option Explicit
Option Base 1

Sub simulationExample()
' 建立碼表物件
Dim benchmark As New Stopwatch
' 模型物件在各範例決定
Dim sde As EulerDescretization:

' 如負利率發生，則將該期間短利設為零
' steps : 500
' paths: 3000
Dim setting As SamplingSetting: setting = simulationSetting(500, 3000, New NegativeValueToZero)

'comparing randnom number generator
Dim rng As NormalRandNum

'設定測試參數
Dim params As New Collection
params.add Item:=0.003, key:="initialValue"
params.add Item:=0.2, key:="kappa"
params.add Item:=0.009, key:="theta"
params.add Item:=0.1, key:="sigma"
params.add Item:=1 / 365, key:="dt"
'模擬結果輸出矩陣
Dim simResult As Matrix

' case 1.  CDF Inverse method + Vasicek model
benchmark.startTimer
 Set sde = modelFactory(Vasicek, params)
 Set rng = rngFactory(CDFInverse, Date)

Set simResult = simulate(sde, rng, setting)

simResult.printToWorksheet
MsgBox "花費時間: " & benchmark.elapsedTime
benchmark.reset
 
'case 2.  Rejection polar method + CIR model
benchmark.startTimer
Set sde = modelFactory(CIR, params)
Set rng = rngFactory(RejectionPolar, Date)

Set simResult = simulate(sde, rng, setting)

simResult.printToWorksheet
MsgBox "花費時間: " & benchmark.elapsedTime

End Sub






















