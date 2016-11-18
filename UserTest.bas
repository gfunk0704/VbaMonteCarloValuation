Attribute VB_Name = "UserTest"
Option Explicit
Option Base 1

Sub simulationExample()
' �إ߽X����
Dim benchmark As New Stopwatch
' �ҫ�����b�U�d�ҨM�w
Dim sde As EulerDescretization:

' �p�t�Q�v�o�͡A�h�N�Ӵ����u�Q�]���s
' steps : 500
' paths: 3000
Dim setting As SamplingSetting: setting = simulationSetting(500, 3000, New NegativeValueToZero)

'comparing randnom number generator
Dim rng As NormalRandNum

'�]�w���հѼ�
Dim params As New Collection
params.add Item:=0.003, key:="initialValue"
params.add Item:=0.2, key:="kappa"
params.add Item:=0.009, key:="theta"
params.add Item:=0.1, key:="sigma"
params.add Item:=1 / 365, key:="dt"
'�������G��X�x�}
Dim simResult As Matrix

' case 1.  CDF Inverse method + Vasicek model
benchmark.startTimer
 Set sde = modelFactory(Vasicek, params)
 Set rng = rngFactory(CDFInverse, Date)

Set simResult = simulate(sde, rng, setting)

simResult.printToWorksheet
MsgBox "��O�ɶ�: " & benchmark.elapsedTime
benchmark.reset
 
'case 2.  Rejection polar method + CIR model
benchmark.startTimer
Set sde = modelFactory(CIR, params)
Set rng = rngFactory(RejectionPolar, Date)

Set simResult = simulate(sde, rng, setting)

simResult.printToWorksheet
MsgBox "��O�ɶ�: " & benchmark.elapsedTime

End Sub






















