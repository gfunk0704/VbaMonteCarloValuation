Attribute VB_Name = "MatrixFunction"
Option Explicit
Option Base 1

Public Function isColVec(inputMatrix As Matrix) As Boolean

If inputMatrix.nCol = 1 Then
    isColVec = True
Else
    isColVec = False
End If

End Function

Public Function isRowVec(inputMatrix As Matrix) As Boolean

If inputMatrix.nRow = 1 Then
    isRowVec = True
esle
    isRowVec = False
End If

End Function

Public Function elementwiseOperation(mat1 As Matrix, operator As ArithmeticOperator, mat2 As Matrix) As Matrix

If (mat1.nRow <> mat2.nRow) Or (mat1.nCol <> mat2.nCol) Then
    Err.Raise Number:=10002, source:="elementwiseOperation", Description:="兩矩陣行列不符"
End If

Dim i As Integer
Dim j As Integer
Dim nRow As Integer: nRow = mat1.nRow
Dim nCol As Integer: nCol = mat1.nCol
Dim data1() As Double: data1 = mat1.data
Dim data2() As Double: data2 = mat2.data
ReDim dataOutput(nRow, nCol) As Double

For i = 1 To nRow
    For j = 1 To nCol
        dataOutput(i, j) = operator.evaluate(data1(i, j), data2(i, j))
    Next j
Next i

Dim outputMat As New Matrix
outputMat.data = dataOutput
Set elementwiseOperation = outputMat

End Function

Public Function matrixMultiplication(mat1 As Matrix, mat2 As Matrix) As Matrix

If mat1.nCol <> mat2.nRow Then
    Err.Raise Number:=10002, source:="matrixMultiplication", Description:="兩矩陣行列不符"
End If

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim data1() As Double: data1 = mat1.data
Dim data2() As Double: data2 = mat2.data
ReDim dataOutput(mat1.nRow, mat2.nCol) As Double

For i = 1 To mat1.nRow
    For j = 1 To mat2.nCol
        For k = 1 To mat1.nCol
            dataOutput(i, j) = dataOutput(i, j) + data1(i, k) * data2(k, j)
        Next k
    Next j
Next i

Dim outputMat As New Matrix
outputMat.data = dataOutput
Set matrixMultiplication = outputMat
End Function

