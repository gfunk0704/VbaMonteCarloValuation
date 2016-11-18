VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Enum Margin
    byRow = 0
    byCol = 1
End Enum

Private data_() As Double
Private nRow_ As Integer
Private nCol_ As Double

Public Property Let data(value() As Double)

If ArrayFunction.getDimension(value) <> 2 Then
    errorInputIncorrect ("data")
End If

data_ = value
nRow_ = UBound(value, 1)
nCol_ = UBound(value, 2)

End Property

Public Property Get data() As Double()
data = data_
End Property

Public Function nRow() As Integer
nRow = nRow_
End Function

Public Function nCol() As Integer
nCol = nCol_
End Function

Public Property Let elem(i As Integer, j As Integer, ByVal value As Double)
data_(i, j) = value
End Property

Public Property Get elem(i As Integer, j As Integer) As Double
elem = data_(i, j)
End Property

Public Property Let col(i As Integer, ByVal colVec As Matrix)

If Not MatrixFunction.isColVec(colVec) Then
    errorInputIncorrect ("col")
End If

Dim j As Integer

For j = 1 To nRow_
    data_(j, i) = colVec.elem(j, 1)
Next j

End Property

Public Property Get col(i As Integer) As Matrix

Dim j As Integer
ReDim vecData(nRow_, 1) As Double

For j = 1 To nRow_
    vecData(j, 1) = data_(j, i)
Next j

Dim colVec As New Matrix
colVec.data = vecData

Set col = colVec
End Property

Public Property Let row(i As Integer, ByVal rowVec As Matrix)

If Not MatrixFunction.isRowVec(rowVec) Then
    errorInputIncorrect ("row")
End If

Dim j As Integer

For j = 1 To nCol_
    data_(i, j) = rowVec.elem(1, j)
Next j

End Property

Public Property Get row(i As Integer) As Matrix

Dim j As Integer
ReDim vecData(1, nCol_) As Double

For j = 1 To nCol_
    vecData(1, j) = data_(i, j)
Next j

Dim rowVec As New Matrix
rowVec.data = vecData
Set row = rowVec

End Property

Public Property Let subMatrix(startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer, ByVal value As Matrix)

If (startRow > endRow) Or (startCol > endCol) Then
    errorInputIncorrect ("subMatrix")
End If

Dim i As Integer
Dim j As Integer

For i = startRow To endRow
    For j = startCol To endCol
        data_(i, j) = value.elem(i - startRow + 1, j - startCol + 1)
    Next j
Next i

End Property

Public Property Get subMatrix(startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer) As Matrix

If (startRow > endRow) Or (startCol > endCol) Then
    errorInputIncorrect ("subMatrix")
End If

Dim i As Integer
Dim j As Integer
ReDim subMatData(endRow - startRow + 1, endCol - startCol + 1) As Double

For i = startRow To endRow
    For j = startCol To endCol
        subMatData(i - startRow + 1, j - startCol + 1) = data_(i, j)
    Next j
Next i

Dim subMat As New Matrix
subMat.data = subMatData
Set subMatrix = subMat

End Property

Public Function findIf(operator As ComparisonOperator, value As Double) As Variant
Dim nElem As Long: nElem = nRow_ * nCol_

Dim counts As Long
Dim i As Integer
Dim j As Integer
ReDim findIndex(2, nRow_ * nCol_) As Integer

For i = 1 To nRow_
    For j = 1 To nCol_
        If operator.compare(data_(i, j), value) Then
            counts = counts + 1
            findIndex(1, counts) = i
            findIndex(2, counts) = j
        End If
    Next j
Next i

If counts > 0 Then
    ReDim Preserve findIndex(2, counts)
    findIf = findIndex
Else
    findIf = Null
End If

End Function

Public Sub transpose()

Dim i As Integer
Dim j  As Integer
ReDim tData(nCol_, nRow_) As Double

For i = 1 To nRow_
    For j = 1 To nCol_
        tData(j, i) = data_(i, j)
    Next j
Next i

nRow_ = UBound(tData, 1)
nCol_ = UBound(tData, 2)
data_ = tData

End Sub

Public Sub operationOnScalar(operator As ArithmeticOperator, value As Double)

Dim i As Integer
Dim j  As Integer


For i = 1 To nRow_
    For j = 1 To nCol_
        data_(i, j) = operator.evaluate(data_(i, j), value)
    Next j
Next i

End Sub

Public Sub cumulative(operator As ArithmeticOperator, by As Margin)

Dim i As Integer
Dim j As Integer

Select Case by
    Case byRow
        For i = 1 To nRow_
            For j = 2 To nCol_
                data_(i, j) = operator.evaluate(data_(i, j), data_(i, j - 1))
            Next j
        Next i
    Case byCol
        For i = 2 To nRow_
            For j = 1 To nCol_
                data_(i, j) = operator.evaluate(data_(i, j), data_(i - 1, j))
            Next j
        Next i
End Select

End Sub

Public Function clone() As Matrix
Dim cloneMat As New Matrix

cloneMat.data = data_
Set clone = cloneMat
End Function

Public Sub fill(value As Double)
Dim i As Integer
Dim j As Integer

For i = 1 To nRow_
    For j = 1 To nCol_
        data_(i, j) = value
    Next j
Next i

End Sub

Public Sub printToWorksheet(Optional sheetName As Variant)

Application.ScreenUpdating = False

Dim nSheets As Integer: nSheets = Worksheets.Count
Worksheets.add After:=Worksheets(nSheets)
Worksheets(nSheets + 1).Range("A1", Cells(nRow_, nCol_)).value = data_

If Not IsMissing(sheetName) Then
    Worksheets(nSheets + 1).name = sheetName
End If

End Sub


Private Sub errorInputIncorrect(errSource As String)
Err.Raise Number:=10001, source:=errSource, Description:="輸入資料格式錯誤"
End Sub


