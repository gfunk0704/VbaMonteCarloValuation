VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tenor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private countBase As String
Private maturityNum As Integer


Public Property Let symbol(tenorString As String)

Dim nCharacter As Integer: nCharacter = Len(tenorString)
Dim countBaseType As String: countBaseType = UCase(Right(tenorString, 1))

maturityNum = Left(tenorString, nCharacter - 1)
Select Case countBaseType
    Case "D"
        countBase = "d"
    Case "W"
        countBase = "ww"
    Case "M"
        countBase = "m"
    Case "Y"
        countBase = "yyyy"
    Case Else
        
End Select

End Property

Public Property Get symbol() As String

Dim countBaseType As String

Select Case countBase
    Case "d"
        countBaseType = "D"
    Case "ww"
        countBaseType = "W"
    Case "m"
        countBaseType = "M"
    Case "yyyy"
        countBaseType = "Y"
End Select

symbol = maturityNum & countBaseType

End Property

Public Function toMaturityDate(valueDate As Date) As Date
toMaturityDate = DateAdd(countBase, maturityNum, valueDate)
End Function
