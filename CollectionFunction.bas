Attribute VB_Name = "CollectionFunction"
Option Explicit
Option Base 1

Public Function hasItem(col As Collection, key As Variant) As Boolean
Dim temp As Variant

On Error GoTo notHasKey:
temp = col(key)
hasItem = True
Exit Function

notHasKey:
hasItem = False
End Function

