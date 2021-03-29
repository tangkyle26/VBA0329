Attribute VB_Name = "Module1"
Option Explicit

Public Function eoqdemo2(demand, holding, fixed) As Integer
eoqdemo2 = (2 * demand * fixed / holding) ^ (1 / 2)
End Function

