Attribute VB_Name = "ConditionControl"
Option Explicit
Function checkMaHs(dataStr$, arrData As Variant) As Boolean
Dim dhCheck As Boolean, numCol%, numLoop&
dhCheck = False
If LBound(arrData, 2) = 1 Then
  numCol = 1
Else
  numCol = 0
End If
  For numLoop = LBound(arrData, 1) To UBound(arrData, 1)
    If dataStr = CStr(arrData(numLoop, numCol)) Then
      dhCheck = True
    Else
    End If
  Next numLoop
checkMaHs = dhCheck
End Function

Function CheckEmpty(ParamArray Values())
Dim numLp As Integer
  For numLp = LBound(Values) To UBound(Values)
    If Values(numLp) = "" Then
      CheckEmpty = False
      Exit Function
    End If
  Next numLp
  
  CheckEmpty = True
End Function
