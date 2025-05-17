Attribute VB_Name = "MathControl"
Option Explicit
Public Function getEndRow(ws As Worksheet, Col As String) As Long
  getEndRow = ws.Range(Col & Rows.Count).End(xlUp).Row
End Function
Public Function getDataRow(ws As Worksheet, Col As String) As Long
  getDataRow = ws.Range(Col & Rows.Count).End(xlUp).Row + 1
End Function
Public Function CnullToZero(cValue As Variant)
  If cValue = "" Or Not IsNumeric(cValue) Then
    CnullToZero = 0
  Else
    CnullToZero = cValue
  End If
End Function
Public Function CEm2None(dhValue As Variant, Optional ByVal dhReverse As Boolean = False)
If dhReverse = False Then
  If dhValue = "" Then
    CEm2None = "None"
  Else
    CEm2None = dhValue
  End If
Else
  If dhValue = "None" Then
    CEm2None = ""
  Else
    CEm2None = dhValue
  End If
End If
End Function
Public Function VNum(dhNum$) As String
Dim PhanNguyen$, PhanThapPhan$
If InStr(1, dhNum, ".", vbTextCompare) > 0 Then
  PhanNguyen = Replace(Split(dhNum, ".", , vbTextCompare)(0), ",", ".", , , vbTextCompare)
  PhanThapPhan = Split(dhNum, ".", , vbTextCompare)(1)
  VNum = Join(Array(PhanNguyen, PhanThapPhan), ",")
Else
  VNum = Replace(dhNum, ",", ".", , , vbTextCompare)
End If
End Function
