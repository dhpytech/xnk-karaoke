Attribute VB_Name = "SheetsControl"
Option Explicit
Function wsData() As Worksheet
  Set wsData = Sheet1
End Function
Function wsLocation() As Worksheet
  Set wsLocation = Sheet2
End Function
Function wsBangTinh() As Worksheet
  Set wsBangTinh = Sheet3
End Function
Function wsCache() As Worksheet
  Set wsCache = Sheet4
End Function
Function wsPhieuTT() As Worksheet
  Set wsPhieuTT = Sheet8
End Function
Function wsCTXD() As Worksheet
  Set wsCTXD = Sheet5
End Function
Function wsChungCu() As Worksheet
  Set wsChungCu = Sheet6
End Function
Function wsOTo() As Worksheet
  Set wsOTo = Sheet7
End Function
Sub HideAllSheet(dhBoolean As Boolean)
Dim num%
  If dhBoolean = True Then
    For num = 1 To 7
      If ThisWorkbook.Sheets(num).Visible = True Then
      Else
        ThisWorkbook.Sheets(num).Visible = True
      End If
    Next num
  Else
    For num = 1 To 7
      If ThisWorkbook.Sheets(num).Visible = True Then
        ThisWorkbook.Sheets(num).Visible = xlVeryHidden
      Else
      End If
    Next num
  End If
End Sub


