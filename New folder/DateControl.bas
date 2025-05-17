Attribute VB_Name = "DateControl"
Option Explicit
Public Function dhDate(dateString As String, Optional ByVal style As String = "dd/mm/yyyy") As Date
Dim dateArr As Variant
Dim dhDay%, dhMonth%, dhYear%
  If dateString Like "??/??/????" Then
    dateArr = Split(dateString, "/", , vbTextCompare)
    dhDay = Int(dateArr(0)): dhMonth = Int(dateArr(1)): dhYear = Int(dateArr(2))
    dhDate = DateSerial(dhYear, dhMonth, dhDay)
  ElseIf dateString Like "??-??-????" Then
    dateArr = Split(dateString, "-", , vbTextCompare)
    dhDay = Int(dateArr(0)): dhMonth = Int(dateArr(1)): dhYear = Int(dateArr(2))
    dhDate = DateSerial(dhYear, dhMonth, dhDay)
  Else
    dhDate = Date
  End If
End Function
Public Function dhDate2Str(Ngay As Date)
  dhDate2Str = "Ngày " & Format(Ngay, "dd") & " Tháng " & Format(Ngay, "mm") & " N" & ChrW(259) & "m " & Format(Ngay, "yyyy")
End Function
