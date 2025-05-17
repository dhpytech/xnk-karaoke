Attribute VB_Name = "MouseControl"
#If Win64 Then
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If
  Private Type POINTAPI
    X As Long
    Y As Long
  End Type
Public Function getXmouse()
Dim Pt As POINTAPI, mWnd As Long
  GetCursorPos Pt 'L?y gi� tr? X v� Y t?i v? tr� con tr? chu?t
  mWnd = WindowFromPoint(Pt.X, Pt.Y) 'Luu bi?n mWnd = v?i Handle t?i d?i tu?ng c� t?a d? X v� Y b?ng h�m WindowFromPoint
  'MsgBox "handle= " & mWnd 'Hi?n Msgbox gi� tr? Handle c?a d?i tu?ng d� ch?n
  getXmouse = Pt.X
End Function
Public Function getYmouse()
Dim Pt As POINTAPI, mWnd As Long
  GetCursorPos Pt 'L?y gi� tr? X v� Y t?i v? tr� con tr? chu?t
  mWnd = WindowFromPoint(Pt.X, Pt.Y) 'Luu bi?n mWnd = v?i Handle t?i d?i tu?ng c� t?a d? X v� Y b?ng h�m WindowFromPoint
  'MsgBox "handle= " & mWnd 'Hi?n Msgbox gi� tr? Handle c?a d?i tu?ng d� ch?n
  getYmouse = Pt.Y
End Function
