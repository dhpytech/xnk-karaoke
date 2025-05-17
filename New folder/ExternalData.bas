Attribute VB_Name = "ExternalData"
Option Explicit
Public Function UserData() As Variant
Dim wbUser As Workbook, wsUser As Worksheet
Dim dhPath$, dhCheck As Boolean
Application.ScreenUpdating = False
dhPath = dhRoot & "\Config.xlsx"
dhCheck = checkPathExists(dhPath)
If dhCheck = True Then
  Set wbUser = Application.Workbooks.Open(dhPath, False, , , "Dhtech")
  Application.Wait (Now + TimeValue("0:00:2"))
  Set wsUser = wbUser.Sheets(1)
  UserData = wsUser.Range("UserList").value
  wbUser.Close SaveChanges:=False
  Set wbUser = Nothing
  Set wsUser = Nothing
Else
  UserData = Array("Invalid", "Path Is Not Exists")
End If
Application.ScreenUpdating = True
End Function
Private Function UserAuthPub(UsName$, UsPass$) As Variant
Dim arrAuth(1 To 3)
Dim arrUser As Variant
arrUser = UserData
Application.ScreenUpdating = False
Dim UsCheck As Boolean, UsPosition$
Dim numUs%
UsCheck = False: UsPosition = "Invalid"
If UBound(arrUser, 1) = 0 Then
Else
  For numUs = LBound(arrUser, 1) To UBound(arrUser, 1)
    If UsName = arrUser(numUs, 1) And UsPass = arrUser(numUs, 2) Then
      UsCheck = True: UsPosition = arrUser(numUs, 3)
      Exit For
    Else
    End If
  Next numUs
End If
UserAuthPub = Array(UsCheck, UsName, UsPosition, arrUser)
Erase arrUser
Application.ScreenUpdating = True
End Function
