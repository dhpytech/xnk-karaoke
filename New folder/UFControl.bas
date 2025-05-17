Attribute VB_Name = "UFControl"
'Khoi Tao Tab cho UserForm
Public Sub tabInit(UF As UserForm, arrName As Variant, Optional dhClear As Boolean = True)
Dim numLoop%, numTab%
  For numLoop = LBound(arrName) To UBound(arrName)
    UF.Controls(arrName(numLoop)).TabIndex = numLoop
    If TypeName(UF.Controls(arrName(numLoop))) = "CheckBox" Then
      UF.Controls(arrName(numLoop)).value = False
    ElseIf TypeName(UF.Controls(arrName(numLoop))) = "TextBox" Then
      If dhClear = True Then
        UF.Controls(arrName(numLoop)).Text = ""
      End If
    Else
    End If
  Next numLoop
End Sub
