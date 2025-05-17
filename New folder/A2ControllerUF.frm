VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A2ControllerUF 
   ClientHeight    =   8400.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14865
   OleObjectBlob   =   "A2ControllerUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A2ControllerUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HSBDS_Click()
  Call TatForm
  A3QlHoSoUF.Show
End Sub
Private Sub HSCANCEL_Click()
  Call TatForm
  A1LoginUF.Show
End Sub
Private Sub HSCC_Click()
  Call TatForm
  A3QlChungCuUF.Show
End Sub

Private Sub HSCTP_Click()
  Call TatForm
  CTPhiUF.Show
End Sub

Private Sub HSNS_Click()
  Call TatForm
  A4QlTaiKhoanUF.Show
End Sub
Private Sub HSOTo_Click()
  Call TatForm
  A3QlOtoUF.Show
End Sub
Private Sub HSSAVE_Click()
  ThisWorkbook.Save
End Sub
Private Sub HSSETTING_Click()
  Call TatForm
  A5QlCaiDatUF.Show
End Sub
Private Sub UserForm_Initialize()
Dim arrLabel, arrImageName
Dim numLoop%
  arrLabel = Array("HSLOGO", "HSBDS", "HSCC", "HSOTO", "HSNS", "HSSAVE", "HSSETTING", "HSCANCEL", "HSCTP")
  arrImageName = Array("logo.jpg", "bds.jpg", "chungcu.jpg", "oto.jpg", "person.jpg", "save.jpg", "setting.jpg", "cancel.jpg", "ctp.jpg")
  For numLoop = LBound(arrLabel, 1) To UBound(arrLabel, 1)
    Me.Controls(arrLabel(numLoop)).Picture = LoadPicture(dhRoot & "\Icons\" & arrImageName(numLoop))
  Next numLoop
End Sub
Private Sub TatForm()
Dim arrLabel
  arrLabel = Array("HSLOGO", "HSBDS", "HSCC", "HSOTO", "HSNS", "HSSAVE", "HSSETTING", "HSCANCEL")
  Call ClearMemory(Me, arrLabel)
  Erase arrLabel
  Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = False Then
  Cancel = True
  End If
End Sub
