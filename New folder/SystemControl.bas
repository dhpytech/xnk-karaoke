Attribute VB_Name = "SystemControl"
Declare PtrSafe Function GetSystemMetrics32 Lib "user32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Public Function widthRatio() As Integer
Dim excelWidth As Integer
    excelWidth = ThisWorkbook.Application.Width
    widthRatio = Int(excelWidth / 1163 * 100)
End Function
Public Function heightRatio() As Integer
Dim excelHeight As Integer
    excelHeight = ThisWorkbook.Application.Height
    heightRatio = Int(excelHeight / 623 * 100)
End Function
Public Function ShowScreenXDimensions()
   Dim X As Long
   Dim Y As Long
   ShowScreenXDimensions = GetSystemMetrics32(SM_CXSCREEN)
End Function
Public Function ShowScreenYDimensions()
   Dim X As Long
   Dim Y As Long
   ShowScreenYDimensions = GetSystemMetrics32(SM_CYSCREEN)
End Function
Public Function DhTechPath(folderName As String)
Dim dhFolder As String, wduser$, userPath$
  'Get User
  wduser = Environ("USERNAME")
  dhFolder = "C:\Users\" & wduser & "\Documents\" & folderName
  dhFolder = Replace(dhFolder, "\", "/", , , vbTextCompare)
' Check Path Or Create Path if Invalid
  If Dir(dhFolder, vbDirectory) = "" Then
    On Error Resume Next
    MkDir dhFolder
  Else
  End If
DhTechPath = dhFolder
End Function
Public Function screenRate()
  dhRatio = ShowScreenXDimensions() / ShowScreenYDimensions() * 9
  screenRate = CStr(dhRatio) & ":9"
End Function
Public Function checkPathExists(Path) As Boolean
Dim X As String
  On Error Resume Next
  X = GetAttr(Path) And 0
  If Err = 0 Then checkPathExists = True Else checkPathExists = False
End Function
Public Sub noticeToUser(notice As String, title As String)
Dim dhNotice As Object
  Set dhNotice = CreateObject("WScript.Shell")
  dhNotice.Popup notice, , title, 0 + 64
Set dhNotice = Nothing
End Sub
Public Function NumLockState() As Boolean
  NumLockState = Word.Application.NumLock
End Function
Function CheckWordRunning() As Boolean
  Dim X
  On Error Resume Next
  Set X = GetObject(, "Word.Application")
  CheckWordRunning = (Err = 0)
End Function
Public Function dhRoot() As String
Dim dhPath$, wduser$, OnedriveId$
wduser = Environ("USERNAME")
dhPath = ThisWorkbook.Path
If dhCheckFolderPath("D:\QLHS THAM DINH") = False Then
  If InStr(1, dhPath, "https://d.docs.live.net/", vbTextCompare) > 0 Then
    OnedriveId = Split(Replace(dhPath, "https://d.docs.live.net/", "", , , vbTextCompare), "/", , vbTextCompare)(0)
    dhPath = "C:\Users\" & wduser & "\OneDrive" & Replace(dhPath, "https://d.docs.live.net/" & OnedriveId, "", , , vbTextCompare)
    dhPath = Replace(dhPath, "/", "\", , , vbTextCompare)
  Else
  End If
Else
  dhPath = "D:\QLHS THAM DINH"
End If
dhRoot = dhPath
dhRoot = ThisWorkbook.Path
End Function
Public Function CMaHS(TenHS$) As String
  CMaHS = V2E(TenHS)
End Function
Public Sub SpeedUp(dhOK As Boolean)
If dhOK = True Then
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  ActiveSheet.DisplayPageBreaks = False
  Application.Calculation = xlCalculationManual
Else
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  ActiveSheet.DisplayPageBreaks = True
  Application.Calculation = xlCalculationAutomatic
End If
End Sub
Sub ClearMemory(dhUF As UserForm, Values As Variant)
Dim numLoop&
  For numLoop = LBound(Values, 1) To UBound(Values, 1)
    dhUF.Controls(Values(numLoop)).Picture = Nothing
  Next numLoop
End Sub
