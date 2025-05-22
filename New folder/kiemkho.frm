VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} kiemkho 
   ClientHeight    =   10245
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   22470
   OleObjectBlob   =   "kiemkho.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "kiemkho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_SETTEXT = &HC
#If Win64 Then
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
#Else
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
#End If
Private Const WM_SETICON = &H80

Private Sub changelocation_Click()
    If Me.listkho.ListIndex = -1 Then
    CreateObject("WScript.Shell").Popup "Ki" & ChrW(7879) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n. Vui l" & ChrW(242) & "ng ch" & ChrW(7885) & "n ki" & ChrW(7879) & "n c" & ChrW(7847) & "n " & ChrW(273) & ChrW(7893) & "i v" & ChrW(7883) & " tr" & ChrW(237) & ".", , "Th" & ChrW(244) & "ng B" & ChrW(225) & "o!", 0 + 64
    Else
    
    location.package = Me.listkho.List(Me.listkho.ListIndex, 0)
    location.itemcode = Me.listkho.List(Me.listkho.ListIndex, 1)
    location.description = Me.listkho.List(Me.listkho.ListIndex, 2)
    location.ngay = Format(Date, "dd/mm/yyyy")
    location.quantity = Format(Me.listkho.List(Me.listkho.ListIndex, 6), "#,##0.000")
    
    location.userName = Sheet5.Range("E2")
    location.class = Me.listkho.List(Me.listkho.ListIndex, 8)
    
    location.location = Me.listkho.List(Me.listkho.ListIndex, 4)
    location.warehouse = Me.listkho.List(Me.listkho.ListIndex, 3)
    location.Show
    End If
End Sub

Private Sub cmdexportExcel_Click()
Dim wbexport As Workbook
Dim wsexport As Worksheet
Dim wbname As String, wduser As String
Set wbexport = Workbooks.Add
Set wsexport = wbexport.Sheets(1)


wduser = CreateObject("WScript.Network").userName
'MsgBox Environ("USERNAME")
wbname = "C:\Users\" & wduser & "\Downloads\lavergne-" & Format(Now(), "ddmmyyyy-hhmmss") & ".xlsx"
row = 0
    For num = 0 To Me.listkho.ListCount - 1
        If Me.listkho.List(num, 0) = "" Then
        Else
            row = row + 1
        End If
    Next num
'MsgBox row
If Me.khotonghopCheck = True Then
    If Me.listkho.List(0, 0) = "Item code" Then
    Else
        wsexport.Range("A1") = "Item code"
        wsexport.Range("B1") = "Description"
        wsexport.Range("C1") = "Product Code"
        wsexport.Range("D1") = "Class"
        wsexport.Range("E1") = "On-hand Qty"
        wsexport.Range("F1") = "Unit"
        wsexport.Range("G1") = "Warehouse"
        wsexport.Range("H1") = "Locationcode"
    End If
    For rowkiemkho = 1 To row
        wsexport.Range("A" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 0)
        wsexport.Range("B" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 1)
        wsexport.Range("C" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 2)
        wsexport.Range("D" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 4)
        wsexport.Range("E" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 7)
        wsexport.Range("F" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 8)
        wsexport.Range("G" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 9)
        wsexport.Range("H" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 10)
    Next rowkiemkho
    'Columns("A:H").EntireColumn.AutoFit
Else
    If Me.listkho.List(0, 0) = "Batch No - Box" Then
    Else
        wsexport.Range("A1") = "Batch-No-Box"
        wsexport.Range("B1") = "Itemcode"
        wsexport.Range("C1") = "Description"
        wsexport.Range("D1") = "Warehouse"
        wsexport.Range("E1") = "Locationcode"
        wsexport.Range("F1") = "Date"
        wsexport.Range("G1") = "On-hand Qty"
        wsexport.Range("H1") = "User Name"
        wsexport.Range("I1") = "Class"
    End If
    For rowkiemkho = 1 To row
        wsexport.Range("A" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 0)
        wsexport.Range("B" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 1)
        wsexport.Range("C" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 2)
        wsexport.Range("D" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 3)
        wsexport.Range("E" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 4)
        wsexport.Range("F" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 5)
        wsexport.Range("G" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 6)
        wsexport.Range("H" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 7)
        wsexport.Range("I" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 8)
    Next rowkiemkho
    'Columns("A:K").EntireColumn.AutoFit
End If
wbexport.SaveAs Filename:=wbname, FileFormat:=51
wbexport.Close SaveChanges:=True
End Sub

Private Sub cmdprint_Click()
Dim num As Long, row As Long, rowkiemkho As Long, endRow As Long, endrowfirst As Long
Dim wskiemkho As Worksheet
Set wskiemkho = Sheet10

Sheet10.Visible = xlSheetVisible
endrowfirst = wskiemkho.Range("A" & Rows.Count).End(xlUp).row
wskiemkho.Range("A1:I" & endrowfirst).clear
row = 0
    For num = 0 To Me.listkho.ListCount - 1
        If Me.listkho.List(num, 0) = "" Then
        Else
            row = row + 1
        End If
    Next num
'MsgBox row
If Me.khotonghopCheck = True Then
    If Me.listkho.List(0, 0) = "Item code" Then
    Else
        wskiemkho.Range("A1") = "Item code"
        wskiemkho.Range("B1") = "Description"
        wskiemkho.Range("C1") = "Product Code"
        wskiemkho.Range("D1") = "Class"
        wskiemkho.Range("E1") = "On-hand Qty"
        wskiemkho.Range("F1") = "Unit"
        wskiemkho.Range("G1") = "Warehouse"
        wskiemkho.Range("H1") = "Locationcode"
    End If
    For rowkiemkho = 1 To row
        wskiemkho.Range("A" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 0)
        wskiemkho.Range("B" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 1)
        wskiemkho.Range("C" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 2)
        wskiemkho.Range("D" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 4)
        wskiemkho.Range("E" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 7)
        wskiemkho.Range("F" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 8)
        wskiemkho.Range("G" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 9)
        wskiemkho.Range("H" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 10)
    Next rowkiemkho
    Columns("A:H").EntireColumn.AutoFit
    endRow = wskiemkho.Range("A" & Rows.Count).End(xlUp).row
    wskiemkho.PageSetup.PrintArea = "$A$1:$H$" & endRow
Else
    If Me.listkho.List(0, 0) = "Batch No - Box" Then
    Else
        wskiemkho.Range("A1") = "Batch-No-Box"
        wskiemkho.Range("B1") = "Itemcode"
        wskiemkho.Range("C1") = "Description"
        wskiemkho.Range("D1") = "Warehouse"
        wskiemkho.Range("E1") = "Locationcode"
        wskiemkho.Range("F1") = "Date"
        wskiemkho.Range("G1") = "On-hand Qty"
        wskiemkho.Range("H1") = "User Name"
        wskiemkho.Range("I1") = "Class"
    End If
    For rowkiemkho = 1 To row
        wskiemkho.Range("A" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 0)
        wskiemkho.Range("B" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 1)
        wskiemkho.Range("C" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 2)
        wskiemkho.Range("D" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 3)
        wskiemkho.Range("E" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 4)
        wskiemkho.Range("F" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 5)
        wskiemkho.Range("G" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 6)
        wskiemkho.Range("H" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 7)
        wskiemkho.Range("I" & rowkiemkho + 1) = Me.listkho.List(rowkiemkho - 1, 8)
    Next rowkiemkho
    Columns("A:K").EntireColumn.AutoFit
    endRow = wskiemkho.Range("A" & Rows.Count).End(xlUp).row
    wskiemkho.PageSetup.PrintArea = "$A$1:$I$" & endRow
End If

With wskiemkho.PageSetup
    .PrintTitleRows = "$1:$1"
    .PrintTitleColumns = ""
    
    .LeftMargin = Application.InchesToPoints(0.4)
    .RightMargin = Application.InchesToPoints(0.4)
    .TopMargin = Application.InchesToPoints(0.4)
    .BottomMargin = Application.InchesToPoints(0.4)
    .HeaderMargin = Application.InchesToPoints(0)
    .FooterMargin = Application.InchesToPoints(0)
    .PrintComments = xlPrintNoComments
    .PrintQuality = 600
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlLandscape
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 100
    .PrintErrors = xlPrintErrorsDisplayed
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .ScaleWithDocHeaderFooter = True
End With

wskiemkho.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
wskiemkho.Range("A1:I" & endRow).clear
Sheet10.Visible = xlSheetVeryHidden
End Sub

Private Sub khochitiet_Click()
    Call updaterepodetail(9)
    Me.changelocation.Visible = True
End Sub

Private Sub khotonghop_Click()
    Call updaterepo(17)
    Me.changelocation.Visible = False
End Sub

Private Sub listkho_Click()

End Sub

Private Sub LogoutButton_Click()
    Unload Me
    adminform.Show
End Sub

Private Sub searchtype_Change()
    Call searchvattu_Change
End Sub

Private Sub searchvattu_Change()
Dim arr_infor(), arr_result(), infor, result As Long, textSearch As String
Dim repo As Worksheet, repodetail As Worksheet

Set repo = Sheet2
Set repodetail = Sheet7
textSearch = LCase(Me.searchvattu.text)

If Me.khotonghopCheck = True Then
    arr_infor = repo.Range("repo").value 'gan ma nguon
    ReDim arr_result(1 To UBound(arr_infor, 1), 1 To 17)
    For infor = 1 To UBound(arr_infor, 1)
        If Me.searchtype = "All" Then
            If LCase(arr_infor(infor, 1)) Like "*" & textSearch & "*" Or _
            LCase(arr_infor(infor, 2)) Like "*" & textSearch & "*" Or _
            LCase(arr_infor(infor, 11)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 17 Step 1
                    If i = 8 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 8) = Format(arr_infor(infor, 8), "#,##0.000")
            End If
        ElseIf Me.searchtype = "Mã SP" Then
            If LCase(arr_infor(infor, 1)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 17 Step 1
                    If i = 8 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 8) = Format(arr_infor(infor, 8), "#,##0.000")
            End If
        ElseIf Me.searchtype = "Kho" Then
            If LCase(arr_infor(infor, 11)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 17 Step 1
                    If i = 8 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 8) = Format(arr_infor(infor, 8), "#,##0.000")
            End If
        Else
        End If
        
    Next infor
Else
    arr_infor = repodetail.Range("repodetail").value 'gan ma nguon
    ReDim arr_result(1 To UBound(arr_infor, 1), 1 To 9)
    For infor = 1 To UBound(arr_infor, 1)
        If Me.searchtype = "All" Then
            If LCase(arr_infor(infor, 2)) Like "*" & textSearch & "*" Or _
            LCase(arr_infor(infor, 5)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 9 Step 1
                    If i = 7 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 7) = Format(arr_infor(infor, 7), "#,##0.000")
            End If
        ElseIf Me.searchtype = "Mã SP" Then
            If LCase(arr_infor(infor, 2)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 9 Step 1
                    If i = 7 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 7) = Format(arr_infor(infor, 7), "#,##0.000")
            End If
        ElseIf Me.searchtype = "Kho" Then
            If LCase(arr_infor(infor, 5)) Like "*" & textSearch & "*" Then
            result = result + 1
                For i = 1 To 9 Step 1
                    If i = 7 Then
                    Else
                    arr_result(result, i) = arr_infor(infor, i)
                    End If
                Next i
                arr_result(result, 7) = Format(arr_infor(infor, 7), "#,##0.000")
            End If
        Else
        End If
        
    Next infor
End If
Me.listkho = ""
Me.listkho.clear
Me.listkho.List = arr_result 'Gan ket qua lai
End Sub

Private Sub UserForm_Initialize()
Dim hWnd&
Dim strIconPath As String
Dim lngIcon As Long
Dim lngHWnd As Long
Dim wsuser As Worksheet, wsrepo As Worksheet
Dim wF As Variant
Set wsuser = Sheet5
Set wsrepo = Sheet2
    ' Change to the path and filename of an icon file
    strIconPath = "C:\Program Files\Inventory Management\Icons\logo.ico"
    ' Get the icon from the source
    lngIcon = ExtractIcon(0, strIconPath, 0)
    ' Get the window handle of the userform
    lngHWnd = FindWindow("ThunderDFrame", Me.Caption)
    'Set the big (32x32) and small (16x16) icons
    SendMessage lngHWnd, WM_SETICON, True, lngIcon
    SendMessage lngHWnd, WM_SETICON, False, lngIcon
    SendMessage lngHWnd, WM_SETICON, False, lngIcon '16x16
    
    hWnd = FindWindow("ThunderDFrame", Caption)
    DefWindowProcW hWnd, WM_SETTEXT, 0, StrPtr("KI" & ChrW(7874) & "M K" & ChrW(202) & " KHO")
    ' Set background
    'Me.commandframe.Picture = LoadPicture("C:\Program Files\Inventory Management\Icons\subbackground1.jpg")
    'Me.framelistvattu.Picture = LoadPicture("C:\Users\Vo Dang Huan\Desktop\Tam\subbackground1.jpg")
    
    
    'Me.repoListView.List = wsrepo.Range("repo").Value
    Call updaterepo(17)
    Me.khochitietCheck = False
    Me.khotonghopCheck = True
    Me.CheckFrame.Enabled = False
    
    With Me.searchtype
    .AddItem "All"
    .AddItem "Kho"
    .AddItem "Mã SP"
    .text = "All"
    End With
    
    Me.StartUpPosition = 2
    'Resize UF
    With Me
    .Zoom = Int(Application.WorksheetFunction.Min(widthRatio, heightRatio) * 0.95)
    .Width = .Width * .Zoom / 100
    .Height = .Height * .Zoom / 100
    End With
    
    Me.changelocation.Visible = False
End Sub

Private Sub updaterepo(numcol As Integer)
Dim original(), resultlist()
Dim numuser As Integer, result As Integer
Set wsrepo = Sheet2

original = wsrepo.Range("repo").value
ReDim resultlist(1 To UBound(original, 1), 1 To 17)
    For numuser = 1 To UBound(original, 1)
        For i = 1 To 17 Step 1
            If i = 8 Then
            Else
                resultlist(numuser, i) = original(numuser, i)
            End If
        Next i
        resultlist(numuser, 8) = Format(original(numuser, 8), "#,##0.000")
    Next numuser
Me.khochitietCheck = False
Me.khotonghopCheck = True
Me.listkho.Height = 420

Me.listkho.clear
Me.listkho.ColumnCount = numcol
Me.listkho.ColumnWidths = "130;320;0;0;170;0;0;80;40;70;70;0;0;0;0;0;0"
Me.listkho.List = resultlist
End Sub

Public Sub updaterepodetail(numcol As Integer)
Dim original(), resultlist()
Dim numuser As Integer, result As Integer
Set wsrepodetail = Sheet7

original = wsrepodetail.Range("repodetail").value
ReDim resultlist(1 To UBound(original, 1), 1 To 9)
    For numuser = 1 To UBound(original, 1)
        For i = 1 To 9 Step 1
            If i = 7 Then
            Else
                resultlist(numuser, i) = original(numuser, i)
            End If
        Next i
        resultlist(numuser, 7) = Format(original(numuser, 7), "#,##0.000")
    Next numuser
Me.khochitietCheck = True
Me.khotonghopCheck = False
Me.listkho.Height = 420
Me.listkho.clear
Me.listkho.ColumnCount = numcol
Me.listkho.ColumnWidths = "110;110;250;70;70;70;70;70;70"
Me.listkho.List = resultlist

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = False Then
Cancel = True
End If
End Sub

