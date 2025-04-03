
' ***************************************************************
' ThisWorkbook
'
' ***************************************************************


Private Sub Workbook_Open()
    Sheets("ss").Activate
    ISIT_FIRST = True
    
     Call clearRowA
     
End Sub

' ***************************************************************
' Sheet5_ss(ss)
'
' ***************************************************************

Private Sub combobox_initialize()

'    Dim tbl As ListObject
'    Dim tableNAME, shNAME As String
'
'    Dim cell As Range
'    Dim i As Integer
'    Dim isFirst As Boolean: isFirst = True
'
'
'    If ISIT_FIRST Then
'        comboAREA.Clear
'
'        If chkboxJIYEOL.Value = True Then
'            tableNAME = "tableJIYEOL"
'            shNAME = "ref1"
'        Else
'            tableNAME = "tableCNU"
'            shNAME = "ref"
'        End If
'
'        Set tbl = Sheets(shNAME).ListObjects(tableNAME)
'
'        i = 0
'        For Each cell In tbl.HeaderRowRange.Cells
'            If isFirst Then
'                isFirst = False
'                GoTo NEXT_ITER
'            End If
'
'             comboAREA.AddItem cell.Value
'NEXT_ITER:
'        Next cell
'    Else
'        ISIT_FIRST = False
'    End If
End Sub


Private Sub CommandButton5_Click()
    UserForm_survey.Show
End Sub


Private Sub CommandButton6_Click()
    Call water_GenerateCopy.Finallize
End Sub

Private Sub Worksheet_Activate()
   
End Sub

'Private Sub chkboxJIYEOL_Click()
'    ISIT_FIRST = True
'    Call combobox_initialize
'    ISIT_FIRST = False
'End Sub


Private Sub comboAREA_DropButtonClick()
    'Call combobox_initialize
End Sub

Private Sub comboAREA_GotFocus()
   'Call combobox_initialize
End Sub


Private Sub comboAREA_Change()
    ' Dim selectedHeader As String
    ' selectedHeader = comboAREA.Value
    ' Range("S21").Value = selectedHeader
End Sub


Private Sub CommandButton1_Click()
    Call insertRow
End Sub

Private Sub CommandButton2_Click()
    ' ���ϼ� �̿���� ��������ǥ
    ' Groundwater Availability Field Survey Sheet
    
    Call mod_MakeFieldList.MakeFieldList
    Sheets("ss").Activate
    
End Sub

Private Sub CommandButton3_Click()
    Popup_MessageBox ("Calculation Compute Q .... ")
    Call water_q.ComputeQ
    Sheets("ss").Activate
End Sub

Private Sub CommandButton4_Click()

   If Sheets("ref").Visible Then
        Sheets("ref").Visible = False
        Sheets("ref1").Visible = False
    Else
        Sheets("ref").Visible = True
        Sheets("ref1").Visible = True
    End If
    
End Sub

Private Sub CommandButtonCopy_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButtonDelete_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButtonInitialClear_Click()
 Call SubModuleInitialClear
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' Sheet2_aa(aa)
'
' ***************************************************************



Private Sub CommandButton1_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButton2_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButton3_Click()
    Call insertRow
End Sub

Private Sub CommandButton4_Click()
    Call ComputeQ
    Sheets("aa").Activate
End Sub

Private Sub CommandButton5_Click()
    ' ���ϼ� �̿���� ��������ǥ
    ' Groundwater Availability Field Survey Sheet
    
    Call mod_MakeFieldList.MakeFieldList
    Sheets("aa").Activate
End Sub

Private Sub CommandButton6_Click()
    Call Finallize
End Sub

Private Sub CommandButtonInitialClear_Click()
 Call SubModuleInitialClear
End Sub


Private Sub Worksheet_Activate()
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' Sheet3_ii(ii)
'
' ***************************************************************



Private Sub CommandButton1_Click()
    Call MainMoudleGenerateCopy
End Sub


Private Sub CommandButton2_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButton3_Click()
    Call insertRow
End Sub

Private Sub CommandButton4_Click()
 Call SubModuleInitialClear
End Sub

Private Sub CommandButton5_Click()
    Call Finallize
End Sub

Private Sub Worksheet_Activate()
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' water_q
'
' ***************************************************************


Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double
Public ISIT_FIRST As Boolean

Public Enum SS_VALUE
    svGAJUNG = 1
    svILBAN = 2
    svSCHOOL = 3
    svGONGDONG = 4
    svMAEUL = 5
End Enum

Public Enum AA_VALUE
    avJEONJAK = 1
    avDAPJAK = 2
    avWONYE = 3
    avCOW = 4
    avPIG = 5
    avCHICKEN = 6
End Enum

Function CheckBoxFind(objNAME As String) As MSForms.CheckBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myCheckBox As MSForms.CheckBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myCheckBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.CheckBox Then
            If obj.Name = objNAME Then
                Set myCheckBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myCheckBox Is Nothing) Then
        ' found
        Set CheckBoxFind = myCheckBox
    Else
        ' not found
        Set CheckBoxFind = Nothing
    End If
End Function

Function ComboBoxFind(objNAME As String) As MSForms.ComboBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myComboBox As MSForms.ComboBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myComboBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.ComboBox Then
            If obj.Name = objNAME Then
                Set myComboBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myComboBox Is Nothing) Then
        ' found
        Set ComboBoxFind = myComboBox
    Else
        ' not found
        Set ComboBoxFind = Nothing
    End If
End Function


Function TextBoxFind(objNAME As String) As MSForms.TextBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myTextBox As MSForms.TextBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myTextBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.TextBox Then
            If obj.Name = objNAME Then
                Set myTextBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myTextBox Is Nothing) Then
        ' found
        Set TextBoxFind = myTextBox
    Else
        ' not found
        Set TextBoxFind = Nothing
    End If
End Function



Function is_Jiyeol(ByVal area As String) As Boolean
    Dim tbl As ListObject
    Dim headerRowArray() As Variant
    
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If headerRowArray(1, i) = area Then
            is_Jiyeol = True
            Exit Function
        End If
    Next i
    
    is_Jiyeol = False
End Function


Sub initialize()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
        
    If is_Jiyeol(TextBox_AREA.Value) Then
        Call initialize_JIYEOL(TextBox_AREA.Value)
    Else
        Call initialize_CNU(TextBox_AREA.Value)
    End If
       
End Sub


Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


' �������
Sub ComputeQ()
    Dim i As Integer
    Dim lastrow As Long

    Call initialize
    
    Sheets("ss").Activate
    lastrow = lastRowByKey("A1")
    
    For i = 2 To lastrow
        Cells(i, "L").Value = ss_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
    
    Sheets("aa").Activate
    lastrow = lastRowByKey("A1")
    
    For i = 2 To lastrow
        Cells(i, "L").Value = aa_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
End Sub


Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double

    If qhp = 0 Then
        Exit Function
    End If

    '���� �ó���
    If CheckSubstring(strPurpose, "��") Then
        ss_water = qhp * 0.01
        Exit Function
    End If
    
    ' �Ϲݿ�
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' ��Ÿ
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' ���Ȱ���
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' û�ҿ�
    If CheckSubstring(strPurpose, "û") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    '���̻����
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    ' �����
    If CheckSubstring(strPurpose, "����") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' �������ÿ�
    If CheckSubstring(strPurpose, "����") Then
        ss_water = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    ' �ι�����
    If CheckSubstring(strPurpose, "�ι�") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' �б���
    If CheckSubstring(strPurpose, "�б�") Then
        ss_water = Round(SS(svSCHOOL, 1) + npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
    
    ' �����
    If CheckSubstring(strPurpose, "����") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' �ҹ��
    If CheckSubstring(strPurpose, "�ҹ�") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
   ss_water = 900
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double
    'nhead - ������ �μ� ....

    If qhp = 0 Then
        Exit Function
    End If

    ' ���ۿ�
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' ���ۿ�
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    ' ���Ȱ���
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    '����
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
    ' ��Ÿ
    If CheckSubstring(strPurpose, "��Ÿ") Then
        aa_water = Round(AA(avDAPJAK, 1) + nhead * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
   aa_water = 900
End Function











' ***************************************************************
' water_GenerationCopy
'
' ***************************************************************

Option Explicit

Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFindAll() As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


' ���⼭ �˻���, "AA" ���� ��쿡��, ���� �ؽ�Ʈ ����Ÿ �Ӹ� �ƴ϶� ...
' =SUMIF(SS_INSIDE_AREA,"O",$L$2:$L$2) �� ��쿡�� ó��, �̷��͵� �˻��� �Ǳ⿡,
' �ϴ��� Ctrl+F �� �˻��� �غ��°��� ��õ�Ѵ�.
' �̰���, ������ �˻��� �̿��ؼ�, ��ġ�ϴ� �Լ��̱� �����̴�.

Private Function lastRowByFind(ByVal str As String) As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find(str, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


Private Sub DoCopy(lastrow As Long)
    Range("F2:H" & lastrow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' ����
    Range("L2:L" & lastrow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastrow).Select
    Selection.Copy
    
    Range("r2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
End Sub



' return letter of range ...
Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
    
    Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function


' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA
Sub ToggleOX()
    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    

    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Debug.Print Alpha_Column(ActiveCell)
    
    
    '2024-12-25, Add Compute Q
    If activeCellColumn = "L" Then
        Popup_MessageBox ("Calculation Compute Q .... ")
        Call water_q.ComputeQ
        Sheets("ss").Activate
    End If
    
    
    
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    

    If activeCellColumn = "B" Then
        If ActiveCell.Value = "�Ű��" Then
            ActiveCell.Value = "�㰡��"
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            ActiveCell.Value = "�Ű��"
             With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End If
    
    If activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "D" & Range(cp).row & ":D" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
    If activeCellColumn = "C" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "C" & Range(cp).row & ":C" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
       
    ' 2024,12,22 - toggle address format
    If activeCellColumn = "M" Then
      Call AddressReset(ActiveSheet.Name)
    End If
    
    
    If ActiveSheet.Name = "ss" And activeCellColumn = "K" Then
        UserForm_SS.Show
    End If
    
    If ActiveSheet.Name = "aa" And activeCellColumn = "K" Then
        UserForm_AA.Show
    End If
    
    If ActiveSheet.Name = "ii" And activeCellColumn = "K" Then
        UserForm_II.Show
    End If
End Sub


' Ctrl+R , Transfer Well Data
' =D2&" "&E2&" ����"
Sub ToggleAddressFormatString()

    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    Dim MainSheet, TargetSheet As String
    
    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
 
    If lastrow = 1048577 Or Range("E" & (lastrow - 1)).Value = "��Ȱ��" Then
        lastrow = 2
    End If
    
    
    Range("E" & lastrow).Select
    ActiveSheet.Paste
    
       
    AddressReset (MainSheet)
    AddressReset (TargetSheet)
    Range("G7").Select

End Sub

' Ctrl+R , Transfer Well Data
' =D2&" "&E2&" ����"
Sub TransferWellData()

    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    Dim MainSheet, TargetSheet As String
    
    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    MainSheet = ActiveSheet.Name
    
    If MainSheet = "aa" Then
        TargetSheet = "ss"
    ElseIf MainSheet = "ss" Then
        TargetSheet = "aa"
    Else
        Exit Sub
    End If
    
    fillRange = "E" & row & ":J" & row
    Range(fillRange).Select
    Selection.Cut
    
    Sheets(TargetSheet).Activate
    lastrow = lastRowByKey("E1") + 1
    
    If lastrow = 1048577 Or Range("E" & (lastrow - 1)).Value = "��Ȱ��" Then
        lastrow = 2
    End If
    
    
    Range("E" & lastrow).Select
    ActiveSheet.Paste
    
       
    AddressReset (MainSheet)
    AddressReset (TargetSheet)
    Range("G7").Select


End Sub

' =D2&" "&E2&" ����"
Sub AddressReset(Optional ByVal shName As String = "option")
    Dim lastrow As Long
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = shName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    
    If Not sheetExists Then
        shName = ActiveSheet.Name
    End If
    
    
    Sheets(shName).Activate
    
    lastrow = lastRowByKey("M2")
    
    If CheckSubstring(Range("M2"), "����") Then
        Range("M2").Formula = "=D2&"" ""&E2"
    Else
        Range("M2").Formula = "=D2&"" ""&E2&"" ����"" "
    End If
    
    
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M" & lastrow)
    Range("M2").Select
  
End Sub

Sub test()
    Dim lastrow As Long
    
    
    lastrow = lastRowByKey("E1") + 1
        
   If lastrow = 1048577 Or ActiveCell.Value = "��Ȱ��" Then
        lastrow = 2
    End If
    
    Range("o2").Value = "ll"
End Sub


Sub MainMoudleGenerateCopy()
    Dim lastrow As Long
        
    lastrow = lastRowByKey("A1")
    Call DoCopy(lastrow)
End Sub


Sub SubModuleInitialClear()
    Dim lastrow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastrow = lastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastrow).Select
    Selection.ClearContents
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    
    If lastrow >= 23 Then
        Rows("23:" & lastrow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    If (ActiveSheet.Name = "ii") Then
        Range("l2").Value = 0
    End If
    
    Range("m2").Select
End Sub


Sub Finallize()
    Dim lastrow As Long
    Dim delStartRow, delEndRow, delAddressStart As Long
    Dim userChoice As VbMsgBoxResult
    
    lastrow = lastRowByKey("A1")
    delStartRow = lastRowByKey("D1") + 1
    
    delAddressStart = lastRowByKey("E1") + 1
    
    
    Select Case ActiveSheet.Name
    
        Case "ss"
            delEndRow = lastRowByFind("����") - 4
            
        Case "aa"
            delEndRow = lastRowByFind("������") - 4
        
        Case "ii"
            delEndRow = lastRowByFind("������") - 6
    
    End Select
    
    '
    'if q is 0 then this section is not have water resource so clear next well
    '
    If Range("L2").Value = 0 Then
        delStartRow = 3
        delEndRow = lastRowByKey("L1")
    Else
        delStartRow = delAddressStart
    End If
    
    
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    If delStartRow = 1048577 Or lastrow = 2 Or (delEndRow - delStartRow <= 2) Then
        Exit Sub
    Else
        Rows(delStartRow & ":" & delEndRow).Select
        Selection.Delete Shift:=xlUp
        Range("A2").Select
    End If
      
End Sub

Sub SubModuleCleanCopySection()
    Dim lastrow As Long
        
    lastrow = lastRowByKey("A1")
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    Range("P14").Select
End Sub


' 2023/4/19 - copy modify
'2024/12/25 -- add short cut (Ctrl+i)

Sub insertRow()
    Dim lastrow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastrow = lastRowByRowsCount("A")
    
    Rows(CStr(lastrow + 1) & ":" & CStr(lastrow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1")
    j = i + AddingRowCount
    
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "K" & i & ":M" & i
    selection_target = "K" & i & ":M" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("S" & i).Select
    Selection.AutoFill Destination:=Range("S" & i & ":S" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
End Sub





' ***************************************************************
' mod_initialize_setting
'
' ***************************************************************


Option Explicit



Private Sub init_nonsan()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.42
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_daejeon()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.43

    SS(svILBAN, 1) = 2.119
    SS(svILBAN, 2) = 0.021
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 7.13
    SS(svGONGDONG, 2) = 0.001
    
    SS(svMAEUL, 1) = 6.463
    SS(svMAEUL, 2) = 0.178
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub


Private Sub init_yeonki()

   
    SS(svGAJUNG, 1) = 0.265
    SS(svGAJUNG, 2) = 0.181
    SS_CITY = 2.75

    SS(svILBAN, 1) = 3.521
    SS(svILBAN, 2) = 0.011
    
    SS(svSCHOOL, 1) = 11.687
    SS(svSCHOOL, 2) = 0.007
    
    SS(svGONGDONG, 1) = 0.265
    SS(svGONGDONG, 2) = 0.181
    
    SS(svMAEUL, 1) = 7.287
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub

Private Sub init_boryoung()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.36
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_dangjin()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.59
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_yesan()

   
    SS(svGAJUNG, 1) = 0.265
    SS(svGAJUNG, 2) = 0.181
    SS_CITY = 2.34
    
    SS(svILBAN, 1) = 3.521
    SS(svILBAN, 2) = 0.011
    
    SS(svSCHOOL, 1) = 11.687
    SS(svSCHOOL, 2) = 0.007
    
    SS(svGONGDONG, 1) = 0.265
    SS(svGONGDONG, 2) = 0.181
    
    SS(svMAEUL, 1) = 7.287
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub



Private Sub init_sejong()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.57

    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 7.13
    SS(svGONGDONG, 2) = 0.001
    
    SS(svMAEUL, 1) = 6.463
    SS(svMAEUL, 2) = 0.178
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub


Public Enum LC_COMBOBOX
    lcDAEJEON = 1
    lcJIYEOL = 2
End Enum

Public IS_FIRST_LOAD As Boolean

Private Sub OptionButton_DAEJEON_Click()

    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub


Private Sub OptionButton_JIYEOL_Click()
    
    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub

Sub PutDataToASheet(ByVal sh As String, ByVal table As String, ByVal area As String, SurveyData As Variant)
    Dim tbl As ListObject
    Dim cell As Range
    Dim i As Integer: i = 1
    
    Set tbl = Sheets(sh).ListObjects(table)
        
    For Each cell In tbl.ListColumns(area).DataBodyRange.Cells
        cell.Value = SurveyData(i)
        i = i + 1
    Next cell
    
End Sub

Function getSurveyData() As Variant
    Dim values() As Variant
    Dim i As Integer: i = 1
    Dim ctl As Control

    ReDim values(1 To 23)
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            values(i) = ctrl.Value
            i = i + 1
        End If
    Next ctrl
    
    getSurveyData = values
    
End Function



Private Sub CommandButton_Insert_Click()

    Dim values As Variant
    Dim area As String
    
    values = getSurveyData()

    area = ComboBox_AREA.Value
    
    If area = "" Then
        area = Default
    End If
    
    
    If OptionButton_JIYEOL.Value Then
        Call PutDataToASheet("ref1", "tableJIYEOL", area, values)
    Else
        Call PutDataToASheet("ref", "tableCNU", area, values)
    End If
    
    Call PutText(area)
    
    Unload Me
    
End Sub


Sub PutText(area As String)
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    TextBox_AREA.Value = area
End Sub

' ***************************************************************
' UserForm_Survey
'
' ***************************************************************



Private Sub CommandButton_LOAD_Click()
    Call LoadSurveyData(ComboBox_AREA.Value)
End Sub


Private Sub ComboBox_AREA_Change()
 Call LoadSurveyData(ComboBox_AREA.Value)
End Sub

Sub Initialize_Setting()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    Debug.Print "TextBox_AREA.Value", "'" & TextBox_AREA.Value & "'"
    
    If is_Jiyeol(TextBox_AREA.Value) Then
        OptionButton_JIYEOL.Value = True
    Else
        OptionButton_DAEJEON.Value = True
    End If
    
    ' Call LoadComboBox
    ' OptionButton.Value = True set is triggered clicked event
    
    ComboBox_AREA.Value = TextBox_AREA.Value
    LoadSurveyData (TextBox_AREA.Value)
    
End Sub


Sub LoadSurveyData(area As String)
    Dim tbl As ListObject
    Dim values() As Variant
    
    
    If OptionButton_JIYEOL.Value Then
        Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    Else
        Set tbl = Sheets("ref").ListObjects("tableCNU")
    End If
    
    
    If area = "" Or area = "0" Then
        area = Default
        Exit Sub
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
    Dim i As Integer: i = 1
        
    Dim ctl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ' MsgBox "Found a TextBox with the name: " & ctrl.NAME
            ctrl.Value = values(i, 1)
            i = i + 1
        End If
    Next ctrl
    
End Sub


Sub LoadComboBox()
    Dim tbl As ListObject
    Dim tableNAME, shName As String
    Dim headerRowArray() As Variant
    
    ComboBox_AREA.Clear
    
    If OptionButton_JIYEOL.Value Then
        tableNAME = "tableJIYEOL"
        shName = "ref1"
    Else
        tableNAME = "tableCNU"
        shName = "ref"
    End If
    
    Set tbl = Sheets(shName).ListObjects(tableNAME)

    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    Dim isFirst As Boolean: isFirst = True
    
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If isFirst Then
            isFirst = False
            GoTo NEXT_LOOP
        End If
        
        ComboBox_AREA.AddItem headerRowArray(1, i)
        
NEXT_LOOP:
    Next i
End Sub



Function is_Jiyeol(ByVal area As String) As Boolean

    Dim tbl As ListObject
    Dim headerRowArray() As Variant
    
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If headerRowArray(1, i) = area Then
            is_Jiyeol = True
            Exit Function
        End If
    Next i
    
    is_Jiyeol = False

End Function


Private Sub UserForm_Initialize()
    Dim i As Integer

    IS_FIRST_LOAD = True
    Call Initialize_Setting
    
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub



Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


' ***************************************************************
' mod_MakeFieldList
'
' ***************************************************************


Option Explicit

' �̰����ٰ� �⺻���� �������� �������ش�.
' �����̸���, �����ϰ����͵��� ...

Const EXPORT_DATE As String = "2025-02-13"
Const EXPORT_ADDR_HEADER As String = "��⵵ �Ⱦ�� "
Const EXPORT_FILE_NAME As String = "d:\05_Send\iyong_template.xlsx"
        
' 1�� 1�ϴ� �޼���, ���������� ���� �˻�
' ��⵵ �Ⱦ��
Const ONEMAN_WATER_SUPPLY As Double = 289.16
        
Public Enum ALLOW_TYPE_VALUE
     at_HEOGA = 0
     at_SINGO = 1
End Enum


Sub SplitAddressHeader()
    Dim arr() As String
    Dim i As Integer
    
    
    arr = Split(EXPORT_ADDR_HEADER, " ")
        
'    If UBound(arr) >= 2 Then
'        Debug.Print arr(1)
'    Else
'        Debug.Print arr(0)
'    End If
'
    For i = 0 To UBound(arr)
        Debug.Print """" & arr(i) & """"
    Next i
End Sub

' ends with given string
' 2025/3/7
'
Function EndsWith(str As String, endStr As String) As Boolean
    If Right(str, 1) = endStr Then
        EndsWith = True
    Else
        EndsWith = False
    End If
End Function


'
' Make Address Header
'
Function MakeAddressHeader(str As String) As String
    Dim arr() As String

    arr = Split(EXPORT_ADDR_HEADER, " ")
    
    If EndsWith(str, "��") Then
        MakeAddressHeader = arr(0) & " " & str
    Else
        MakeAddressHeader = EXPORT_ADDR_HEADER & str
    End If
    
End Function




Sub delay(ti As Integer)
    Application.Wait Now + TimeSerial(0, 0, ti)
End Sub


Sub MakeFieldList()
    Call make_datamid
    Call Delete_Outside_Boundary
    Call ExportData
End Sub


Sub ExportData()
    ' data_mid ����, �߰��������� ������� ����Ÿ�� �ҷ��ͼ�, ���̽�ó���� ������Ʈ�� �����.
    Call Make_DataOut
    Call ExportCurrentWorksheet("data_out")
End Sub

Sub ExportCurrentWorksheet(sh As String)
    Dim filePath As String
    
    If Not ActivateSheet(sh) Then
        Debug.Print "ActivateSheet Error, maybe sheet does not exist ...."
        Exit Sub
    End If
        
    'filePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    ' filePath = "d:\05_Send\aaa.xlsx"
    
    filePath = EXPORT_FILE_NAME
    
    If VarType(filePath) = vbString Then
    
        If Dir(filePath) <> "" Then
            ' Delete the file
            Kill filePath
    
'            If MsgBox("The file " & filePath & " already exists. Do you want to overwrite it?", _
'                      vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
'                Exit Sub
'            End If
        End If
    
    
        If Sheets(sh).Visible = False Then
            Sheets(sh).Visible = True
        End If
        
        Sheets(sh).Activate
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=xlLocalSessionChanges
        ActiveWorkbook.Close savechanges:=False
    End If
End Sub


Private Sub DeleteFile(filePath As String)
    ' Check if the file exists before attempting to delete
    If Dir(filePath) <> "" Then
        ' Delete the file
        Kill filePath
        ' MsgBox "File deleted successfully.", vbInformation
    Else
        ' MsgBox "File not found.", vbExclamation
    End If
End Sub



Function ActivateSheet(sh As String) As Boolean
    On Error GoTo ErrorHandler
    Sheets(sh).Activate
    ActivateSheet = True
    Exit Function
    
ErrorHandler:
'    MsgBox "An error occurred while trying to activate the sheet." & vbNewLine & _
'           "Please check that the sheet name is correct and try again.", _
'           vbExclamation, "Error"

    ActivateSheet = False
End Function

Sub Make_DataOut()
    Dim str_, address, id, purpose As String
    Dim allowType, i, lastrow  As Integer
    Dim simdo, diameter, hp, capacity, tochool, Q As Double
    Dim setting As String
    
    Dim ag_start, ag_end, ag_year  As String
    Dim sayong_gagu, sayong_ingu, sayong_ilin_geupsoo As Double
    Dim usage_day, usage_month, usage_year As Double
    
    str_ = ChrW(&H2714)
    
    
    If Not Sheets("data_mid").Visible Then
        Sheets("data_mid").Visible = True
    End If
    
    Sheets("data_mid").Activate
    
    Call initialize
    lastrow = getlastrow()
    
    For i = 2 To lastrow
    
        Call GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
        
        If allowType = at_HEOGA Then
            setting = setting & "b,"
            ' �㰡�ü�
        Else
            setting = setting & "c,"
            ' �Ű�ü�
        End If
        
'       ������������, �Ϲݰ��������� �˻��ؼ� �߰����� ...
'       If (diameter >= 150) And (hp >= 1#) Then
'            setting = setting & "aq,"
'       Else
'            setting = setting & "ap,"
'       End If

        setting = setting & IIf(diameter >= 150 And hp >= 1#, "aq,", "ap,")

       
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "f,"
                setting = setting & SS_StringCheck(purpose)
                setting = setting & SS_PublicCheck(purpose)
            
            Case "a"
                setting = setting & "u,"
                setting = setting & AA_StringCheck(purpose)
                
                If allowType = at_HEOGA Then
                    setting = setting & "ab,"
                Else
                    setting = setting & AA_PublicCheck(purpose)
                End If
                                            
            Case "i"
                setting = setting & "o,"
                setting = setting & II_StringCheck(purpose)
                setting = setting & II_PublicCheck(purpose)
                
                
        End Select
        
        ' ����� �ΰ� , �������ִ� ���ΰ� ?
        If IsDrinking(purpose) Then
            setting = setting & "ah,"
        Else
            setting = setting & "ai,"
        End If
        
        
        
        ' ad = ���߻��
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
            
            Case "a"
                '����� : 3 ~ 11������
                ag_start = "3"
                ag_end = "11"
                ag_year = "274"
            
            
            Case "i"
                ' ������ - ���߻��
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
                
        End Select
        
        
        '�����, ��밡��, ����α�, ���α޼����� ������
        If IsDrinking(purpose) Then
                 ' �뵵��, �������϶� ...
                 If CheckSubstring(purpose, "����") Then
                        sayong_gagu = 1
                        sayong_ingu = SS_CITY
                        sayong_ilin_geupsoo = Q / SS_CITY
                 End If
                
                 ' https://kosis.kr/statHtml/statHtml.do?orgId=110&tblId=DT_11001N_2013_A055
                 ' �뵵�� ���̻���� �϶� ...
                 If CheckSubstring(purpose, "����") Then
                        sayong_gagu = 30
                        sayong_ingu = 90
                        sayong_ilin_geupsoo = ONEMAN_WATER_SUPPLY
                End If
        Else
            sayong_gagu = 0
            sayong_ingu = 0
            sayong_ilin_geupsoo = 0
        End If
                
         
        ' �ϻ�뷮 ���
        usage_day = Q
        usage_month = Q * 29
        
        If LCase(Left(id, 1)) = "s" Then
            usage_year = usage_month * 12
        Else
            usage_year = usage_month * 8
        End If
        
        
        '�㰡�� -  av,aw,ay,az,ba,
        
        ' ������Ȳ üũ
        Select Case LCase(Left(id, 1))
            Case "s"
                If allowType = at_SINGO Then ' �Ű�ü��̸�
                    If CheckSubstring(purpose, "�Ϲ�") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "����") Then setting = setting & "av,aw,ax,ay,az,ba,"
                    If CheckSubstring(purpose, "����") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "�ι�") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "�б�") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "û��") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "����") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "���") Then setting = setting & "av,aw,ay,"
                Else ' �㰡�ü��̸�
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            Case "a"
                If allowType = at_SINGO Then ' �Ű�ü��̸�
                    If CheckSubstring(purpose, "����") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "����") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "����") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "���") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "�����") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "���") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "��Ÿ") Then setting = setting & "aw,ay,"
                Else ' �㰡�ü��̸�
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            
            Case "i"
                ' ������ - ���߻��
                setting = setting & "ad,"
                If allowType = at_SINGO Then
                    ' �Ű�ü��̸�
                    setting = setting & "aw,ay,"
                    
                Else
                    ' �㰡�ü��̸�
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
                
        End Select
        
        
        
        
        Debug.Print "**********************************"
        Debug.Print setting
        
        Call PutDataSheetOut(i, setting, address, simdo, diameter, hp, capacity, tochool, Q, ag_start, ag_end, ag_year, _
                             sayong_gagu, sayong_ingu, sayong_ilin_geupsoo, usage_day, usage_month, usage_year)
        
        
        setting = ""
    
    Next i

' =INDEX(itable[value], MATCH("d1", itable[key], 0))

End Sub

Sub PutDataSheetOut(ii As Variant, setting As Variant, address As Variant, simdo As Variant, diameter As Variant, hp As Variant, _
                    capacity As Variant, tochool As Variant, Q As Variant, _
                    ag_start As Variant, ag_end As Variant, ag_year As Variant, _
                    sayong_gagu As Variant, sayong_ingu As Variant, sayong_ilin_geupsoo As Variant, _
                    usage_day As Variant, usage_month As Variant, usage_year As Variant)

    Dim out() As String
    Dim i As Integer
    Dim index, str_, setting_1 As String
    
    Sheets("data_out").Activate
    
    With Range("A" & CStr(ii) & ":BB" & CStr(ii))
        .Value = " "
    End With

    str_ = ChrW(&H2714)
    
    
    setting_1 = DeepCopyString(CStr(setting))
    
    out = FilterString(setting_1)
    
    For i = LBound(out) To UBound(out)
        index = out(i)
        Sheets("data_out").Cells(ii, index).Value = str_
    Next i
    
    '  myString = Format(myDate, "yyyy-mm-dd")
    Sheets("data_out").Cells(ii, "a").Value = " " & Format(EXPORT_DATE, "yyyy-mm-dd") & "."
    Sheets("data_out").Cells(ii, "e").Value = address
    Sheets("data_out").Cells(ii, "ar").Value = simdo
    Sheets("data_out").Cells(ii, "as").Value = diameter
    Sheets("data_out").Cells(ii, "at").Value = hp
    Sheets("data_out").Cells(ii, "au").Value = capacity
    Sheets("data_out").Cells(ii, "av").Value = tochool
    
    Sheets("data_out").Cells(ii, "ae").Value = ag_start
    Sheets("data_out").Cells(ii, "af").Value = ag_end
    Sheets("data_out").Cells(ii, "ag").Value = ag_year
    
    ' ����� �϶���, ��밡��, ����α�, 1�α޼� ����
    If Sheets("data_out").Cells(ii, "ah").Value = ChrW(&H2714) Then
        Sheets("data_out").Cells(ii, "aj").Value = CStr(Format(sayong_gagu, "0.00"))
        Sheets("data_out").Cells(ii, "ak").Value = CStr(Format(sayong_ingu, "0.00"))
        Sheets("data_out").Cells(ii, "al").Value = CStr(Format(sayong_ilin_geupsoo, "0.00"))
    End If
    
    Sheets("data_out").Cells(ii, "am").Value = CStr(Format(usage_day, "0.00"))
    Sheets("data_out").Cells(ii, "an").Value = CStr(Format(usage_month, "#,##0"))
    Sheets("data_out").Cells(ii, "ao").Value = CStr(Format(usage_year, "#,##0"))
    

End Sub
                             
                          
' GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
Sub GetDataFromSheet(i As Variant, id As Variant, address As Variant, allowType As Variant, _
                     simdo As Variant, diameter As Variant, hp As Variant, capacity As Variant, tochool As Variant, _
                     purpose As Variant, Q As Variant)
    
    id = Sheets("data_mid").Cells(i, "a").Value
    address = Sheets("data_mid").Cells(i, "b").Value
    allowType = Sheets("data_mid").Cells(i, "c").Value
    simdo = Sheets("data_mid").Cells(i, "d").Value
    diameter = Sheets("data_mid").Cells(i, "e").Value
    hp = Sheets("data_mid").Cells(i, "f").Value
    capacity = Sheets("data_mid").Cells(i, "g").Value
    tochool = Sheets("data_mid").Cells(i, "h").Value
    purpose = Sheets("data_mid").Cells(i, "i").Value
    Q = Sheets("data_mid").Cells(i, "j").Value
    
End Sub


Function getlastrow() As Integer
    ' ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    getlastrow = ActiveSheet.Range("A3333").End(xlUp).row
End Function


' 2024-1-11 , modify last cell check
' using cell reference SUM_SS, SUM_AA, SUM_II

Sub LastRowFindAll(row_ss As Variant, row_aa As Variant, row_ii As Variant)
    
    If Range("SUM_SS").Value = 0 Then
        row_ss = 0
    Else
        Sheets("ss").Activate
        row_ss = getlastrow() - 1
    End If
           
    If Range("SUM_AA").Value = 0 Then
        row_aa = 0
    Else
        Sheets("aa").Activate
        row_aa = getlastrow() - 1
    End If
      
    
    If Range("SUM_II").Value = 0 Then
        row_ii = 0
        Exit Sub
    Else
        Sheets("ii").Activate
        row_ii = getlastrow() - 1
    End If
    
End Sub

Sub EraseSheetData()
    Worksheets("data_mid").Range("A2:J1000").Delete
    Worksheets("data_out").Range("A2:BD1000").Delete
End Sub




' allowType = 1 - �Ű��
' allowType = 0 - �㰡��
Public Sub make_datamid()
    Dim i, j, row_end As Integer
    Dim newAddress, id, purpose As String
    Dim allowType As Integer
    Dim well_data(1 To 5) As Double
    Dim Q As Double
    Dim boundary As String
    Dim row_ss, row_aa, row_ii As Integer
    
    Call LastRowFindAll(row_ss, row_aa, row_ii)
    Call EraseSheetData
    
    Sheets("ss").Activate
    ' Debug.Print row_end
    For i = 1 To row_ss
        id = Cells(i + 1, "a").Value
        ' �ּ����, ������ ���� ���� �ٽ� �������־�� �Ѵ�.
        ' newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        newAddress = MakeAddressHeader(Cells(i + 1, "c")) & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "�Ű��" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "L").Value
        boundary = Cells(i + 1, "s").Value
        
        
        If Q <> 0 Then
            Call putdata(i, id, newAddress, allowType, well_data, purpose, Q, boundary)
        End If
    Next i
    
    
    Sheets("aa").Activate
    ' Debug.Print row_end
    For i = 1 To row_aa
    
        id = Cells(i + 1, "a").Value
        newAddress = MakeAddressHeader(Cells(i + 1, "c")) & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "�Ű��" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        If Q <> 0 Then
            Call putdata(i + row_ss, id, newAddress, allowType, well_data, purpose, Q, boundary)
        End If
    Next i
    
    Sheets("ii").Activate
    ' Debug.Print row_end
    
    For i = 1 To row_ii
    
        id = Cells(i + 1, "a").Value
        newAddress = MakeAddressHeader(Cells(i + 1, "c")) & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "�Ű��" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        If Q <> 0 Then
            Call putdata(i + row_ss + row_aa, id, newAddress, allowType, well_data, purpose, Q, boundary)
        End If
    Next i
    
End Sub

' 2024-1-11
' delete outside boundary

Private Sub Delete_Outside_Boundary()

    Dim row_ss, row_aa, row_ii As Integer
    Dim i, j As Integer
        
    j = 2
    Sheets("data_mid").Activate
    
    For i = 1 To getlastrow()
        
        If Cells(j, "K").Value = "O" Then
            j = j + 1
        Else
            Rows(j & ":" & j).Select
            Selection.Delete Shift:=xlUp
        End If
    
    Next i

End Sub

Sub putdata(i As Variant, id As Variant, newAddress As Variant, allowType As Variant, well_data As Variant, _
            purpose As Variant, Q As Variant, boundary As Variant)
        
    Sheets("data_mid").Cells(i + 1, "a").Value = id
    Sheets("data_mid").Cells(i + 1, "b").Value = newAddress
    Sheets("data_mid").Cells(i + 1, "c").Value = allowType
    Sheets("data_mid").Cells(i + 1, "d").Value = well_data(1)
    Sheets("data_mid").Cells(i + 1, "e").Value = well_data(2)
    Sheets("data_mid").Cells(i + 1, "f").Value = well_data(3)
    Sheets("data_mid").Cells(i + 1, "g").Value = well_data(4)
    Sheets("data_mid").Cells(i + 1, "h").Value = well_data(5)
    Sheets("data_mid").Cells(i + 1, "i").Value = purpose
    Sheets("data_mid").Cells(i + 1, "j").Value = Q
    Sheets("data_mid").Cells(i + 1, "k").Value = boundary
    
End Sub













' ***************************************************************
' Sheet1_index(index)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' Sheet4_data_out(data_out)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' Sheet6_data_mid(data_mid)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' mod_UniCOde
'
' ***************************************************************


Option Explicit

Sub GetCheckMarkCode()
    Dim checkMark As String
    Dim code As Long

    checkMark = "?" ' the check mark symbol

    code = AscW(checkMark)

    Debug.Print "The Unicode code point for " & checkMark & " is " & code
End Sub

Sub InsertCheckMark()
    Dim checkMark As String

    checkMark = ChrW(&H2714) ' &H2714 is the Unicode code point for the check mark symbol

    Range("A1").Value = checkMark ' Replace "A1" with the cell where you want to insert the check mark symbol
End Sub

Sub TestUniCode()
    Dim i As Integer
    Dim str_check As String
    Dim code As Long
    Dim index As Variant
        
    str_check = Sheets("index").Range("a1").Value
    code = AscW(str_check)
    
    index = Array("a", "b", "c", "f", "k")
    
    For i = LBound(index) To UBound(index)
        Cells(33, index(i)).Value = ChrW(&H2714)   ' str_check
    Next i
    
    Debug.Print "strcheck", code
End Sub



' ***************************************************************
' mod_FilterString
'
' ***************************************************************


Option Explicit

Function FilterString(str As String) As Variant
    Dim elements() As String
    Dim element As Variant
    Dim out() As String
    Dim i As Long
    
    elements = Split(str, ",")
    For Each element In elements
        If element = "" Then Exit For
        ReDim Preserve out(i)
        out(i) = Trim(element)
        i = i + 1
    Next element

    FilterString = out
End Function


'Function DeepCopyString(originalStr As String) As String
'    Dim copiedStr As String
'
'    copiedStr = StrConv(originalStr, vbFromUnicode)
'    DeepCopyString = copiedStr
'
'    ' Debug.Print "Original string: " & originalStr
'    ' Debug.Print "Copied string: " & copiedStr
'End Function


Function DeepCopyString(originalStr As String) As String
    
    DeepCopyString = Left$(originalStr, Len(originalStr))
End Function


Sub TestFilterString()
    Dim out() As Variant
    Dim i As Integer
    
    out = FilterString("a,b, c,d,af, ae, k, x, ag")
    
    Debug.Print "***************************"
    
    For i = LBound(out) To UBound(out)
        
        Debug.Print out(i)
        Debug.Print "***************************"
    
    Next i
End Sub

' ***************************************************************
' mod_CheckString
'
' ***************************************************************


Option Explicit


Function CheckSubstring(str As String, chk As String) As Boolean
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function


Function IsDrinking(str As String) As Boolean
   ' ������ - �缳
    If CheckSubstring(str, "����") Then
            IsDrinking = True
            Exit Function
    End If
    
    ' �Ϲݿ� - �缳
    If CheckSubstring(str, "�Ϲ�") Then
           IsDrinking = False
            Exit Function
    End If
    
    ' �б��� - ����
    If CheckSubstring(str, "�б�") Then
             IsDrinking = True
            Exit Function
    End If
        
    ' �ι����� - ����
    If CheckSubstring(str, "�ι�") Then
             IsDrinking = False
            Exit Function
    End If
    
    ' �������ÿ� - �缳
    If CheckSubstring(str, "����") Then
             IsDrinking = True
            Exit Function
    End If
    
    ' ���̻���� - ����
    If CheckSubstring(str, "����") Then
             IsDrinking = True
            Exit Function
    End If
    
    ' ���Ȱ��� - �缳
    If CheckSubstring(str, "���") Then
             IsDrinking = False
            Exit Function
    End If
    
    ' ��Ÿ - �缳
    If CheckSubstring(str, "��Ÿ") Then
             IsDrinking = False
            Exit Function
    End If
    
    IsDrinking = False
End Function



Function SS_StringCheck(str As String) As String
    ' ������ - �缳
    If CheckSubstring(str, "����") Then
            SS_StringCheck = "g,"
            Exit Function
    End If
    
    ' �Ϲݿ� - �缳
    If CheckSubstring(str, "�Ϲ�") Then
            SS_StringCheck = "h,"
            Exit Function
    End If
    
    ' �б��� - ����
    If CheckSubstring(str, "�б�") Then
            SS_StringCheck = "i,"
            Exit Function
    End If
        
    ' �ι����� - ����
    If CheckSubstring(str, "�ι�") Then
            SS_StringCheck = "j,"
            Exit Function
    End If
    
    ' �������ÿ� - �缳
    If CheckSubstring(str, "����") Then
            SS_StringCheck = "k,"
            Exit Function
    End If
    
    ' ���̻���� - ����
    If CheckSubstring(str, "����") Then
            SS_StringCheck = "l,"
            Exit Function
    End If
    
    ' ���Ȱ��� - �缳
    If CheckSubstring(str, "���") Then
            SS_StringCheck = "m,"
            Exit Function
    End If
    
    ' ��Ÿ - �缳
    If CheckSubstring(str, "��Ÿ") Then
            SS_StringCheck = "n,"
            Exit Function
    End If
    
    SS_StringCheck = "n,"
End Function

Function AA_StringCheck(str As String) As String
    
    ' ������� ���� �缳, ���� �㰡�� - ����
    If CheckSubstring(str, "����") Then
            AA_StringCheck = "v,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            AA_StringCheck = "w,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            AA_StringCheck = "x,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            AA_StringCheck = "y,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            AA_StringCheck = "z,"
            Exit Function
    End If
    
    If CheckSubstring(str, "��Ÿ") Then
            AA_StringCheck = "aa,"
            Exit Function
    End If
    
    AA_StringCheck = "aa,"
End Function


Function II_StringCheck(str As String) As String
    ' �ذ�, ����, ��� - ����
    If CheckSubstring(str, "����") Then
            II_StringCheck = "p,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            II_StringCheck = "q,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            II_StringCheck = "r,"
            Exit Function
    End If
    
    ' ��������, ��Ÿ - �缳
    If CheckSubstring(str, "��������") Then
            II_StringCheck = "s,"
            Exit Function
    End If
    
    If CheckSubstring(str, "��Ÿ") Then
            II_StringCheck = "t,"
            Exit Function
    End If

    II_StringCheck = "t,"
End Function



Function SS_PublicCheck(str As String) As String
    ' ������ - �缳
    If CheckSubstring(str, "����") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' �Ϲݿ� - �缳
    If CheckSubstring(str, "�Ϲ�") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' �б��� - ����
    If CheckSubstring(str, "�б�") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
        
    ' �ι����� - ����
    If CheckSubstring(str, "�ι�") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' �������ÿ� - �缳
    If CheckSubstring(str, "����") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' ���̻���� - ����
    If CheckSubstring(str, "����") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' ���Ȱ��� - �缳
    If CheckSubstring(str, "���") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' ��Ÿ - �缳
    If CheckSubstring(str, "��Ÿ") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    SS_PublicCheck = "ac,"
End Function

Function AA_PublicCheck(str As String) As String
    ' ������� ���� �缳, ���� �㰡�� - ����
    If CheckSubstring(str, "����") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "��Ÿ") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    AA_PublicCheck = "ac,"
End Function


Function II_PublicCheck(str As String) As String
    ' �ذ�, ����, ��� - ����
    If CheckSubstring(str, "����") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    If CheckSubstring(str, "����") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    If CheckSubstring(str, "���") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' ��������, ��Ÿ - �缳
    If CheckSubstring(str, "��������") Then
            II_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "��Ÿ") Then
            II_PublicCheck = "ac,"
            Exit Function
    End If

    II_PublicCheck = "ac,"
End Function


' ***************************************************************
' Sheet1(ref)
'
' ***************************************************************


Option Explicit

Private Sub CommandButton1_Click()
    ActiveSheet.Visible = False
End Sub
' ***************************************************************
' Sheet2(ref1)
'
' ***************************************************************




Option Explicit

Private Sub CommandButton1_Click()
    ActiveSheet.Visible = False
End Sub
' ***************************************************************
' modTable
'
' ***************************************************************


Option Explicit

Sub test_tableindex()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("tableCNU")
    
    Dim values() As Variant
    values = tbl.ListColumns("nonsan").DataBodyRange.Value
    
    Dim i As Long
    For i = 1 To UBound(values, 1)
        ActiveSheet.Cells(29, Chr(Asc("A") + i)).Value = values(i, 1)
    Next i
End Sub


Sub initialize_CNU(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref").ListObjects("tableCNU")
    
    Dim values() As Variant
    
    If area = "" Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
         
    '���󳲵�, ������, 2020 ȯ��� ���ϼ�����������ħ
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
    
End Sub


Sub initialize_JIYEOL(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    Dim values() As Variant
    
    
    If (area = "") Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
           
    '���󳲵�, ������, 2020 ȯ��� ���ϼ�����������ħ
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
End Sub




' ***************************************************************
' UserForm_SS
'
' ***************************************************************

' Optionbutton1 - ������
' Optionbutton2 - �Ϲݿ�
' Optionbutton3 - û�ҿ�
' Optionbutton4 - �ι�����
' Optionbutton5 - �б���
' Optionbutton6 - �������ÿ�
' Optionbutton7 - ���̻����
' Optionbutton8 - ���Ȱ���
' Optionbutton9 - ��Ÿ
' Optionbutton10 - �����
' Optionbutton11 - �����ó���
' Optionbutton12 - �����
' Optionbutton13 - �ҹ��

Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("������", "�Ϲݿ�", "û�ҿ�", "�ι�����", "�б���", "�������ÿ�", "���̻����", "���Ȱ���", "��Ÿ", "�����", "�����ó���", "�����", "�ҹ��")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 12
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 Then
'        Unload Me
'    End If
'End Sub
' ***************************************************************
' UserForm_AA
'
' ***************************************************************


' Optionbutton1 - ���ۿ�
' Optionbutton2 - ���ۿ�
' Optionbutton3 - ������
' Optionbutton4 - ����
' Optionbutton5 - ������
' Optionbutton6 - ��Ÿ


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("���ۿ�", "���ۿ�", "������", "����", "������", "��Ÿ")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 5
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub


Private Sub CommandButton2_Click()
  Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
    
End Sub




Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub



Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim fName As String
    
    Dim pathToExport As String
    pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        
        
        If item.CodeModule.CountOfLines <> 0 Then
            lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        Else
            lineToPrint = "'This Module is Empty "
        End If
        
        
        fName = item.CodeModule.Name
        Debug.Print lineToPrint
        SaveTextToFile lineToPrint, pathToExport & fName & ".bas"
        
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub


' ***************************************************************
' UserForm_II
'
' ***************************************************************

' Optionbutton1 - ����������ü
' Optionbutton2 - ��Ÿ
' Optionbutton3 - �������
' Optionbutton4 - �������
' Optionbutton5 - �����������
' Optionbutton6 - ����������


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("����������ü", "��Ÿ", "�������", "�������", "�����������", "����������")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 5
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 Then
'        Unload Me
'    End If
'End Sub



Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:02"), "Popup_CloseUserForm"
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
   
    Me.TextBox1.Text = "this is Sample initialize"
End Sub

Sub Popup_MessageBox(ByVal msg As String)
    UserForm1.TextBox1.Text = msg
    UserForm1.Show
End Sub

Sub Popup_CloseUserForm()
    Unload UserForm1
End Sub

Sub test()
    ' Application.OnTime Now + TimeValue("00:00:01"), "Popup_CloseUserForm"
    Popup_MessageBox ("Automatic Close at One Seconds ...")
End Sub


