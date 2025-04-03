Attribute VB_Name = "mod_DuplicatetWellSpec"


' �⺻���������͸� , �������� ���� GetOtherFileName
Function GetOtherFileName(Optional ByVal SearchText As String = "����Ÿ") As String
    Dim Workbook As Workbook
    Dim WBNAME As String
    Dim i As Long

    If Workbooks.count <> 2 Then
        GetOtherFileName = "NOTHING"
        Exit Function
    End If

    For Each Workbook In Application.Workbooks
        WBNAME = Workbook.name
        If StrComp(ThisWorkbook.name, WBNAME, vbTextCompare) = 0 Then
        ' �̸��� thisworkbook.name �� ���ٸ� , �����б��
            GoTo NEXT_ITERATION
        End If
        
        If ThisWorkbook.name <> Workbook.name And CheckSubstring(WBNAME, SearchText) Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    
    If ThisWorkbook.name <> WBNAME And CheckSubstring(WBNAME, SearchText) Then
        GetOtherFileName = WBNAME
    Else
        GetOtherFileName = "NOTHING"
    End If
End Function


Sub CheckSheetExists(WB_NAME As String)
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "All" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' Do something if sheet exists
    If sheetExists Then
        MsgBox "Sheet 'All' exists!"
        ' Place your code here to do something
    Else
        MsgBox "Sheet 'All' does not exist."
    End If
End Sub


'**********************************************************************************************************************

Sub InteriorCopyDirection(this_WBNAME As String, well_no As Integer, IS_OVER180 As Boolean)

    Workbooks(this_WBNAME).Worksheets(CStr(well_no)).Activate
    
    If IS_OVER180 Then
        Range("K12").Font.Bold = True
        Range("L12").Font.Bold = False
        
        CellBlack (ActiveSheet.Range("K12"))
        CellLight (ActiveSheet.Range("L12"))
    Else
        Range("K12").Font.Bold = False
        Range("L12").Font.Bold = True
        
        CellBlack (ActiveSheet.Range("L12"))
        CellLight (ActiveSheet.Range("K12"))
    End If
End Sub


Private Sub CellBlack(S As Range)
    S.Select
    
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .themeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub

Private Sub CellLight(S As Range)
    S.Select
    
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub


Private Function get_direction() As Long
    ' get direction is cell is bold
    ' ���� ���尪�̸� ������ �Ѵ�.  ������ �ΰ��߿��� �ϳ��� �����ϰ� �ȴ�.
    ' 2019/10/18��
    
    Range("k12").Select
    
    If Selection.Font.Bold Then
        get_direction = Range("k12").value
    Else
        get_direction = Range("L12").value
    End If
End Function


'**********************************************************************************************************************


Sub DuplicateWellSpec(ByVal this_WBNAME As String, ByVal WB_NAME As String, ByVal well_no As Integer, obj As Class_Boolean)
    ' Dim WB_NAME As String
    Dim i As Integer
    Dim long_axis, short_axis, well_distance, well_height, surface_water_height As Long
    Dim degree_of_flow As Double
    Dim IS_OVER180 As Boolean


'    obj.Result = False, ��������
'    obj.Result = True , ��������
      
    If Workbooks.count <> 2 Then
        MsgBox "Please Open, �⺻��������Ÿ�� ����,  �⺻��������Ÿ ���� �ϳ��� �ҷ��ü��� �ֽ��ϴ�. ", vbOKOnly
        obj.result = True
        Exit Sub
    End If
   
    
    ' WB_NAME = GetOtherFileName
    IS_OVER180 = False
    
    If WB_NAME = "NOTHING" Then
        GoTo SheetDoesNotExist
    End If
    
    On Error GoTo SheetDoesNotExist
    
    With Workbooks(WB_NAME).Worksheets(CStr(well_no))
        long_axis = .Range("K6").value
        short_axis = .Range("K7").value
        degree_of_flow = .Range("K12").value
        
        If .Range("K12").Font.Bold Then
            IS_OVER180 = True
        End If
        
        well_distance = .Range("K13").value
        well_height = .Range("K14").value
        surface_water_height = .Range("K15").value
    End With
    

    With Workbooks(this_WBNAME).Worksheets(CStr(well_no))
        .Range("K6") = long_axis
        .Range("K7") = short_axis
        .Range("K12") = degree_of_flow
        .Range("K13") = well_distance
        .Range("K14") = well_height
        .Range("K15") = surface_water_height
    End With
    
    Call InteriorCopyDirection(this_WBNAME, well_no, IS_OVER180)

    obj.result = False
    Exit Sub

SheetDoesNotExist:
    MsgBox "Please Open, �⺻��������Ÿ ������ �ƴմϴ�. ", vbOKOnly
    obj.result = True
    
End Sub

Sub Duplicate_WATER(ByVal this_WBNAME As String, ByVal WB_NAME As String)

    Dim cpRange As String
    
    cpRange = "E7:L8"
    
'    Workbooks(WB_NAME).Sheets("water").Visible = True
'    Workbooks(this_WBNAME).Sheets("water").Visible = True
    
    Workbooks(WB_NAME).Worksheets("water").Activate
    Workbooks(WB_NAME).Worksheets("water").Range(cpRange).Select
    Selection.Copy
    
    
    Workbooks(this_WBNAME).Worksheets("water").Activate
    Workbooks(this_WBNAME).Worksheets("water").Range("E7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    Workbooks(WB_NAME).Sheets("water").Visible = False
'    Workbooks(this_WBNAME).Sheets("water").Visible = False

End Sub





Function aCellContains(searchRange As Range, searchValue As String) As Boolean
    aCellContains = InStr(1, LCase(searchRange.value), LCase(searchValue)) > 0
End Function


Function aFindCellByLoopingPartialMatch(wb As Workbook) As String

    Dim ws As Worksheet
    Dim cell As Range
    Dim Address As String
     
     For Each cell In wb.Worksheets("Well").Range("A1:AZ1").Cells
        Debug.Print cell.Address, cell.value
    
        If aCellContains(cell, "") Then
            Address = cell.Address
            Exit For
        End If
    Next
    
    aFindCellByLoopingPartialMatch = Address
    
End Function



Sub Duplicate_WELL_MAIN(ByVal this_WBNAME As String, ByVal WB_NAME As String, ByVal nofwell As Integer)

   Dim cpRange, Title As String
    
    cpRange = "A4:P" & (nofwell + 4 - 1)
    
    Workbooks(WB_NAME).Worksheets("Well").Activate
    Workbooks(WB_NAME).Worksheets("Well").Range(cpRange).Select
    Selection.Copy
    
    Workbooks(this_WBNAME).Worksheets("Well").Activate
    Workbooks(this_WBNAME).Worksheets("Well").Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    
    ' 2024/6/26��, Copy Title
    ' 2024/12/26 Search Title location
    
    titleCell = aFindCellByLoopingPartialMatch(Workbooks(WB_NAME))
    Title = Workbooks(WB_NAME).Worksheets("Well").Range(titleCell).value
    EraseCellData ("A1:G1")
    Workbooks(this_WBNAME).Worksheets("Well").Range("D1") = Title
    
    ' End of Copy Title
    
    
    Application.CutCopyMode = False
    Range("A4").Select
End Sub



Sub ImportWellSpec_OLD(ByVal well_no As Integer, obj As Class_Boolean)
    Dim WkbkName As Object
    Dim WBNAME As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, DeltaS As Double
    Dim Casing As Integer

    WBNAME = "A" & GetNumeric2(well_no) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBNAME) Then
        MsgBox "Please open the yangsoo data ! " & WBNAME
        obj.result = True
        Exit Sub
    Else
        obj.result = False
    End If

    ' delta s : ����1���� ��������
    DeltaS = Workbooks(WBNAME).Worksheets("SkinFactor").Range("b4").value
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
    nl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i4").value
    sl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i6").value
    Casing = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i10").value
    
    ' WkbkName.Close
    T1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("D5").value
    S1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("E10").value
    T2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("H13").value
    S2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i16").value
    S3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i13").value
    
    ' Skin Coefficient
    Skin = Workbooks(WBNAME).Worksheets("SkinFactor").Range("G6").value
    
    ' yangsoo radius of influence
    RI1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C13").value
    RI2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C18").value
    RI3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C23").value
    
    ' ��ȿ�칰�ݰ� , �������� ����
    ir = GetEffectiveRadius(WBNAME)
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
    Range("c20") = nl
    Range("c20").numberFormat = "0.00"
    
    Range("c21") = sl
    Range("c21").numberFormat = "0.00"
    
    Range("c10") = 5
    Range("c11") = Casing - 5
    
    'in recover test, s' value
    Range("G6") = S3
        
    Range("E5") = T1
    Range("E5").numberFormat = "0.0000"
     
    Range("E6") = T2
    Range("E6").numberFormat = "0.0000"
    
    Range("g5") = S2
    Range("g5").numberFormat = "0.0000000"
    
    Range("G4") = S1
    
    Range("h5") = Skin 'skin coefficient
    Range("h6") = ir 'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(DeltaS, 2) 'deltas
End Sub
