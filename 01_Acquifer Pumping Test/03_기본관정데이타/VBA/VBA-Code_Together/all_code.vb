
Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").SingleColor.value = True
    Sheets("Recharge").cbCheSoo.value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub




Private Sub CommandButton_boryung_Click()
    Call importRainfall_button("BORYUNG")
End Sub

Private Sub CommandButton_buyeo_Click()
    Call importRainfall_button("BUYEO")
End Sub

Private Sub CommandButton_cheonan_Click()
    Call importRainfall_button("CHEONAN")
End Sub

Private Sub CommandButton_cheongju_Click()
    Call importRainfall_button("CHEONGJU")
End Sub

Private Sub CommandButton_daejeon_Click()
    Call importRainfall_button("DAEJEON")
End Sub

Private Sub CommandButton_geumsan_Click()
   Call importRainfall_button("GEUMSAN")
End Sub

Private Sub CommandButton_incheon_Click()
    Call importRainfall_button("incheon")
End Sub

Private Sub CommandButton_seosan_Click()
    Call importRainfall_button("SEOSAN")
End Sub


Private Sub CommandButton_Seoul_Click()
     Call importRainfall_button("SEOUL")
End Sub

Private Sub CommandButton_suwon_Click()
    Call importRainfall_button("suwon")
End Sub

Private Sub CommandButton1_Click()
    Call importRainfall
End Sub

Private Sub CommandButton2_Click()
    Range("b5:n34").ClearContents
End Sub



' this is empty

Private Sub CommandButton1_Click()
    Call getMotorPower
    Call FindMaxMin
End Sub

Private Sub CommandButton2_Click()
    Rows("36:90").Select
    Selection.Delete Shift:=xlUp
    Range("g34").Select
End Sub

Private Sub FindMaxMin()

    Dim nWell As Integer
    Dim maxVal, minVal As Double
    Dim qMax, qMin As Double
    
    
    nWell = GetNumberOfWell()
    
    If nWell <= 1 Then
        Exit Sub
    End If
    
    maxVal = Application.WorksheetFunction.max(Range("B46:" & ColumnNumberToLetter(nWell + 1) & "46"))
    minVal = Application.WorksheetFunction.min(Range("B46:" & ColumnNumberToLetter(nWell + 1) & "46"))
    
    qMax = Application.WorksheetFunction.max(Range("B41:" & ColumnNumberToLetter(nWell + 1) & "41"))
    qMin = Application.WorksheetFunction.min(Range("B41:" & ColumnNumberToLetter(nWell + 1) & "41"))
    
    
    Range("k52") = minVal
    Range("k53") = maxVal
    
    Range("l52") = qMin
    Range("l53") = qMax

End Sub

Private Sub ShowLocation_Click()
      Sheets("location").Visible = True
      Sheets("location").Activate
End Sub


Sub hfill_red(ByVal i As Integer)
    Range("C" & i & ":P" & i).Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Sub hfill_clear()
    Range("C5:P17").Select
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Private Sub CommandButton3_Click()
    Dim i As Integer
    Dim max, min As Single
    
    max = Range("o15").value
    min = Range("o16").value
    
    Range("B5:P14").Select
    Selection.Font.Bold = False
     
    Range("a1").Activate
    Call hfill_clear
    
    For i = 5 To 14
        If Cells(i, "O").value = max Or Cells(i, "O").value = min Then
            Call hfill_red(i)
            Union(Cells(i, "B"), Cells(i, "O")).Select
            Selection.Font.Bold = True
        End If
    Next i
End Sub


Option Explicit

Private Sub CommandButton_PressAll_Click()
    Call PressAll_Button
End Sub

Private Sub CommandButton1_Click()
' add well

    BaseData_ETC_02.TurnOffStuff
    Call AddWell_CopyOneSheet
    BaseData_ETC_02.TurnOnStuff
    
End Sub


'AggChart Button
Private Sub CommandButton10_Click()
    Sheets("AggChart").Visible = True
    Sheets("AggChart").Select
End Sub


'AggFx Button
Private Sub CommandButton11_Click()
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
End Sub

Private Sub CommandButton12_Click()
    Sheets("water").Visible = True
    Sheets("water").Select
End Sub

Private Sub CommandButton13_Click()
' �̰���, FX���� ����Ϻ� �����͸� ������ ����,
' ������ ����, 1, 2, 3 �� ������
' FX ���� ������ �´�.
' Import All Well Spec

    Call ImportAll_EachWellSpec
End Sub


Private Sub CommandButton14_Click()
    'wSet, WellSpec Setting

    Dim nofwell, i As Integer

    nofwell = sheets_count()
    
    For i = 1 To nofwell
        Cells(i + 3, "E").formula = "=Recharge!$I$24"
        Cells(i + 3, "F").formula = "=All!$B$2"
        Cells(i + 3, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i
End Sub

Private Sub CommandButton15_Click()
' 2024/6/24 - dupl, duplicate basic well data ...
' �⺻��������Ÿ �����ϴ°�
' ������ ��ȸ�ϸ鼭, �ű⿡�� �����͸� ������ ���µ� ��
' ���� , �����, �����
' ����, �Ÿ�, ��������, ��ǥ��ǥ�� �̷��� ������ ���� �ɵ��ϴ�.

' k6 - ����� / long axis
' k7 - ����� / short axis
' k12 - degree of flow
' k13 - well distance
' k14 - well height
' k15 - surfacewater height

    Call DuplicateBasicWellData
  
End Sub

'AggSum Button
Private Sub CommandButton3_Click()
    Sheets("AggSum").Visible = True
    Sheets("AggSum").Select
End Sub



'Aggregate1 Button
Private Sub CommandButton4_Click()
    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
End Sub




Private Sub CommandButton5_Click()
' Aggregate2 Button
' �����Լ� 2��

    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
End Sub


'AggWhpa Button
Private Sub CommandButton7_Click()
    Sheets("aggWhpa").Visible = True
    Sheets("aggWhpa").Select
End Sub


Private Sub CommandButton9_Click()
  Sheets("AggStep").Visible = True
  Sheets("AggStep").Select
End Sub


'Jojung Button
'add new feature - correct border frame ...
Private Sub CommandButton2_Click()
    Call JojungButton
End Sub



' delete last
Private Sub CommandButton8_Click()
    Call DeleteLast
End Sub

Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub

'one button / delete all well except for one ...

Private Sub CommandButton6_Click()
   Call Make_OneButton
End Sub






Private Sub CommandButton4_Click()
    Call delete_allWhpaData
End Sub



Private Sub CommandButton2_Click()
    Call TurnOffStuff
    Call main_drasticindex
    Call print_drastic_string
    Call TurnOnStuff
End Sub

Private Sub CommandButton3_Click()
    Call getWhpaData_AllWell
End Sub

Private Sub CommandButton7_Click()
   Call getWhpaData_EachWell
End Sub



Private Sub CommandButton5_Click()
    Call BaseData_DrasticIndex.ToggleDirection
End Sub


Private Function get_rf_number() As String
    Dim rf_num As String

    '=(max*rf_1*E17/1000)
    get_rf_number = VBA.Mid(Range("F17").formula, 10, 1)

End Function


Private Sub Set_RechargeFactor_One()

    Range("F17").formula = "=(max*rf_1*E17/1000)"
    Range("F19").formula = "=(max*rf_1*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio"
    Range("G19").formula = "=F19*allow_ratio"
    
    Range("E13").formula = "=Recharge!I24"
    Range("F13").formula = "=rf_1"
    Range("G13").formula = "=allow_ratio"
    
    Range("E26").formula = "=Recharge!C30"
    
End Sub

Private Sub Set_RechargeFactor_Two()

    Range("F17").formula = "=(max*rf_2*E17/1000)"
    Range("F19").formula = "=(max*rf_2*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio2"
    Range("G19").formula = "=F19*allow_ratio2"
    
    
    Range("E13").formula = "=Recharge!I25"
    Range("F13").formula = "=rf_2"
    Range("G13").formula = "=allow_ratio2"
    
    
    Range("E26").formula = "=Recharge!D30"
End Sub


Private Sub Set_RechargeFactor_Three()

    Range("F17").formula = "=(max*rf_3*E17/1000)"
    Range("F19").formula = "=(max*rf_3*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio3"
    Range("G19").formula = "=F19*allow_ratio3"
    
    Range("E13").formula = "=Recharge!I26"
    Range("F13").formula = "=rf_3"
    Range("G13").formula = "=allow_ratio3"
    
    Range("E26").formula = "=Recharge!E30"
    
End Sub



Private Sub CommandButton6_Click()
'Select Recharge Factor

    
   If Frame1.Controls("optionbutton1").value = True Then
        Call Set_RechargeFactor_One
   End If
    
   If Frame1.Controls("optionbutton2").value = True Then
        Call Set_RechargeFactor_Two
   End If
    
   If Frame1.Controls("optionbutton3").value = True Then
        Call Set_RechargeFactor_Three
   End If
    

End Sub



' 2022/6/9 Import YangSoo Data
' Radius of Influence - �������ݰ�
' Effective Radius - ��ȿ�칰�ݰ�
' 2024/6/7 - ��Ų��� �߰����� ...
' 2024/7/9 - ������ ����Ʈ �ؿ��°���, FX ���� �����´�.

Private Sub CommandButton8_Click()
   
   Call modWell_Each.ImportEachWell(Range("E15").value)
        
End Sub

Private Sub Worksheet_Activate()

    Select Case get_rf_number
    
        Case "1"
             Frame1.Controls("optionbutton1").value = True
             
        Case "2"
             Frame1.Controls("optionbutton2").value = True
             
        Case "3"
             Frame1.Controls("optionbutton3").value = True
             
        Case Else
            Frame1.Controls("optionbutton1").value = True
           
    End Select

End Sub


'�⺻��������Ÿ - ���ƽ�ε���
Option Explicit

Dim Dr, Rr  As Single

Public Enum DRASTIC_MODE
    dmGENERAL = 0
    dmCHEMICAL = 1
End Enum

Sub ShiftNewYear()
    Range("B6:N34").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.Copy
    
    Range("B5").Select
    ActiveSheet.PasteSpecial format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
    
    Range("B34:N34").Select
    Selection.ClearContents
    
    ActiveWindow.SmallScroll Down:=18
    Range("B42:N50").Select
    Selection.Copy
    Range("B41").Select
    ActiveSheet.PasteSpecial format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
    Range("B50:N50").Select
    Selection.ClearContents
End Sub

Sub ToggleDirection()
    If Range("k12").Font.Bold Then
        Range("K12").Font.Bold = False
        Range("L12").Font.Bold = True
        
        CellBlack (ActiveSheet.Range("L12"))
        CellLight (ActiveSheet.Range("K12"))
    Else
        Range("K12").Font.Bold = True
        Range("L12").Font.Bold = False
        
        CellBlack (ActiveSheet.Range("K12"))
        CellLight (ActiveSheet.Range("L12"))
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

'Drastic Index �� ��� ���ֱ� ���� �Լ� ...
' 2017/11/21 ȭ����

' 1, ���ϼ����� ���� ����� ���
Private Function Rating_UnderGroundWater(ByVal water_level As Single) As Integer
    Select Case water_level
        Case Is < 1.52
            Rating_UnderGroundWater = 10
        Case Is < 4.57
            Rating_UnderGroundWater = 9
        Case Is < 9.14
            Rating_UnderGroundWater = 7
        Case Is < 15.24
            Rating_UnderGroundWater = 5
        Case Is < 22.86
            Rating_UnderGroundWater = 3
        Case Is < 30.48
            Rating_UnderGroundWater = 2
        Case Else
            Rating_UnderGroundWater = 1
    End Select
End Function


'2, ������ �����Ծ緮
Private Function Rating_NetRecharge(ByVal value As Single) As Integer
    Select Case value
        Case Is < 5.08
            Rating_NetRecharge = 1
        Case Is < 10.16
            Rating_NetRecharge = 3
        Case Is < 17.78
            Rating_NetRecharge = 6
        Case Is < 25.4
            Rating_NetRecharge = 8
        Case Else
            Rating_NetRecharge = 9
    End Select
End Function



'3, �����
Private Function Rating_AqMedia(ByVal value As String) As Integer
    Dim ratings As New Dictionary
    
    ratings.Add "Massive Shale", 2
    ratings.Add "Metamorphic/Igneous", 3
    ratings.Add "Weathered Metamorphic / Igneous", 4
    ratings.Add "Glacial Till", 5
    ratings.Add "Bedded SandStone", 6
    ratings.Add "Massive Sandstone", 6
    ratings.Add "Massive Limestone", 6
    ratings.Add "Sand And Gravel", 8
    ratings.Add "Basalt", 9
    ratings.Add "Karst Limestone", 10

    If ratings.Exists(value) Then
        Rating_AqMedia = ratings(value)
    Else
        Rating_AqMedia = 0
    End If
End Function


'4 ���Ư���� ���� ���

Private Function Rating_SoilMedia(ByVal value As String) As Integer
    Select Case value
        Case "Thin Or Absent", "Gravel"
            Rating_SoilMedia = 10
        Case "Sand"
            Rating_SoilMedia = 9
        Case "Peat"
            Rating_SoilMedia = 8
        Case "Shrinking Or Aggregated Clay"
            Rating_SoilMedia = 7
        Case "Sandy Loam"
            Rating_SoilMedia = 6
        Case "Loam"
            Rating_SoilMedia = 5
        Case "Silty Loam"
            Rating_SoilMedia = 4
        Case "Clay Loam"
            Rating_SoilMedia = 3
        Case "Mud"
            Rating_SoilMedia = 2
        Case "Nonshrinking And Nonaggregated Clay"
            Rating_SoilMedia = 1
    End Select
End Function


' 5, ��������
Private Function Rating_Topo(ByVal value As Single) As Integer
    Select Case value
        Case Is < 2
            Rating_Topo = 10
        Case Is < 6
            Rating_Topo = 9
        Case Is < 12
            Rating_Topo = 5
        Case Is < 18
            Rating_Topo = 3
        Case Else
            Rating_Topo = 1
    End Select
End Function



'6 ����ȭ���� ���⿡ ���� ��� Ir
'
Private Function Rating_Vadose(ByVal value As String) As Integer
    Select Case value
        Case "Confining Layer"
            Rating_Vadose = 1
        Case "Silt/Clay", "Shale"
            Rating_Vadose = 3
        Case "Limestone", "Sandstone", "Bedded Limestone, Sandstone, Shale", "Sand And Gravel With Significant Silt And Clay"
            Rating_Vadose = 6
        Case "Metamorphic/Igneous"
            Rating_Vadose = 4
        Case "Sand And Gravel"
            Rating_Vadose = 8
        Case "Basalt"
            Rating_Vadose = 9
        Case "Karst Limestone"
            Rating_Vadose = 10
    End Select
End Function


' 7, ������� ������������ ���� ��� : Cr

Private Function Rating_EC(ByVal value As Double) As Integer
    Select Case value
        Case Is < 0.0000472
            Rating_EC = 1
        Case Is < 0.000142
            Rating_EC = 2
        Case Is < 0.00033
            Rating_EC = 4
        Case Is < 0.000472
            Rating_EC = 6
        Case Is < 0.000944
            Rating_EC = 8
        Case Else
            Rating_EC = 10
    End Select
End Function



Public Sub find_average()
    ' 2019/10/18 �� �ۼ���
    ' get ���������, �����, ����, ��������� ����� ���ؾ��Ѵ�.
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    Dim nTooSoo     As Single: nTooSoo = 0
    Dim nDaeSoo     As Single: nDaeSoo = 0
    Dim nDirection  As Single: nDirection = 0
    Dim nGradient   As Single: nGradient = 0
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        
        Worksheets(CStr(i)).Activate
        
        nTooSoo = nTooSoo + Range("E7").value
        nDaeSoo = nDaeSoo + Range("C14").value
        nDirection = nDirection + get_direction()
        nGradient = nGradient + Range("K18").value
        
    Next i
    
    Worksheets("1").Activate
    
    Range("J3").value = nTooSoo / n_sheets
    Range("J4").value = nDaeSoo / n_sheets
    Range("J5").value = nDirection / n_sheets
    Range("J6").value = nGradient / n_sheets
    
    Range("k3").formula = "=round(j3,4)"
    Range("k4").formula = "=round(j4,1)"
    Range("k5").formula = "=round(j5,1)"
    Range("k6").formula = "=round(j6,4)"
    
    Call make_frame
End Sub

Public Sub find_average2(ByVal sheet As Integer, ByVal nof_well As Integer)
    ' 2019/10/18 �� �ۼ���
    ' get ���������, �����, ����, ��������� ����� ���ؾ��Ѵ�.
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    Dim nTooSoo     As Single: nTooSoo = 0
    Dim nDaeSoo     As Single: nDaeSoo = 0
    Dim nDirection  As Single: nDirection = 0
    Dim nGradient   As Single: nGradient = 0
    
    Worksheets(CStr(sheet)).Activate
    
    For i = 1 To nof_well
        Worksheets(CStr(i + sheet - 1)).Activate
        
        nTooSoo = nTooSoo + Range("E7").value
        nDaeSoo = nDaeSoo + Range("C14").value
        nDirection = nDirection + get_direction()
        nGradient = nGradient + Range("K18").value
    Next i
    
    Worksheets(CStr(sheet)).Activate
    
    Range("J3").value = nTooSoo / nof_well
    Range("J4").value = nDaeSoo / nof_well
    Range("J5").value = nDirection / nof_well
    Range("J6").value = nGradient / nof_well
    
    Range("k3").formula = "=round(j3,4)"
    Range("k4").formula = "=round(j4,1)"
    Range("k5").formula = "=round(j5,1)"
    Range("k6").formula = "=round(j6,4)"
    
    Call make_frame2(sheet)
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

Sub main_drasticindex()
    Dim water_level, net_recharge, topo, EC As Single
    Dim AQ, Soil, Vadose As String
    Dim drastic_string As String
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    ' ��Ʈ�� ���� ..., �˻��� ���� ����
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        
        '1
        water_level = Range("D26").value
        Range("D27").value = Rating_UnderGroundWater(water_level)
        
        '2
        net_recharge = Range("E26").value
        Range("E27").value = Rating_NetRecharge(net_recharge)
        
        '3
        AQ = Range("F26").value
        Range("F27").value = Rating_AqMedia(AQ)
        
        '4
        Soil = Range("G26").value
        Range("G27").value = Rating_SoilMedia(Soil)
        
        '5
        topo = Range("H26").value
        Range("H27").value = Rating_Topo(topo)
        
        '6 Iv, Vadose
        Vadose = Range("I26").value
        Range("I27").value = Rating_Vadose(Vadose)
        
        '7
        EC = Range("J26").value
        Range("J27").value = Rating_EC(EC)
        
    Next i
End Sub


Function check_DrasticIndex(ByVal dmMode As Integer) As String
    ' dmGENERAL = 0
    ' dmCHEMICAL = 1
    
    Dim value As Integer
    Dim result As String
    
    If dmMode = dmGENERAL Then
        value = Range("K30").value
    Else
        value = Range("K31").value
    End If
    
    Select Case value
        Case Is <= 100
            result = "�ſ쳷��"
        Case Is <= 120
            result = "����"
        Case Is <= 140
            result = "��������"
        Case Is <= 160
            result = "�߰�����"
        Case Is <= 180
            result = "����"
        Case Else
            result = "�ſ����"
    End Select
    
    check_DrasticIndex = result
End Function



Public Sub print_drastic_string()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Range("k26").value = check_DrasticIndex(dmGENERAL)
        Range("k27").value = check_DrasticIndex(dmCHEMICAL)
    Next i
End Sub

' this is empty


Private Sub CommandButton1_Click()
' QT - Quality Test
' Import Quality Test From YangSoo
  Call ImportAll_QT
End Sub


'Get Water Spec from YanSoo ilbo
Private Sub CommandButton2_Click()

    Call GetWaterSpecFromYangSoo_Q1

End Sub


' Get(Ec, Ph, Temp) Range - ���������� ��賻�� �Լ� ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub CommandButton3_Click()
    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC_Q1(cellLOW, i)
        hiEC(i) = getEC_Q1(cellHI, i)
        
        lowPH(i) = getPH_Q1(cellLOW, i)
        hiPH(i) = getPH_Q1(cellHI, i)
        
        lowTEMP(i) = getTEMP_Q1(cellLOW, i)
        hiTEMP(i) = getTEMP_Q1(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"
End Sub



' make summary page
Private Sub CommandButton4_Click()
    Dim nofwell As Integer
    Dim i As Integer
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    
    nofwell = GetNumberOfWell()
    
    For i = 1 To nofwell
        DuplicateQ1Page (i)
    Next i
End Sub


' delete all summary page
Private Sub CommandButton5_Click()

    Call modWaterQualityTest.DeleteAllSummaryPage("Q1")

End Sub























Public Sub MalgunGothic()
    ActiveWindow.SmallScroll Down:=78
    Cells.Select
    Range("A200").Activate
    With Selection.Font
        .name = "���� ���"
    End With
    
    Range("C186").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub

Public Sub make_frame2(ByVal sh As Integer)
    Worksheets(CStr(sh)).Activate
    
    Range("i3").value = "���������"
    Range("i4").value = "������β�"
    Range("i5").value = "����"
    Range("i6").value = "�������"
    
    Range("I3:K6").Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    With Selection.Font
        .name = "���� ���"
        .Size = 11
    End With
    
    Range("I3:K6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("k3").Select
End Sub

Public Sub make_frame()
    Range("i3").value = "���������"
    Range("i4").value = "������β�"
    Range("i5").value = "����"
    Range("i6").value = "�������"
    
    Range("I3:K6").Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    With Selection.Font
        .name = "���� ���"
        .Size = 11
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    
    Range("I3:K6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub


Public Sub draw_motor_frame(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    Debug.Print lastRow()
    mychar = ColumnNumberToLetter(nof_sheets + 1)
    
    Range("A" & (po) & ":" & mychar & (po + 12)).Select
    
    Call draw_border
    Range("A" & (po) & ":" & mychar & (po + 1)).Select
    Call draw_border
    Range("A" & (po + 11) & ":" & mychar & (po + 12)).Select
    Call draw_border
    Range("A" & (po) & ":" & "A" & (po + 12)).Select
    Call draw_border
    
    Range("A" & (po + 2) & ":" & "A" & (po + 10)).Select
    Call draw_border
    
    Range("A" & (po) & ":B" & (po)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Range("a" & (po)).value = "�������»����ϴ°�"
    Range("a" & (po + 2)).value = "�����ɵ�"
    Range("a" & (po + 3)).value = "Q(����)-�����"
    Range("a" & (po + 4)).value = "Depth(���ͼ�ġ�ɵ�)"
    Range("a" & (po + 5)).value = "Height(������)"
    Range("a" & (po + 6)).value = "Sum (�հ�)"
    Range("a" & (po + 7)).value = "E (ȿ��)"
    Range("a" & (po + 9)).value = "����"
    Range("a" & (po + 11)).value = "�㰡������ ����"
    Range("a" & (po + 12)).value = "�̷л� ����ɷ�"
    
    Call decorationPumpHP(nof_sheets, po)
    Call decorationInerLine(nof_sheets, po)
    Call alignTitle(nof_sheets, po)
End Sub

Private Sub alignTitle(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    mychar = ColumnNumberToLetter(nof_sheets + 1)
    
    'Range("A57:B57").Select
    Range("A" & (po) & ":" & "B" & (po)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Selection.Font.Italic = False
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .Underline = xlUnderlineStyleNone
        .themeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("A59:A69").Select
    Range("A" & (po + 2) & ":" & "A" & (po + 12)).Select
    
    With Selection.Font
        .name = "���� ���"
        .Size = 11
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    Range("O65").Select
End Sub

Private Sub decorationPumpHP(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = ColumnNumberToLetter(nof_sheets + 1)
    'Range("B58:N69").Select
    Range("B" & (po + 1) & ":" & mychar & (po + 12)).Select
    
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B59:N59").Select
    Range("B" & (po + 2) & ":" & mychar & (po + 2)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .color = -16776961
        .TintAndShade = 0
    End With
    
    'Range("B60:N60").Select
    Range("B" & (po + 3) & ":" & mychar & (po + 3)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    
    'Range("B63:N63").Select
    Range("B" & (po + 6) & ":" & mychar & (po + 6)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    ActiveWindow.SmallScroll Down:=3
    
    'Range("B64:N64").Select
    Range("B" & (po + 7) & ":" & mychar & (po + 7)).Select
    Selection.NumberFormatLocal = "0.00"
    
    'Range("B67:N67").Select
    Range("B" & (po + 10) & ":" & mychar & (po + 10)).Select
    Selection.Font.Bold = True
    With Selection.Font
        .name = "���� ���"
        .Size = 14
        .Italic = True
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B68:N69").Select
    Range("B" & (po + 11) & ":" & mychar & (po + 12)).Select
    
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .color = -16776961
        .TintAndShade = 0
    End With
End Sub

Private Sub decorationInerLine(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = ColumnNumberToLetter(nof_sheets + 1)
    
    'Range("A60:N61").Select
    Range("A" & (po + 3) & ":" & mychar & (po + 4)).Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B67:N67").Select
    Range("B" & (po + 10) & ":" & mychar & (po + 10)).Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B59:N67").Select
    Range("B" & (po + 2) & ":" & mychar & (po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A59:A67").Select
    Range("A" & (po + 2) & ":" & "A" & (po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("B68:N69").Select
    Range("B" & (po + 11) & ":" & mychar & (po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A68:A69").Select
    Range("A" & (po + 11) & ":" & "A" & (po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
End Sub

Private Sub draw_border()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Option Explicit

Public ip, ip2  As Long

Private Function nColorsInArray(ByRef array_tabcolor() As Variant, ByVal CHECK As Variant) As Integer
    ' array_tabcolor :
    ' check : color��
    ' ������ �����ϴ� ������ ��� �޶�� �Ѵ�.
    ' �÷����� ������ ��� �ִ����� ����
    
    Dim i, limit    As Integer
    Dim count       As Integer: count = 0
    
    limit = UBound(array_tabcolor, 1)
    
    For i = 1 To limit
        If array_tabcolor(i) = CHECK Then
            count = count + 1
        End If
    Next i
    
    nColorsInArray = count
End Function

Private Function getans_tabcolors() As Variant
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors(), ans_tabcolors() As Variant
    
    'uc : unique colors
    Dim uc          As Integer
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim ans_tabcolors(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    ans_tabcolors(0) = 1
    
    For i = 1 To limit
        uc = nColorsInArray(arr_tabcolors, new_tabcolors(i - 1))
        ans_tabcolors(i) = ans_tabcolors(i - 1) + uc
    Next i
    
    getans_tabcolors = ans_tabcolors
End Function

Private Function getkey_tabcolors() As Object
    ' return value using by dictionary
    ' 1:(3), 2:(2), 3:(2), 4:(1)
    ' c.Add Item:=CStr(uc), key:=CStr(i)
    
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    'uc : unique colors
    Dim uc          As Integer
    
    'c colors code
    Dim C           As Collection
    Set C = New Collection
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    
    For i = 1 To limit
        uc = nColorsInArray(arr_tabcolors, new_tabcolors(i - 1))
        C.Add item:=CStr(uc), key:=CStr(i)
    Next i
    
    Set getkey_tabcolors = C
End Function

Private Sub get_tabsize(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Integer)
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
End Sub

Sub getWhpaData_AllWell()
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_sheets
        
        Call find_average2(i, 1)
        
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub getWhpaData_EachWell()
    ' 2019�� ��������Ʈ 7�������� ó��, 2019�� 6�� ����
    ' rc : return collection
    
    Dim r_ans()     As Variant
    Dim rc          As Collection
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Set rc = getkey_tabcolors()
    r_ans = getans_tabcolors()
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    Debug.Print rc(1)
    Debug.Print r_ans(0)
    Debug.Print nof_sheets
    Debug.Print nof_unique_tab
    
    ' Call find_average2(1, rc(1))
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_unique_tab
        
        sheet = r_ans(i - 1)
        Call find_average2(sheet, rc(i))
        
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub delete_allWhpaData()
    Dim n_sheets    As Long
    Dim i           As Long
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        
        Range("I3:K6").Select
        Selection.Clear
        Range("H9").Select
    Next i
End Sub


'@2023-7-15 , Refactor

Private Function get_efficiency_A(ByVal Q As Variant) As Variant
    Dim thresholds As Variant
    Dim results As Variant
    
    Dim i, result As Integer
    Dim CHECK As Boolean
    
    CHECK = True
    
    thresholds = Array(57.6, 72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(40, 42, 45, 48, 50, 52, 54, 57, 59, 61, 62, 64, 65)
    
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i - 1)
            CHECK = False
            Exit For
        End If
    Next i
        
    get_efficiency_A = result
    
    If Q < 57.6 Then get_efficiency_A = 40
    If Q > 1440 Then get_efficiency_A = 65
    
End Function



'@2023-7-15 , Refactor

Private Function get_efficiency_B(ByVal Q As Variant) As Variant
    Dim thresholds As Variant
    Dim results As Variant
    
    Dim i, result As Integer
    Dim CHECK As Boolean
    
    CHECK = True
    
    thresholds = Array(57.6, 72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(34, 36, 38, 41, 42, 44, 46, 48, 50, 52, 53, 54, 55)
    
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i - 1)
            CHECK = False
            Exit For
        End If
    Next i
        
    get_efficiency_B = result
    
    If Q < 57.6 Then get_efficiency_B = 34
    If Q > 1440 Then get_efficiency_B = 55
    
End Function



'@2023-7-15 , Refactor

Private Function get_efficiency_dongho(ByVal Q As Variant) As Variant
    Dim results As Variant
    Dim thresholds As Variant
    Dim i, result As Integer
    
    
    thresholds = Array(72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(38, 40.25, 43, 45.25, 47, 49, 51.25, 53.5, 55.5, 57, 58.25, 59.5)
        
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i)
            Exit For
        End If
    Next i
    
    
    get_efficiency_dongho = result
    If Q < 57.6 Then get_efficiency_dongho = 37
    If Q > 1440 Then get_efficiency_dongho = 60
    
End Function



Public Sub getMotorPower()
    Dim r_ans()     As Variant
    Dim rc          As Collection        'return collection
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Dim Title()     As Variant
    Dim simdo()     As Variant
    Dim pump_q()    As Variant
    Dim motor_depth() As Variant
    Dim efficiency() As Variant
    Dim Hp()        As Variant
    Dim stable_height()        As Variant
    
    Set rc = getkey_tabcolors()
    r_ans = getans_tabcolors()
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    ReDim Title(1 To nof_sheets)
    ReDim simdo(1 To nof_sheets)
    ReDim pump_q(1 To nof_sheets)
    ReDim motor_depth(1 To nof_sheets)
    ReDim efficiency(1 To nof_sheets)
    ReDim Hp(1 To nof_sheets)
    ReDim stable_height(1 To nof_sheets)
    
    ip = lastRow() + 4
    ip2 = ip + 15
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_sheets
        Worksheets(CStr(i)).Activate
        
        Title(i) = Range("b2").value
        simdo(i) = Range("c7").value
        
        ' ä����ȹ���� �����Ұ�����, ������� �����Ұ�����
        If Sheets("Recharge").cbCheSoo.value = True Then
            pump_q(i) = Range("c15").value
        Else
            pump_q(i) = Range("c16").value
        End If
        
        motor_depth(i) = Range("c18").value
        
        '2022/8/4 select efficiency
        If Sheets("Recharge").OptionButton1.value Then
          efficiency(i) = get_efficiency_A(pump_q(i))
        ElseIf Sheets("Recharge").OptionButton2.value Then
          efficiency(i) = get_efficiency_B(pump_q(i))
        Else
           efficiency(i) = get_efficiency_dongho(pump_q(i))
        End If
        
        Hp(i) = Range("c17").value
        stable_height(i) = Range("c21").value
    Next i
    
    Sheet_Recharge.Activate
    
    Call draw_motor_frame(nof_sheets, ip)
    
    For i = 1 To nof_sheets
        Call insert_basic_entry(Title(i), simdo(i), pump_q(i), motor_depth(i), efficiency(i), Hp(i), i, ip)
        Call insert_cell_function(i, ip)
    Next i
    
    
    ' -----------------------------------
    ' 2023-07-15
    ' -----------------------------------
    
    For i = 1 To nof_sheets
        Call insert_downform(pump_q(i), motor_depth(i), efficiency(i), Title(i), ip2 + i - 1, stable_height(i))
    Next i
    
    Call DecoLine(i, ip2)
    Call Range_AlignLeft(Range("G54:G94"))
    
    
    Application.ScreenUpdating = True
End Sub


' -----------------------------------
' 2025-03-03
' -----------------------------------
Sub Range_AlignLeft(ByVal rng As Range)
 rng.Select
 With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("G38").Select
End Sub

Public Sub insert_downform(pump_q As Variant, motor_simdo As Variant, e As Variant, Title As Variant, ByVal po As Integer, ByVal stable_height As Variant)
    Dim tenper As Double
    Dim sum_simdo As Double
    

    tenper = Round(motor_simdo / 10, 1)
    sum_simdo = motor_simdo + tenper
    
    Cells(po, "A").value = Title
    Cells(po, "B").value = pump_q
    Cells(po, "C").value = motor_simdo
    Cells(po, "D").value = tenper
    Cells(po, "E").value = sum_simdo
    Cells(po, "F").value = e
    Cells(po, "G").value = "-"
    Cells(po, "H").value = Round((pump_q * (motor_simdo + tenper)) / (6572.5 * (e / 100)), 4)
    Cells(po, "I").value = find_P2(Cells(po, "H").value)
    
    ' -----------------------------------
    ' 2025-02-27
    ' -----------------------------------
    Cells(po, "J").value = stable_height
    Cells(po, "J").numberFormat = "0.00"
    
    ' -----------------------------------
    ' 2025-03-2
    ' inject formula in a cell
    ' -----------------------------------
    Cells(po, "G").value = "{ " & pump_q & " TIMES " & sum_simdo & " } over { " & "6,572.5" & " TIMES " & (e / 100) & " }"

    Debug.Print "{ " & pump_q & " TIMES " & sum_simdo & " } over { " & "6,572.5" & " TIMES " & (e / 100) & " }"

End Sub




Function find_P2(ByVal num As Double) As Double
    Dim thresholds As Variant
    Dim i As Integer
    thresholds = Array(1, 2, 3, 5, 7.5, 10, 15, 20, 25, 30)
    
    find_P2 = thresholds(0)
    
    For i = 1 To UBound(thresholds)
        If num >= thresholds(i - 1) And num < thresholds(i) Then
            find_P2 = thresholds(i)
            Exit For
        End If
    Next i
End Function



Sub DecoLine(ByVal i As Integer, ByVal po As Integer)
    Rows(po & ":" & (po + i - 2)).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Selection.Font
        .name = "Arial"
        .Size = 12
        .Italic = True
    End With
    
        
    Range("A" & po & ":I" & (po + i - 2)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
        
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("C62").Select
End Sub

    
Function RoundUpNumber(ByVal num As Double)
    Dim roundedNum As Double
    roundedNum = Application.WorksheetFunction.RoundUp(num, 0)
    RoundUpNumber = roundedNum
End Function


Private Sub insert_cell_function(ByVal n As Integer, ByVal position As Integer)
    'height1 : ������
    'height : �����հ�
    
    Dim mychar
    Dim height, height1, eq, round_hp, theory_hp As String
    Dim h1, h2      As Integer
    
    h1 = position + 4
    h2 = position
    
    mychar = ColumnNumberToLetter(n + 1)
    ' Debug.Print mychar
    
    height = "=" & mychar & CStr(h1) & "+" & mychar & CStr(h1 + 1)
    height1 = "=round(" & mychar & CStr(h2 + 4) & "/10,1)"
    
    eq = "=round((" & mychar & CStr(h2 + 3) & "*" & mychar & CStr(h2 + 6) & ")/(6572.5*" & mychar & CStr(h2 + 7) & "),4)"
    round_hp = "=roundup(" & mychar & CStr(h2 + 9) & ",0)"
    theory_hp = "=round((" & mychar & CStr(h2 + 11) & "*" & mychar & CStr(h2 + 7) & "*6572.5)" & "/" & mychar & CStr(h2 + 6) & ",1)"
    
    Range(mychar & CStr(h2 + 5)).formula = height1        '������
    Range(mychar & CStr(h2 + 6)).formula = height        '�հ�
    
    Range(mychar & CStr(h2 + 9)).formula = eq
    Range(mychar & CStr(h2 + 10)).formula = round_hp
    Range(mychar & CStr(h2 + 12)).formula = theory_hp

End Sub


'ip : insertion point

Private Sub insert_basic_entry(Title As Variant, simdo As Variant, Q As Variant, motor_depth As Variant, _
                                e As Variant, Hp As Variant, ByVal i As Integer, ByVal po As Variant)
    Dim mychar As String
    
    mychar = ColumnNumberToLetter(i + 1)
    Range(mychar & CStr(po + 1)).value = Title
    Range(mychar & CStr(po + 2)).value = simdo
    Range(mychar & CStr(po + 3)).value = Q
    Range(mychar & CStr(po + 4)).value = motor_depth
    Range(mychar & CStr(po + 7)).value = e / 100
    Range(mychar & CStr(po + 11)).value = Hp
    
    Call SetFontMalgun(mychar, ip)
End Sub

Sub SetFontMalgun(ByVal col As String, ByVal ip As Integer)
    With Range("col" & CStr(ip) & ":" & col & CStr((ip + 11))).Font
        .name = "Arial"
        .Size = 12
        .Bold = True
        .Italic = True
    End With
End Sub



Option Explicit

Public Sub rows_and_column()
    Debug.Print Cells(20, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Debug.Print Range("a20").row & " , " & Range("a20").Column
    
    Range("B2:Z44").Rows(3).Select
End Sub

Public Sub ShowNumberOfRowsInSheet1Selection()
    Dim AREA        As Range
    
    ' Worksheets("Sheet1").Activate
    Dim selectedRange As Excel.Range
    Set selectedRange = Selection
    
    Dim areaCount   As Long
    areaCount = Selection.Areas.count
    
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
               Selection.Rows.count & " rows."
    Else
        Dim areaIndex As Long
        areaIndex = 1
        For Each AREA In Selection.Areas
            MsgBox "Area " & areaIndex & " of the selection contains " & _
                   AREA.Rows.count & " rows." & " Selection 2 " & Selection.Areas(2).Rows.count & " rows."
            areaIndex = areaIndex + 1
        Next
    End If
End Sub


' Refactor 2023/10/20
Function myRandBetween(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div) * SIGN
End Function

Function myRandBetween2(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div)
End Function


' Refactor 2023/10/20
Public Sub rnd_between()
    Dim i As Integer
    
    For i = 14 To 24
        Dim SIGN As Integer
        SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
        
        Cells(i, 14).value = (WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        
        With Cells(i, 14)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00"
        End With
    Next i
End Sub


Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub

'
' 2023/3/15, make summary page
' i think, this procedure made before .. but ican't the source excel file ..
' so i make again
' get page from pagenum
' quotient, remainder
' pagenum - 7 : 7/3 - 2, 1 (if remainder 1 = w1, remain = 2, w2)
'
Private Sub CommandButton3_Click()
' make summary page

    Dim n_sheets As Integer
    Dim result() As Integer
    Dim w3page, wselect, restpage As Integer
    'wselect = 1 --> only w1
    'wselect = 2 --> w1, w2
    
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    n_sheets = sheets_count()
    result = DivideWellsBy3(n_sheets)
    
    
    ' result(0) = quotient
    ' result(1) = remainder
    w3page = result(0)
    
    Select Case result(1)
        Case 0
            restpage = 0
            wselect = 0
            
        Case 1
            restpage = 1
            wselect = 1
            
        Case 2
            restpage = 1
            wselect = 2
    End Select
    
    
    Call DuplicateQ3Page(w3page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call modWaterQualityTest.DuplicateRestQ3(wselect, w3page)
    End If
    
End Sub

'
' 2024/2/28, Get Water Spec From YangSoo
'

Private Sub CommandButton2_Click()

  Call GetWaterSpecFromYangSoo_Q3

End Sub



Private Sub CommandButton4_Click()
 
 Call modWaterQualityTest.DeleteAllSummaryPage("Q3")
   
End Sub

Private Sub CommandButton5_Click()
' get ec, ph, temp

    Call DataAnalysis
End Sub


' Get(Ec, Ph, Temp) Range - ���������� ��賻�� �Լ� ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub DataAnalysis()
    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    

    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC_Q3(cellLOW, i)
        hiEC(i) = getEC_Q3(cellHI, i)
        
        lowPH(i) = getPH_Q3(cellLOW, i)
        hiPH(i) = getPH_Q3(cellHI, i)
        
        lowTEMP(i) = getTEMP_Q3(cellLOW, i)
        hiTEMP(i) = getTEMP_Q3(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"
End Sub


Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub



Private Sub CommandButton5_Click()
' Get(Ec, Ph, Temp) Range - ���������� ��賻�� �Լ� ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...

    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC_Q2(cellLOW, i)
        hiEC(i) = getEC_Q2(cellHI, i)
        
        lowPH(i) = getPH_Q2(cellLOW, i)
        hiPH(i) = getPH_Q2(cellHI, i)
        
        lowTEMP(i) = getTEMP_Q2(cellLOW, i)
        hiTEMP(i) = getTEMP_Q2(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"

End Sub



Private Sub CommandButton3_Click()
' make summary page

    Dim result() As Integer
    Dim w2page, wselect, restpage As Integer
    'wselect = 1 --> only w1
       
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
       
       
    result = DivideWellsBy2(sheets_count())
    
    ' result(0) = quotient
    ' result(1) = remainder
    
    w2page = result(0)
    restpage = result(1)
    
    Call DuplicateQ2Page(w2page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call modWaterQualityTest.DuplicateRestQ2(w2page)
    End If

End Sub


Private Sub CommandButton2_Click()
' get waterspec from yangsoo
  
  Call GetWaterSpecFromYangSoo_Q2

  
End Sub



Private Sub CommandButton4_Click()

 Call modWaterQualityTest.DeleteAllSummaryPage("Q2")

End Sub




Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c6").value
    'MsgBox (sheetDate)
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.value = nYear
    ComboBoxMonth.value = nMonth
    ComboBoxDay.value = nDay
    
    ComboBoxHour.value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.value = True
    Else
        OptionButtonAM.value = True
    End If
    
    Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.value, ComboBoxMonth.value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, nTime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.value
    nMonth = ComboBoxMonth.value
    nDay = ComboBoxDay.value
        
    nHour = ComboBoxHour.value
    nMinute = ComboBoxMinute.value
            
            
    nHour = nHour + IIf(OptionButtonPM.value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    nTime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + nTime
         
    Range("c6").value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    
End Sub

Option Explicit

Public Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub

Public Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'
'Function UppercaseString(inputString As String) As String
'    UppercaseString = UCase(inputString)
'End Function
'


Public Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column
    
    Dim lRow        As Long
    Dim lCol        As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.count, 1).End(xlUp).row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
           "Last Column: " & lCol
End Sub

Public Function lastRow() As Long
    Dim lRow        As Long
    lRow = Cells(Rows.count, 1).End(xlUp).row
    
    lastRow = lRow
End Function

'Public Function Contains(Col As Collection, key As Variant) As Boolean
'    On Error Resume Next
'    Col (key)                                    ' Just try it. If it fails, Err.Number will be nonzero.
'    Contains = (Err.number = 0)
'    Err.Clear
'End Function

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o           As Object
    On Error Resume Next
    
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function IsSheetExists(ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = sheet_name Then
            sheetExists = True
            Exit For
        End If
    Next ws
   
    
    If sheetExists Then
        CheckSheetExists = True
    Else
        CheckSheetExists = False
    End If
End Function



Function RemoveDupesDict(myArray As Variant) As Variant
    'DESCRIPTION: Removes duplicates from your array using the dictionary method.
    'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
    ' the Tools > References menu.
    ' (1.b) This is necessary because I use Early Binding in this function.
    ' Early Binding greatly enhances the speed of the function.
    ' (2) The scripting dictionary will not work on the Mac OS.
    'SOURCE: https://wellsr.com
    '-----------------------------------------------------------------------
    Dim i           As Long
    Dim d           As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    With d
        For i = LBound(myArray) To UBound(myArray)
            If IsMissing(myArray(i)) = False Then
                .item(myArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function

Public Function GetLength(A As Variant) As Integer
    ' if array is empty return 0
    ' else return number of array item
    
    If IsEmpty(A) Then
        GetLength = 0
    Else
        GetLength = UBound(A) - LBound(A) + 1
    End If
End Function

Public Function getUnique(ByRef array_tabcolor As Variant) As Variant
    ' remove duplicated item in array
    ' and return unique array value
    
    Dim array_size  As Variant
    Dim new_array   As Variant
    
    new_array = RemoveDupesDict(array_tabcolor)
    getUnique = new_array
End Function


'
' 2024,3,4 Convert to Double
' for Summary Tab

Function ConvertToDouble(inputString As String) As Double
    ConvertToDouble = CDbl(Replace(inputString, " m", ""))
End Function


Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)        'TRY to convert to an Integer value
    Exit Function        'If we reach this point, then we succeeded so exit
    
ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then        'error # 13 is Type mismatch
        Err.Clear
        ConvertToLongInteger = 0
        Exit Function
    Else
        On Error GoTo 0
        Resume
    End If
End Function


Option Explicit

'------------------------------------------------------------------------------------------
' 2022/6/11

Public Enum cellLowHi
    cellLOW = 0
    cellHI = 1
End Enum


Sub EraseCellData(ByVal str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Sub Delay(ByVal msg As String, ByVal n As Integer)
    Application.Wait (Now + TimeValue("0:00:" & n))
    MsgBox msg, vbOKOnly
End Sub


'Function GetNumberOfWell() As Integer
'    Dim save_name As String
'    Dim n As Integer
'
'    save_name = ActiveSheet.Name
'    Sheets("Well").Activate
'    Sheets("Well").Range("A30").Select
'    Selection.End(xlUp).Select
'    n = CInt(GetNumeric2(Selection.value))
'
'    GetNumberOfWell = n
'End Function

Function GetRangeStringFromSelection()
    Dim selectedRange As Range
    Dim rangeAddress As String

    ' Set the selected range to a variable
    Set selectedRange = Selection

    ' Get the address of the selected range
    rangeAddress = selectedRange.Address
    GetRangeStringFromSelection = rangeAddress
    
    ' Display the range address
    ' MsgBox "The address of the selected range is: " & rangeAddress
End Function



Function ColumnNumberToLetter(ByVal columnNumber As Integer) As String
    Dim dividend As Integer
    Dim modulo As Integer
    Dim columnName As String
    Dim result As String
    
    dividend = columnNumber
    result = ""
    
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = (dividend - modulo) \ 26
    Loop
    
    ColumnNumberToLetter = columnName
End Function


Function ColumnLetterToNumber(ByVal columnLetter As String) As Long
    Dim i As Long
    Dim result As Long

    result = 0
    For i = 1 To Len(columnLetter)
        result = result * 26 + (Asc(UCase(Mid(columnLetter, i, 1))) - 64)
    Next i

    ColumnLetterToNumber = result
End Function


Sub BackGroundFill(rngLine As Range, FLAG As Boolean)

If FLAG Then
    rngLine.Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
Else
    rngLine.Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

End Sub

Function GetRowColumn(name As String) As Variant
    Dim acColumn, acRow As Variant
    Dim result(1 To 2) As Variant

    acColumn = Split(Range(name).Address, "$")(1)
    acRow = Split(Range(name).Address, "$")(2)

    '  Row = ActiveCell.Row
    '  col = ActiveCell.Column
    
    
    result(1) = acColumn
    result(2) = acRow

    Debug.Print acColumn, acRow
    GetRowColumn = result
End Function


' �̰���, Well ���� ���� ������ �˻����ϴ°��̶�, ���̰� �����.
Function GetNumberOfWell() As Integer
    Dim save_name As String
    Dim n As Integer
    
    save_name = ActiveSheet.name
    With Sheets("Well")
        n = .Cells(.Rows.count, "A").End(xlUp).row
        n = CInt(GetNumeric2(.Cells(n, "A").value))
    End With
    
    GetNumberOfWell = n
End Function

Public Function sheets_count() As Long
    Dim i As Integer
    Dim nSheetsCount As Long
    Dim nWell As Long
    Dim strSheetsName() As String

    nSheetsCount = ThisWorkbook.Sheets.count
    nWell = 0

    ReDim strSheetsName(1 To nSheetsCount)

    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).name
        If ConvertToLongInteger(strSheetsName(i)) <> 0 Then
            nWell = nWell + 1
        End If
    Next i

    sheets_count = nWell
End Function


'BaseData_ETC : ������赥����, An_OriginalSaveFile
Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.name, Workbook.name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.name, "OriginalSaveFile") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    If Workbook Is Nothing Then
        GetOtherFileName = "Empty"
    Else
        GetOtherFileName = Workbook.name
    End If
    
End Function


Function CheckSubstring(ByVal str As String, ByVal chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function



Function ExtractNumberFromString(inputString As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "\d+"
    End With
    
    If regex.test(inputString) Then
        Set matches = regex.Execute(inputString)
        ExtractNumberFromString = matches(0)
    Else
        ExtractNumberFromString = "No numbers found"
    End If
End Function



Function GetNumeric2(ByVal CellRef As String) As String
    Dim StringLength, i  As Integer
    Dim result      As String
    
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i
    GetNumeric2 = result
End Function

'********************************************************************************************************************************************************************************
'Function Name                    : IsWorkBookOpen(ByVal OWB As String)
'Function Description             : Function to check whether specified workbook is open
'Data Parameters                  : OWB:- Specify name or path to the workbook. eg: "Book1.xlsx" or "C:\Users\Kannan.S\Desktop\Book1.xlsm"

'********************************************************************************************************************************************************************************
Function IsWorkBookOpen(ByVal OWB As String) As Boolean
    IsWorkBookOpen = False
    Dim wb          As Excel.Workbook
    Dim WBNAME      As String
    Dim WBPath      As String
    Dim OWBArray    As Variant
    
    Err.Clear
    
    On Error Resume Next
    OWBArray = Split(OWB, Application.PathSeparator)
    Set wb = Application.Workbooks(OWBArray(UBound(OWBArray)))
    WBNAME = OWBArray(UBound(OWBArray))
    WBPath = wb.Path & Application.PathSeparator & WBNAME
    
    If Not wb Is Nothing Then
        If UBound(OWBArray) > 0 Then
            If LCase(WBPath) = LCase(OWB) Then IsWorkBookOpen = True
        Else
            IsWorkBookOpen = True
        End If
    End If
    Err.Clear
    
End Function

'------------------------------------------------------------------------------------------

Public Function GetLengthByColor(ByVal tabColor As Variant) As Integer
    Dim n_sheets, i, j, nTab As Integer
    n_sheets = sheets_count()
    
    nTab = 0
    
    For i = 1 To n_sheets
        If (Sheets(CStr(i)).Tab.color = tabColor) Then
            nTab = nTab + 1
        End If
    Next i
    
    GetLengthByColor = nTab
End Function

Sub get_tabsize_by_well(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Variant, ByRef n_tabcolors As Variant)
    ' n_tabcolors : return value
    ' nof_unique_tab : return value
    
    Dim n_sheets, i, j As Integer
    Dim limit()     As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim limit(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    
    For i = 0 To UBound(new_tabcolors)
        limit(i) = GetLengthByColor(new_tabcolors(i))
    Next i
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
    n_tabcolors = new_tabcolors
End Sub
Option Explicit

Dim ColorValue(1 To 20) As Long

Public Sub InitialSetColorValue()
    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)
End Sub

Sub initialize_wellstyle()
    
    Dim rng, cell As Range

    Set rng = Range("C3:C22")
    Range("C3:C22").Select
    
    Selection.numberFormat = "General"
        
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 2024/6/15
    For Each cell In rng
        SetFontAndInteriorColorBasedOnBackground cell
    Next cell

    
    Range("E19:G19").Select
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .themeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("E21:G21").Select
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B25:K29").Select
    With Selection.Font
        .name = "���� ���"
        .Size = 11
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("d23").Select
End Sub


' 2024/6/15
Sub test_SetFontAndInteriorColorBasedOnBackground()

    Dim cell, rng As Range
    
    Set rng = ActiveSheet.Range("c7, c8, c9")

    For Each cell In rng
        SetFontAndInteriorColorBasedOnBackground cell
    Next cell

End Sub


' 2024/6/15
Function GetBackgroundColor(ByVal cell As Range) As Long
    GetBackgroundColor = cell.Interior.color
End Function


' Subroutine to set the font and interior color based on the background color
Sub SetFontAndInteriorColorBasedOnBackground(ByVal cell As Range)
    Dim bgColor As Long
    
    ' Get the background color of the cell
    bgColor = GetBackgroundColor(cell)
    
    ' Determine if the background color is dark
    If IsDarkColor(bgColor) Then
        With cell.Font
            .name = "���� ���"
            .Size = 10
            .themeColor = xlThemeColorDark1 ' Light font color for dark background
            .ThemeFont = xlThemeFontNone
        End With
    Else
        With cell.Font
            .name = "���� ���"
            .Size = 10
            .themeColor = xlThemeColorLight1 ' Dark font color for light background
            .ThemeFont = xlThemeFontNone
        End With
    End If
End Sub

' Function to determine if a color is dark
Function IsDarkColor(color As Long) As Boolean
    Dim R As Long, G As Long, B As Long
    R = (color Mod 256)
    G = ((color \ 256) Mod 256)
    B = ((color \ 65536) Mod 256)
    
    ' Calculate brightness (perceived luminance)
    ' Using the formula: 0.299*R + 0.587*G + 0.114*B
    If (0.299 * R + 0.587 * G + 0.114 * B) < 128 Then
        IsDarkColor = True
    Else
        IsDarkColor = False
    End If
End Function


' 2024/6/15
Sub DetermineThemeColor()
    Dim ws As Worksheet
    Dim cell As Range
    Dim themeColor As Long
    
    ' Set your worksheet and cell
    Set ws = ActiveSheet
    Set cell = ws.Range("c8")
    
    ' Get the theme color if it exists
    On Error Resume Next
    themeColor = cell.Interior.themeColor
    On Error GoTo 0
    
    ' Check if the theme color is valid
    If themeColor <> xlColorIndexNone Then
        MsgBox "The theme color of the cell is: " & ThemeColorName(themeColor)
    Else
        MsgBox "The cell does not have a theme color."
    End If
End Sub


' 2024/6/15
Function ThemeColorName(themeColor As Long) As String
    Select Case themeColor
        Case xlThemeColorDark1
            ThemeColorName = "Dark1"
        Case xlThemeColorLight1
            ThemeColorName = "Light1"
        Case xlThemeColorDark2
            ThemeColorName = "Dark2"
        Case xlThemeColorLight2
            ThemeColorName = "Light2"
        Case xlThemeColorAccent1
            ThemeColorName = "Accent1"
        Case xlThemeColorAccent2
            ThemeColorName = "Accent2"
        Case xlThemeColorAccent3
            ThemeColorName = "Accent3"
        Case xlThemeColorAccent4
            ThemeColorName = "Accent4"
        Case xlThemeColorAccent5
            ThemeColorName = "Accent5"
        Case xlThemeColorAccent6
            ThemeColorName = "Accent6"
        Case xlThemeColorHyperlink
            ThemeColorName = "Hyperlink"
        Case xlThemeColorFollowedHyperlink
            ThemeColorName = "Followed Hyperlink"
        Case Else
            ThemeColorName = "Unknown Theme Color"
    End Select
End Function



Private Sub change_font_size()
    Range("J25").Select
    Selection.Font.Size = 10
    Range("F26").Select
    Selection.Font.Size = 10
End Sub

Public Sub make_wellstyle()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    Call TurnOffStuff
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Call initialize_wellstyle
        Call change_font_size
    Next i
    
    Call TurnOnStuff
    
End Sub

Sub JojungData(ByVal nsheet As Integer)
    Dim nselect     As String
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    '=Well!D7
    nselect = Mid(Range("c2").formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect
    
    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    ' minhwasoo 2023/10/13
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub

Sub SetMyTabColor(ByVal index As Integer)
    If Sheets("Well").SingleColor.value Then
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .color = 192
            .TintAndShade = 0
        End With
    Else
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .color = ColorValue(index)
            .TintAndShade = 0
        End With
    End If
End Sub

'������ ��Ʈ�� ��ȸ�ϸ鼭, ���� �������� ���߾��ش�.
'
Public Sub JojungSheetData()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Cells(i + 3, "A").value = "W" & i
    Next i
    
    For i = 1 To n_sheets
        Sheets(CStr(i)).Activate
        Range("B26").value = "W-" & i
        
        Call JojungData(i)
        Call SetMyTabColor(i)
    Next i
End Sub

Option Explicit

Private Sub HideLocation_Click()
  Sheets("location").Visible = False
  Sheets("Recharge").Activate
End Sub
Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - �Ѱ����� - b13
'q2 - ��ä���� - b7
'q3 - �����ȹ�� - b15
'ratio - b11
'qq1 - 1�ܰ� �����


' Agg1_Tentative_Water_Intake : ����������� ���
'
Private Sub CommandButton2_Click()
' Collect Data
    Call TurnOffStuff
    Call modAgg1.ImportAggregateData(999, False)
    Call TurnOnStuff
End Sub


' �������� ������ �����̸�, �ҷ�����
Private Sub CommandButton3_Click()
    ' SingleWell Import
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = BaseData_ETC.GetOtherFileName
    'MsgBox WB_NAME
    
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
        
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call modAgg1.ImportAggregateData(singleWell, True)

End Sub



Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    ' Collect All Data
    Call TurnOffStuff
    Call modAgg2.GROK_ImportWellSpec(999, False)
    Call TurnOnStuff
End Sub




' �������� ������ �����̸�, �ҷ�����
Private Sub CommandButton3_Click()
    ' SingleWell Import
    ' ������ �������, ���ϰ��� ����Ʈ �ؾ� �Ұ�쿡 ....
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = BaseData_ETC.GetOtherFileName
    'MsgBox WB_NAME
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call modAgg2.GROK_ImportWellSpec(singleWell, True)

End Sub







Option Explicit

'0 : skin factor
'1 : Re1
'2 : Re2
'3 : Re3

Public Enum ER_VALUE
    erRE0 = 0
    erRE1 = 1
    erRE2 = 2
    erRE3 = 3
End Enum

Function GetER_Mode(ByVal WB_NAME As String) As Integer
    Dim Er, R       As String
    
    ' er = Range("h10").value
    Er = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("h10").value
    'MsgBox er
    R = Mid(Er, 5, 1)
    
    If R = "F" Then
        GetER_Mode = 0
    Else
        GetER_Mode = val(R)
    End If
End Function



Function GetEffectiveRadius(ByVal WB_NAME As String) As Double
    Dim i, Er As Integer
    
    If Not IsWorkBookOpen(WB_NAME) Then
        MsgBox "Please open the yangsoo data ! " & WB_NAME
        Exit Function
    End If
    
    Er = GetER_Mode(WB_NAME)
    'Worksheets("SkinFactor").Range("k8").value  - ����� 1�� (RE1)
    'Worksheets("SkinFactor").Range("k9").value  - ����� 2�� (RE2)
    'Worksheets("SkinFactor").Range("k10").value  - ����� 3�� (RE3)
    
    
    Select Case Er
        Case erRE1
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k8").value
        Case erRE2
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k9").value
        Case erRE3
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k10").value
        Case Else
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("C8").value
    End Select

End Function


Function GetER_ModeFX(ByVal well_no As Integer) As Integer
    Dim Er, R  As String
    Dim wsYangSoo As Worksheet
    
    Set wsYangSoo = Worksheets("YangSoo")
    
    ' ak : ER Mode
    Er = wsYangSoo.Cells(4 + well_no, "ak").value
    
    'MsgBox er
    R = Mid(Er, 5, 1)
    
    If R = "F" Then
        GetER_ModeFX = 0
    Else
        GetER_ModeFX = val(R)
    End If
End Function



Function GetEffectiveRadiusFromFX(ByVal well_no As Integer) As Double
    Dim i, Er As Integer
    Dim wsYangSoo As Worksheet
    
    Set wsYangSoo = Worksheets("YangSoo")
    
    Er = GetER_ModeFX(well_no)
    i = well_no
    
    'Worksheets("SkinFactor").Range("k8").value  - ����� 1�� (RE1)
    'Worksheets("SkinFactor").Range("k9").value  - ����� 2�� (RE2)
    'Worksheets("SkinFactor").Range("k10").value  - ����� 3�� (RE3)
    
    Select Case Er
        Case erRE1
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AL").value
        Case erRE2
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AM").value
        Case erRE3
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AN").value
        Case Else
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "Z").value
    End Select

End Function

Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub



Private Sub CommandButton2_Click()
' Collect Data

    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q, DaeSoo, T1, S1, direction, gradient As Double
    
    nofwell = sheets_count()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    Call EraseCellData("C4:O34")
    
    TurnOffStuff
    
    For i = 1 To nofwell
        Q = Sheets(CStr(i)).Range("c16").value
        DaeSoo = Sheets(CStr(i)).Range("c14").value
        
        T1 = Sheets(CStr(i)).Range("e7").value
        S1 = Sheets(CStr(i)).Range("g7").value
        
        direction = getDirectionFromWell(i)
        gradient = Sheets(CStr(i)).Range("k18").value
        
        Call modAggWhpa.WriteWellData_Single(Q, DaeSoo, T1, S1, direction, gradient, i)
    Next i
    
    Sheets("aggWhpa").Select
    
    Call MakeAverageAndMergeCells(nofwell)
    Call DrawOutline
    TurnOnStuff
    
End Sub





Option Explicit


' AggSum_Intake : �����ȹ��
' AggSum_Simdo : �����ɵ�
' AggSum_MotorHP : ��������
' AggSum_NaturalLevel : �ڿ�����
' AggSum_StableLevel : ��������
' AggSum_ToChool : ���ⱸ��
' AggSum_MotorSimdo : ���ͽɵ�
'

' AggSum_ROI : radius of influence
' AggSum_DI : drastic index
' AggSum_ROI_Stat :
' AggSum_26_AC : 26, AquiferCharacterization
' AggSum_26_RightAC : 26, Right AquiferCharacterizationn

Const WELL_BUFFER = 30


Private Sub CommandButton1_Click()
    Sheets("AggSum").Visible = False
    Sheets("Well").Select
End Sub


' Summary Button
Private Sub CommandButton2_Click()
   Call Summary_Button
End Sub


' 2025-2-12, CheckBoxClick �߰����� ...
Private Sub CheckBox1_Click()
  Application.Run "Sheet_AggSum.Summary_Button"
End Sub


' 2025-2-12, CheckBoxClick �߰����� ...
Sub Summary_Button()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggSum" Then Sheets("AggSum").Select

    ' Summary, Aquifer Characterization  Appropriated Water Analysis
    BaseData_ETC_02.TurnOffStuff
    
    Call Write23_SummaryDevelopmentPotential
    Call Write26_AquiferCharacterization(nofwell)
    Call Write26_Right_AquiferCharacterization(nofwell)
    
    Call Write_RadiusOfInfluence(nofwell)
    Call Write_WaterIntake(nofwell)
    Call Check_DI
    
    Call Write_DiggingDepth(nofwell)
    Call Write_MotorPower(nofwell)
    Call Write_DrasticIndex(nofwell)
    
    Call Write_NaturalLevel(nofwell)
    Call Write_StableLevel(nofwell)
    
    
    Call Write_MotorTochool(nofwell)
    Call Write_MotorSimdo(nofwell)
    
    Call Write_WellDiameter(nofwell)
    Call Write_CasingDepth(nofwell)


    Range("D5").Select
    BaseData_ETC_02.TurnOnStuff
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
Option Explicit



Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggStep").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data
    Call TurnOffStuff
    Call WriteStepTestData(999, False)
    Call TurnOnStuff
End Sub



Private Sub CommandButton3_Click()
'Single Well Import

'single well import

Dim singleWell  As Integer
Dim WB_NAME As String



' �������� ������ �����̸�, �ҷ�����
WB_NAME = BaseData_ETC.GetOtherFileName
'MsgBox WB_NAME

'If Workbook Is Nothing Then
'    GetOtherFileName = "Empty"
'Else
'    GetOtherFileName = Workbook.name
'End If
    
If WB_NAME = "Empty" Then
    MsgBox "WorkBook is Empty"
    Exit Sub
Else
    singleWell = CInt(ExtractNumberFromString(WB_NAME))
'   MsgBox (SingleWell)
End If

Call WriteStepTestData(singleWell, True)

End Sub


Option Explicit
'Sheet_AggChart


Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data

   Call TurnOffStuff
   Call WriteAllCharts(999, False)
   Call TurnOnStuff

End Sub


Private Sub CommandButton3_Click()
    'single well import
    
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
            
    
    ' �������� ������ �����̸�, �ҷ�����
    WB_NAME = BaseData_ETC.GetOtherFileName
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call WriteAllCharts(singleWell, True)

End Sub




Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Sub DeleteAllCharts()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Sub DeleteAllImages(ByVal singleWell As Integer)
    Dim ws As Worksheet
    Dim sh As Shape
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    
    If singleWell = 999 Then
        For Each sh In ws.Shapes
            If sh.Type = msoPicture Then
                sh.Delete
            End If
        Next sh
    End If
    
End Sub

Sub WriteAllCharts(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
'AggChart ChartImport

    Dim fName, source_name As String
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggChart" Then Sheets("AggChart").Select
    
    ' Call DeleteAllCharts
    
    Call TurnOffStuff
    If isSingleWellImport Then
        Call DeleteAllImages(singleWell)
    Else
        Call DeleteAllImages(999)
    End If
    
    
    source_name = ActiveWorkbook.name
    
    
    Call TurnOffStuff
    
    For i = 1 To nofwell
    
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
        
SINGLE_ITERATION:
        Call Write_InsertChart(i, source_name)
        
NEXT_ITERATION:
    Next i
    
    Call TurnOnStuff
End Sub

Sub Write_InsertChart(well As Integer, source_name As String)
    Dim fName As String
    Dim imagePath As String

    imagePath = Environ("TEMP") & "\tempChartImage.png"

    fName = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data ! " & fName
        Exit Sub
    End If

    Call SaveAndInsertChart(well, source_name, "Chart 5", "d" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 7", "j" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 9", "p" & CStr(3 + 16 * (well - 1)))
End Sub


Sub SaveAndInsertChart(well As Integer, source_name As String, chartName As String, targetRange As String)
    Dim imagePath As String
    Dim fName As String
    Dim targetCell As Range
    Dim picWidth As Double, picHeight As Double
    
    imagePath = Environ("TEMP") & "\tempChartImage.png"
    fName = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"

    Windows(fName).Activate
    Worksheets("Input").ChartObjects(chartName).Activate
    ActiveChart.Export fileName:=imagePath, FilterName:="PNG"
    
    With ActiveChart.Parent
        picWidth = .Width
        picHeight = .height
    End With
    

    Windows(source_name).Activate
    Set targetCell = Sheets("AggChart").Range(targetRange)
        
    
    Sleep (1000)
    Sheets("AggChart").Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=picWidth, _
        height:=picHeight
        
End Sub


Sub ActivateChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    Set ws = ThisWorkbook.Worksheets("Input")
    Set chartObj = ws.ChartObjects("Chart 5")
    chartObj.Activate
End Sub



Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
  'Collect Data
    Dim fName As String
    
    fName = "A1_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "YangSoo File Does not OPEN ... ! " & fName
        Exit Sub
    End If
    
    Call TurnOffStuff
    Call GetBaseDataFromYangSoo(999, False)
    Call TurnOnStuff
End Sub


Private Sub CommandButton3_Click()
    ' Write Formula Button
       
       Call WriteFormula
    ' End of Write Formula Button
End Sub


Private Sub CommandButton4_Click()
    'single well import
    
    Dim WellNumber  As Integer
    Dim WB_NAME As String
    
    
    ' �������� ������ �����̸�, �ҷ�����
    WB_NAME = BaseData_ETC.GetOtherFileName
    
    'MsgBox WB_NAME
        
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
        
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        WellNumber = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call GetBaseDataFromYangSoo(WellNumber, True)

End Sub






Private Sub CommandButton1_Click()
   Sheets("water").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Dim WB_NAME, cpRange  As String

    If Workbooks.count = 1 Then
        MsgBox "Please Open �������, ���� ", vbOKOnly
        Exit Sub
    End If
   
    ' ������� ������ �ҷ����� ���� ����
    WB_NAME = Sheet4_Water.GetOtherFileName
    
    cpRange = GetCopyPoint(WB_NAME)
    Call CopyFromGWAN_JUNG(WB_NAME, cpRange)
    Call FormulaInjection
    
End Sub

' 2024-01-14
' inject formula ...

Private Sub FormulaInjection()
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    For i = 4 To nofwell + 3
        Sheets("Well").Cells(i, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i

End Sub



'�⺻���������͸� ������ �´�.
Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.name, Workbook.name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.name, "����") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    GetOtherFileName = Workbook.name
End Function


'
'Function lastRowByKey(cell As String) As Long
'    lastRowByKey = Range(cell).End(xlDown).Row
'End Function


Function GetCopyPoint(ByVal fName As String) As String

  Dim ip1, ip2 As Integer

  ip1 = Workbooks(fName).Worksheets("ss").Range("b1").End(xlDown).row + 4
  ip2 = ip1 + 2
  
  GetCopyPoint = "B" & ip1 & ":J" & ip2
  ThisWorkbook.Activate

End Function


Sub CopyFromGWAN_JUNG(ByVal fName As String, ByVal cpRange As String)

    Workbooks(fName).Worksheets("ss").Activate
    Workbooks(fName).Worksheets("ss").Range(cpRange).Select
    Selection.Copy
    
    ThisWorkbook.Sheets("water").Activate
    
    Range("d6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.name & vbCrLf
    Next
    
    Cells(1, 1).value = workbookNames
End Sub


Sub DumpRangeToArrayAndSaveTest()
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.value
    
    ' Save array to a file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    SaveArrayToFileByExcelForm myArray, filePath
    
End Sub



Sub SaveArrayToFileByExcelForm(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, "myArray(" & i & ", " & j & ") = ";
            
            ' Separate values with a comma (CSV format)
            If j <= UBound(myArray, 2) Then
                Print #FileNum, myArray(i, j);
            End If
            
            Print #FileNum, ""
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Close FileNum
End Sub


Sub importFromArray()
    Dim myArray As Variant
    Dim rng As Range
    
    indexString = "data_" & UCase(Range("s11").value)
    
    myArray = Application.Run(indexString)
    
    
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    rng.value = myArray
       
End Sub





'***********************
'Year 2024
'***********************
'data_GEUMSAN
'data_BORYUNG
'data_DAEJEON
'data_BUYEO
'data_SEOSAN
'data_CHEONAN
'data_CHEUNGJU
'***********************
'data_HONGSUNG
'data_SEJONG
'***********************


Function data_TEMP() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_TEMP = myArray

End Function

Function data_CHUNGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1995
    myArray(1, 2) = 11.9
    myArray(1, 3) = 6.8
    myArray(1, 4) = 29.3
    myArray(1, 5) = 46.2
    myArray(1, 6) = 45.5
    myArray(1, 7) = 19
    myArray(1, 8) = 290.8
    myArray(1, 9) = 802
    myArray(1, 10) = 16
    myArray(1, 11) = 23
    myArray(1, 12) = 23.9
    myArray(1, 13) = 4.2
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 27.8
    myArray(2, 3) = 1.6
    myArray(2, 4) = 101.7
    myArray(2, 5) = 36.5
    myArray(2, 6) = 29.1
    myArray(2, 7) = 203.5
    myArray(2, 8) = 207
    myArray(2, 9) = 126
    myArray(2, 10) = 26.5
    myArray(2, 11) = 83
    myArray(2, 12) = 74.1
    myArray(2, 13) = 22.2
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 5
    myArray(3, 3) = 44.3
    myArray(3, 4) = 23.2
    myArray(3, 5) = 60
    myArray(3, 6) = 193.5
    myArray(3, 7) = 147.8
    myArray(3, 8) = 308
    myArray(3, 9) = 179.9
    myArray(3, 10) = 52.5
    myArray(3, 11) = 35.6
    myArray(3, 12) = 128.5
    myArray(3, 13) = 41.4
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 22.5
    myArray(4, 3) = 26.4
    myArray(4, 4) = 33.8
    myArray(4, 5) = 145.3
    myArray(4, 6) = 90
    myArray(4, 7) = 217.7
    myArray(4, 8) = 286.6
    myArray(4, 9) = 541.7
    myArray(4, 10) = 183
    myArray(4, 11) = 64.5
    myArray(4, 12) = 38.5
    myArray(4, 13) = 2.5
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 3.1
    myArray(5, 3) = 1.9
    myArray(5, 4) = 54.2
    myArray(5, 5) = 96
    myArray(5, 6) = 103.2
    myArray(5, 7) = 168.5
    myArray(5, 8) = 112.7
    myArray(5, 9) = 299.6
    myArray(5, 10) = 239.6
    myArray(5, 11) = 195.1
    myArray(5, 12) = 34.5
    myArray(5, 13) = 7
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 47.5
    myArray(6, 3) = 3.3
    myArray(6, 4) = 14.5
    myArray(6, 5) = 42.5
    myArray(6, 6) = 54
    myArray(6, 7) = 248.5
    myArray(6, 8) = 259.5
    myArray(6, 9) = 260.5
    myArray(6, 10) = 260.8
    myArray(6, 11) = 25.5
    myArray(6, 12) = 35.5
    myArray(6, 13) = 17.5
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 47
    myArray(7, 3) = 52.2
    myArray(7, 4) = 8.3
    myArray(7, 5) = 12
    myArray(7, 6) = 4.6
    myArray(7, 7) = 238.9
    myArray(7, 8) = 241.6
    myArray(7, 9) = 82.6
    myArray(7, 10) = 13.8
    myArray(7, 11) = 82.1
    myArray(7, 12) = 4.4
    myArray(7, 13) = 10.6
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 49.6
    myArray(8, 3) = 4.2
    myArray(8, 4) = 28.5
    myArray(8, 5) = 151.2
    myArray(8, 6) = 105
    myArray(8, 7) = 74.7
    myArray(8, 8) = 190.2
    myArray(8, 9) = 653
    myArray(8, 10) = 92.6
    myArray(8, 11) = 52.2
    myArray(8, 12) = 9.8
    myArray(8, 13) = 58.6
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 17.2
    myArray(9, 3) = 59.2
    myArray(9, 4) = 58.2
    myArray(9, 5) = 170.8
    myArray(9, 6) = 117.8
    myArray(9, 7) = 152.1
    myArray(9, 8) = 382.8
    myArray(9, 9) = 314.7
    myArray(9, 10) = 268.1
    myArray(9, 11) = 27.5
    myArray(9, 12) = 55
    myArray(9, 13) = 17.8
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 16.6
    myArray(10, 3) = 32
    myArray(10, 4) = 29.4
    myArray(10, 5) = 81
    myArray(10, 6) = 124.9
    myArray(10, 7) = 335
    myArray(10, 8) = 410.7
    myArray(10, 9) = 192.2
    myArray(10, 10) = 144.1
    myArray(10, 11) = 1.4
    myArray(10, 12) = 32.5
    myArray(10, 13) = 25.4
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 2.8
    myArray(11, 3) = 20.8
    myArray(11, 4) = 43.1
    myArray(11, 5) = 63.1
    myArray(11, 6) = 53.9
    myArray(11, 7) = 178.7
    myArray(11, 8) = 381.6
    myArray(11, 9) = 226.1
    myArray(11, 10) = 320
    myArray(11, 11) = 63.4
    myArray(11, 12) = 15.7
    myArray(11, 13) = 11.7
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 27.1
    myArray(12, 3) = 34.9
    myArray(12, 4) = 5.9
    myArray(12, 5) = 91.8
    myArray(12, 6) = 95.1
    myArray(12, 7) = 128.5
    myArray(12, 8) = 666.9
    myArray(12, 9) = 71.5
    myArray(12, 10) = 21.7
    myArray(12, 11) = 23.1
    myArray(12, 12) = 53.1
    myArray(12, 13) = 14.3
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 5.5
    myArray(13, 3) = 38.5
    myArray(13, 4) = 112.7
    myArray(13, 5) = 18.3
    myArray(13, 6) = 116.5
    myArray(13, 7) = 90.1
    myArray(13, 8) = 282.7
    myArray(13, 9) = 366
    myArray(13, 10) = 332.7
    myArray(13, 11) = 32.8
    myArray(13, 12) = 22
    myArray(13, 13) = 21.4
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 29.3
    myArray(14, 3) = 8.2
    myArray(14, 4) = 43.1
    myArray(14, 5) = 31.5
    myArray(14, 6) = 70.9
    myArray(14, 7) = 78.1
    myArray(14, 8) = 319.8
    myArray(14, 9) = 192.5
    myArray(14, 10) = 71.1
    myArray(14, 11) = 16
    myArray(14, 12) = 10.3
    myArray(14, 13) = 11.7
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 16.7
    myArray(15, 3) = 15.8
    myArray(15, 4) = 52
    myArray(15, 5) = 30.7
    myArray(15, 6) = 97.1
    myArray(15, 7) = 89.5
    myArray(15, 8) = 316.2
    myArray(15, 9) = 142.5
    myArray(15, 10) = 70.6
    myArray(15, 11) = 45
    myArray(15, 12) = 31.2
    myArray(15, 13) = 29.5
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 44.3
    myArray(16, 3) = 70.8
    myArray(16, 4) = 85.3
    myArray(16, 5) = 69.5
    myArray(16, 6) = 97
    myArray(16, 7) = 50.6
    myArray(16, 8) = 112.2
    myArray(16, 9) = 345.1
    myArray(16, 10) = 287.8
    myArray(16, 11) = 21
    myArray(16, 12) = 14.3
    myArray(16, 13) = 14.4
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 2.7
    myArray(17, 3) = 45.9
    myArray(17, 4) = 30.6
    myArray(17, 5) = 157.8
    myArray(17, 6) = 187.7
    myArray(17, 7) = 452.6
    myArray(17, 8) = 603.9
    myArray(17, 9) = 289.4
    myArray(17, 10) = 158.6
    myArray(17, 11) = 61.5
    myArray(17, 12) = 67
    myArray(17, 13) = 15.6
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 9.6
    myArray(18, 3) = 1.7
    myArray(18, 4) = 66.4
    myArray(18, 5) = 84.5
    myArray(18, 6) = 61
    myArray(18, 7) = 58.8
    myArray(18, 8) = 265.7
    myArray(18, 9) = 403.3
    myArray(18, 10) = 177.2
    myArray(18, 11) = 62
    myArray(18, 12) = 48
    myArray(18, 13) = 52.1
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 40.5
    myArray(19, 3) = 36.9
    myArray(19, 4) = 48
    myArray(19, 5) = 84.7
    myArray(19, 6) = 92.5
    myArray(19, 7) = 126.6
    myArray(19, 8) = 240.7
    myArray(19, 9) = 222.2
    myArray(19, 10) = 122.2
    myArray(19, 11) = 12.1
    myArray(19, 12) = 44
    myArray(19, 13) = 32.2
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 14
    myArray(20, 3) = 18.9
    myArray(20, 4) = 37.7
    myArray(20, 5) = 39.6
    myArray(20, 6) = 26.3
    myArray(20, 7) = 63.3
    myArray(20, 8) = 92.6
    myArray(20, 9) = 284.3
    myArray(20, 10) = 122.7
    myArray(20, 11) = 153.8
    myArray(20, 12) = 23.5
    myArray(20, 13) = 22.9
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 15.6
    myArray(21, 3) = 22.8
    myArray(21, 4) = 31.7
    myArray(21, 5) = 88.9
    myArray(21, 6) = 23
    myArray(21, 7) = 75
    myArray(21, 8) = 181.6
    myArray(21, 9) = 71.8
    myArray(21, 10) = 33.8
    myArray(21, 11) = 60.2
    myArray(21, 12) = 89.9
    myArray(21, 13) = 37.5
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 1.8
    myArray(22, 3) = 50.1
    myArray(22, 4) = 11.9
    myArray(22, 5) = 97.3
    myArray(22, 6) = 70
    myArray(22, 7) = 38.9
    myArray(22, 8) = 374.4
    myArray(22, 9) = 44
    myArray(22, 10) = 60.8
    myArray(22, 11) = 102.6
    myArray(22, 12) = 22.9
    myArray(22, 13) = 42.4
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 18
    myArray(23, 3) = 36.2
    myArray(23, 4) = 22.5
    myArray(23, 5) = 71.4
    myArray(23, 6) = 32.2
    myArray(23, 7) = 43.7
    myArray(23, 8) = 464.3
    myArray(23, 9) = 257.9
    myArray(23, 10) = 62.4
    myArray(23, 11) = 21.2
    myArray(23, 12) = 17.6
    myArray(23, 13) = 25.5
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 14.3
    myArray(24, 3) = 35.8
    myArray(24, 4) = 75.1
    myArray(24, 5) = 107.9
    myArray(24, 6) = 180
    myArray(24, 7) = 63.7
    myArray(24, 8) = 149.1
    myArray(24, 9) = 353.3
    myArray(24, 10) = 184.9
    myArray(24, 11) = 96
    myArray(24, 12) = 50
    myArray(24, 13) = 39
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 4.1
    myArray(25, 3) = 29
    myArray(25, 4) = 27.9
    myArray(25, 5) = 58.5
    myArray(25, 6) = 15.4
    myArray(25, 7) = 59.6
    myArray(25, 8) = 161.4
    myArray(25, 9) = 102.6
    myArray(25, 10) = 165.9
    myArray(25, 11) = 59
    myArray(25, 12) = 84.6
    myArray(25, 13) = 27.5
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 60.2
    myArray(26, 3) = 61.9
    myArray(26, 4) = 20.7
    myArray(26, 5) = 25.9
    myArray(26, 6) = 109.7
    myArray(26, 7) = 112.1
    myArray(26, 8) = 352.2
    myArray(26, 9) = 505.6
    myArray(26, 10) = 146.4
    myArray(26, 11) = 10.7
    myArray(26, 12) = 30
    myArray(26, 13) = 11.1
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 13.6
    myArray(27, 3) = 12.3
    myArray(27, 4) = 80.5
    myArray(27, 5) = 63.4
    myArray(27, 6) = 178.4
    myArray(27, 7) = 130.4
    myArray(27, 8) = 310.7
    myArray(27, 9) = 239.9
    myArray(27, 10) = 240.3
    myArray(27, 11) = 45.5
    myArray(27, 12) = 44.9
    myArray(27, 13) = 5.6
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 2
    myArray(28, 3) = 4.3
    myArray(28, 4) = 79.9
    myArray(28, 5) = 45.8
    myArray(28, 6) = 8.6
    myArray(28, 7) = 219
    myArray(28, 8) = 350.5
    myArray(28, 9) = 457.3
    myArray(28, 10) = 102
    myArray(28, 11) = 96.5
    myArray(28, 12) = 79.5
    myArray(28, 13) = 18.8
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 23
    myArray(29, 3) = 3.3
    myArray(29, 4) = 16.7
    myArray(29, 5) = 32.1
    myArray(29, 6) = 123.8
    myArray(29, 7) = 239.6
    myArray(29, 8) = 554.2
    myArray(29, 9) = 248
    myArray(29, 10) = 242.5
    myArray(29, 11) = 31.3
    myArray(29, 12) = 50.8
    myArray(29, 13) = 96.3
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 34.6
    myArray(30, 3) = 86.6
    myArray(30, 4) = 50.8
    myArray(30, 5) = 59
    myArray(30, 6) = 128.7
    myArray(30, 7) = 126.1
    myArray(30, 8) = 503.2
    myArray(30, 9) = 71.4
    myArray(30, 10) = 211.4
    myArray(30, 11) = 122.8
    myArray(30, 12) = 49.1
    myArray(30, 13) = 4.3
    

    data_CHUNGJU = myArray

End Function


Function data_CHUPUNGNYEONG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 16.7
    myArray(1, 3) = 15.7
    myArray(1, 4) = 55.6
    myArray(1, 5) = 67.5
    myArray(1, 6) = 37
    myArray(1, 7) = 43.8
    myArray(1, 8) = 132.2
    myArray(1, 9) = 472.4
    myArray(1, 10) = 56.9
    myArray(1, 11) = 22
    myArray(1, 12) = 29.1
    myArray(1, 13) = 5.1
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 19.4
    myArray(2, 3) = 2
    myArray(2, 4) = 117.2
    myArray(2, 5) = 28.1
    myArray(2, 6) = 58.9
    myArray(2, 7) = 463.9
    myArray(2, 8) = 118.2
    myArray(2, 9) = 89.4
    myArray(2, 10) = 18.1
    myArray(2, 11) = 73.6
    myArray(2, 12) = 58.9
    myArray(2, 13) = 24.3
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 15.2
    myArray(3, 3) = 32.8
    myArray(3, 4) = 27.9
    myArray(3, 5) = 49.7
    myArray(3, 6) = 148.7
    myArray(3, 7) = 205.2
    myArray(3, 8) = 249.2
    myArray(3, 9) = 154.4
    myArray(3, 10) = 30
    myArray(3, 11) = 4.2
    myArray(3, 12) = 140.7
    myArray(3, 13) = 44.6
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 28.2
    myArray(4, 3) = 49.6
    myArray(4, 4) = 30.9
    myArray(4, 5) = 201.5
    myArray(4, 6) = 87.8
    myArray(4, 7) = 227
    myArray(4, 8) = 244.8
    myArray(4, 9) = 368.9
    myArray(4, 10) = 282.7
    myArray(4, 11) = 49.9
    myArray(4, 12) = 15.2
    myArray(4, 13) = 4.4
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 2.8
    myArray(5, 3) = 20.1
    myArray(5, 4) = 86.7
    myArray(5, 5) = 76.4
    myArray(5, 6) = 113.2
    myArray(5, 7) = 167.3
    myArray(5, 8) = 143.3
    myArray(5, 9) = 240.1
    myArray(5, 10) = 294.5
    myArray(5, 11) = 93.5
    myArray(5, 12) = 16.3
    myArray(5, 13) = 15.3
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 29.3
    myArray(6, 3) = 5
    myArray(6, 4) = 26
    myArray(6, 5) = 43.4
    myArray(6, 6) = 26.1
    myArray(6, 7) = 155.7
    myArray(6, 8) = 293.4
    myArray(6, 9) = 318.5
    myArray(6, 10) = 304.6
    myArray(6, 11) = 30.4
    myArray(6, 12) = 51
    myArray(6, 13) = 11.1
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 47.7
    myArray(7, 3) = 63.9
    myArray(7, 4) = 9.7
    myArray(7, 5) = 16.5
    myArray(7, 6) = 39.6
    myArray(7, 7) = 202.6
    myArray(7, 8) = 154.2
    myArray(7, 9) = 23.4
    myArray(7, 10) = 112.9
    myArray(7, 11) = 116.5
    myArray(7, 12) = 10.7
    myArray(7, 13) = 24.5
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 71.9
    myArray(8, 3) = 8.9
    myArray(8, 4) = 54.4
    myArray(8, 5) = 188.2
    myArray(8, 6) = 125.3
    myArray(8, 7) = 49.3
    myArray(8, 8) = 204.3
    myArray(8, 9) = 597.1
    myArray(8, 10) = 55.5
    myArray(8, 11) = 39.6
    myArray(8, 12) = 19.1
    myArray(8, 13) = 46.1
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 20.9
    myArray(9, 3) = 54.8
    myArray(9, 4) = 50.6
    myArray(9, 5) = 182.7
    myArray(9, 6) = 178.9
    myArray(9, 7) = 157.3
    myArray(9, 8) = 548.3
    myArray(9, 9) = 330.6
    myArray(9, 10) = 222.9
    myArray(9, 11) = 24
    myArray(9, 12) = 47.2
    myArray(9, 13) = 17.1
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 16
    myArray(10, 3) = 24.9
    myArray(10, 4) = 20.6
    myArray(10, 5) = 70.8
    myArray(10, 6) = 112.5
    myArray(10, 7) = 249.1
    myArray(10, 8) = 391.5
    myArray(10, 9) = 317.6
    myArray(10, 10) = 175.1
    myArray(10, 11) = 2.6
    myArray(10, 12) = 40.2
    myArray(10, 13) = 23.3
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 12.8
    myArray(11, 3) = 35.9
    myArray(11, 4) = 50.2
    myArray(11, 5) = 31.3
    myArray(11, 6) = 47
    myArray(11, 7) = 131.2
    myArray(11, 8) = 252.3
    myArray(11, 9) = 291.8
    myArray(11, 10) = 107.7
    myArray(11, 11) = 13
    myArray(11, 12) = 20.6
    myArray(11, 13) = 14.6
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 19.4
    myArray(12, 3) = 30.4
    myArray(12, 4) = 8.7
    myArray(12, 5) = 89.5
    myArray(12, 6) = 102.5
    myArray(12, 7) = 128
    myArray(12, 8) = 697.6
    myArray(12, 9) = 43
    myArray(12, 10) = 36.9
    myArray(12, 11) = 36
    myArray(12, 12) = 61.4
    myArray(12, 13) = 19.7
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 9.8
    myArray(13, 3) = 45.9
    myArray(13, 4) = 85
    myArray(13, 5) = 24.2
    myArray(13, 6) = 73
    myArray(13, 7) = 152.1
    myArray(13, 8) = 209.5
    myArray(13, 9) = 267.9
    myArray(13, 10) = 386.1
    myArray(13, 11) = 20.4
    myArray(13, 12) = 7.2
    myArray(13, 13) = 29.9
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 40.5
    myArray(14, 3) = 6.1
    myArray(14, 4) = 27.7
    myArray(14, 5) = 47.2
    myArray(14, 6) = 57.6
    myArray(14, 7) = 172.6
    myArray(14, 8) = 152.2
    myArray(14, 9) = 172.9
    myArray(14, 10) = 59
    myArray(14, 11) = 56.5
    myArray(14, 12) = 17
    myArray(14, 13) = 9.2
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 12
    myArray(15, 3) = 34.7
    myArray(15, 4) = 37.2
    myArray(15, 5) = 34
    myArray(15, 6) = 112.1
    myArray(15, 7) = 87.4
    myArray(15, 8) = 436.9
    myArray(15, 9) = 97.7
    myArray(15, 10) = 54.7
    myArray(15, 11) = 16.5
    myArray(15, 12) = 52.7
    myArray(15, 13) = 34.8
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 22.5
    myArray(16, 3) = 72.6
    myArray(16, 4) = 88.6
    myArray(16, 5) = 54.6
    myArray(16, 6) = 115.3
    myArray(16, 7) = 38.2
    myArray(16, 8) = 201.7
    myArray(16, 9) = 443.2
    myArray(16, 10) = 150
    myArray(16, 11) = 22.3
    myArray(16, 12) = 15.1
    myArray(16, 13) = 36.3
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 3.6
    myArray(17, 3) = 47.6
    myArray(17, 4) = 20.5
    myArray(17, 5) = 95.5
    myArray(17, 6) = 163.7
    myArray(17, 7) = 187.5
    myArray(17, 8) = 284.7
    myArray(17, 9) = 369.7
    myArray(17, 10) = 59.9
    myArray(17, 11) = 58.5
    myArray(17, 12) = 95.8
    myArray(17, 13) = 14.8
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 19.8
    myArray(18, 3) = 1.6
    myArray(18, 4) = 93.8
    myArray(18, 5) = 90.6
    myArray(18, 6) = 31.3
    myArray(18, 7) = 73.8
    myArray(18, 8) = 228.6
    myArray(18, 9) = 490.1
    myArray(18, 10) = 282.6
    myArray(18, 11) = 47
    myArray(18, 12) = 51
    myArray(18, 13) = 55.6
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 42.7
    myArray(19, 3) = 38.4
    myArray(19, 4) = 48.6
    myArray(19, 5) = 79.2
    myArray(19, 6) = 80.8
    myArray(19, 7) = 119.7
    myArray(19, 8) = 186.6
    myArray(19, 9) = 106.1
    myArray(19, 10) = 94.8
    myArray(19, 11) = 48
    myArray(19, 12) = 48.1
    myArray(19, 13) = 27.6
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 6.7
    myArray(20, 3) = 13.8
    myArray(20, 4) = 93.9
    myArray(20, 5) = 100.6
    myArray(20, 6) = 20.8
    myArray(20, 7) = 102.9
    myArray(20, 8) = 74.8
    myArray(20, 9) = 402.3
    myArray(20, 10) = 88.8
    myArray(20, 11) = 121.1
    myArray(20, 12) = 78.7
    myArray(20, 13) = 32
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 34.3
    myArray(21, 3) = 30.4
    myArray(21, 4) = 47.7
    myArray(21, 5) = 100.7
    myArray(21, 6) = 23.7
    myArray(21, 7) = 83.8
    myArray(21, 8) = 148.3
    myArray(21, 9) = 89.4
    myArray(21, 10) = 24.2
    myArray(21, 11) = 79.1
    myArray(21, 12) = 127.6
    myArray(21, 13) = 39.8
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 18.9
    myArray(22, 3) = 34.2
    myArray(22, 4) = 57.5
    myArray(22, 5) = 155.8
    myArray(22, 6) = 60.1
    myArray(22, 7) = 45.5
    myArray(22, 8) = 304.4
    myArray(22, 9) = 90.6
    myArray(22, 10) = 187.2
    myArray(22, 11) = 123.8
    myArray(22, 12) = 30.2
    myArray(22, 13) = 38.9
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 20.4
    myArray(23, 3) = 45
    myArray(23, 4) = 30.9
    myArray(23, 5) = 64.7
    myArray(23, 6) = 21.3
    myArray(23, 7) = 60.6
    myArray(23, 8) = 273.4
    myArray(23, 9) = 208.2
    myArray(23, 10) = 105.9
    myArray(23, 11) = 48.2
    myArray(23, 12) = 17.2
    myArray(23, 13) = 26
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 29.2
    myArray(24, 3) = 44.3
    myArray(24, 4) = 127.7
    myArray(24, 5) = 132
    myArray(24, 6) = 81.1
    myArray(24, 7) = 89.6
    myArray(24, 8) = 123.8
    myArray(24, 9) = 335.4
    myArray(24, 10) = 92.3
    myArray(24, 11) = 196.4
    myArray(24, 12) = 24.4
    myArray(24, 13) = 28
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 7.7
    myArray(25, 3) = 31.1
    myArray(25, 4) = 33.5
    myArray(25, 5) = 91.7
    myArray(25, 6) = 46.3
    myArray(25, 7) = 116.5
    myArray(25, 8) = 185.4
    myArray(25, 9) = 218.1
    myArray(25, 10) = 191.8
    myArray(25, 11) = 179.4
    myArray(25, 12) = 35.4
    myArray(25, 13) = 27.2
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 71
    myArray(26, 3) = 68.2
    myArray(26, 4) = 21
    myArray(26, 5) = 39.7
    myArray(26, 6) = 65.3
    myArray(26, 7) = 182.8
    myArray(26, 8) = 499.3
    myArray(26, 9) = 333.5
    myArray(26, 10) = 226.8
    myArray(26, 11) = 5
    myArray(26, 12) = 38.3
    myArray(26, 13) = 3.1
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 18.2
    myArray(27, 3) = 14.8
    myArray(27, 4) = 83.8
    myArray(27, 5) = 43.8
    myArray(27, 6) = 165.7
    myArray(27, 7) = 50.8
    myArray(27, 8) = 217.1
    myArray(27, 9) = 285.6
    myArray(27, 10) = 161.2
    myArray(27, 11) = 41
    myArray(27, 12) = 43.9
    myArray(27, 13) = 3.2
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 3.4
    myArray(28, 3) = 2.3
    myArray(28, 4) = 71.2
    myArray(28, 5) = 57
    myArray(28, 6) = 4.7
    myArray(28, 7) = 98.5
    myArray(28, 8) = 189.1
    myArray(28, 9) = 272.6
    myArray(28, 10) = 117.1
    myArray(28, 11) = 62.5
    myArray(28, 12) = 38.7
    myArray(28, 13) = 11.5
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 23.4
    myArray(29, 3) = 8.2
    myArray(29, 4) = 32.3
    myArray(29, 5) = 43.9
    myArray(29, 6) = 144.8
    myArray(29, 7) = 254.9
    myArray(29, 8) = 390.3
    myArray(29, 9) = 288
    myArray(29, 10) = 199.8
    myArray(29, 11) = 12.2
    myArray(29, 12) = 36.6
    myArray(29, 13) = 117.1
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 42.5
    myArray(30, 3) = 100
    myArray(30, 4) = 65.6
    myArray(30, 5) = 49.3
    myArray(30, 6) = 78.5
    myArray(30, 7) = 95.4
    myArray(30, 8) = 443.6
    myArray(30, 9) = 20.6
    myArray(30, 10) = 221.3
    myArray(30, 11) = 74.9
    myArray(30, 12) = 31.2
    myArray(30, 13) = 8.6

    data_CHUPUNGNYEONG = myArray

End Function

Function data_CHEONGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 21.5
    myArray(1, 3) = 14
    myArray(1, 4) = 34.4
    myArray(1, 5) = 64
    myArray(1, 6) = 70.7
    myArray(1, 7) = 30.9
    myArray(1, 8) = 204.9
    myArray(1, 9) = 835.4
    myArray(1, 10) = 17.5
    myArray(1, 11) = 22.6
    myArray(1, 12) = 20.3
    myArray(1, 13) = 3.6
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 27.9
    myArray(2, 3) = 4.2
    myArray(2, 4) = 98.4
    myArray(2, 5) = 28.6
    myArray(2, 6) = 36.8
    myArray(2, 7) = 255.8
    myArray(2, 8) = 170.5
    myArray(2, 9) = 128.6
    myArray(2, 10) = 11.2
    myArray(2, 11) = 67.1
    myArray(2, 12) = 77.2
    myArray(2, 13) = 22.5
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 12.9
    myArray(3, 3) = 39.1
    myArray(3, 4) = 31.6
    myArray(3, 5) = 58.5
    myArray(3, 6) = 179.1
    myArray(3, 7) = 210.3
    myArray(3, 8) = 425.5
    myArray(3, 9) = 211.1
    myArray(3, 10) = 55.5
    myArray(3, 11) = 8.4
    myArray(3, 12) = 180.3
    myArray(3, 13) = 44.3
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 22
    myArray(4, 3) = 28.9
    myArray(4, 4) = 30.9
    myArray(4, 5) = 153.1
    myArray(4, 6) = 92.8
    myArray(4, 7) = 247
    myArray(4, 8) = 253
    myArray(4, 9) = 460.6
    myArray(4, 10) = 225.9
    myArray(4, 11) = 74.2
    myArray(4, 12) = 44.7
    myArray(4, 13) = 7.1
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 1.6
    myArray(5, 3) = 3.6
    myArray(5, 4) = 54.1
    myArray(5, 5) = 91.4
    myArray(5, 6) = 102.4
    myArray(5, 7) = 191.1
    myArray(5, 8) = 122.4
    myArray(5, 9) = 197.4
    myArray(5, 10) = 281.3
    myArray(5, 11) = 252.4
    myArray(5, 12) = 15.4
    myArray(5, 13) = 13.4
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 38.7
    myArray(6, 3) = 1.3
    myArray(6, 4) = 10.4
    myArray(6, 5) = 56.1
    myArray(6, 6) = 42.1
    myArray(6, 7) = 185.7
    myArray(6, 8) = 300
    myArray(6, 9) = 390.4
    myArray(6, 10) = 244.6
    myArray(6, 11) = 32.1
    myArray(6, 12) = 37.3
    myArray(6, 13) = 18.9
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 56.9
    myArray(7, 3) = 50.3
    myArray(7, 4) = 11.3
    myArray(7, 5) = 12.7
    myArray(7, 6) = 14.3
    myArray(7, 7) = 217.5
    myArray(7, 8) = 171.5
    myArray(7, 9) = 135.5
    myArray(7, 10) = 11.8
    myArray(7, 11) = 75.9
    myArray(7, 12) = 6.9
    myArray(7, 13) = 19.5
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 58.7
    myArray(8, 3) = 9
    myArray(8, 4) = 25.9
    myArray(8, 5) = 132
    myArray(8, 6) = 106.9
    myArray(8, 7) = 57.9
    myArray(8, 8) = 186.2
    myArray(8, 9) = 482.4
    myArray(8, 10) = 90.5
    myArray(8, 11) = 58
    myArray(8, 12) = 26.3
    myArray(8, 13) = 48
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 16.2
    myArray(9, 3) = 45
    myArray(9, 4) = 38.9
    myArray(9, 5) = 192.7
    myArray(9, 6) = 113.5
    myArray(9, 7) = 186
    myArray(9, 8) = 467.2
    myArray(9, 9) = 293.9
    myArray(9, 10) = 150.6
    myArray(9, 11) = 32.5
    myArray(9, 12) = 33.1
    myArray(9, 13) = 12.2
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 12.5
    myArray(10, 3) = 42.3
    myArray(10, 4) = 67.3
    myArray(10, 5) = 61
    myArray(10, 6) = 121.8
    myArray(10, 7) = 421.5
    myArray(10, 8) = 318.9
    myArray(10, 9) = 247.6
    myArray(10, 10) = 139
    myArray(10, 11) = 2
    myArray(10, 12) = 34
    myArray(10, 13) = 38
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 4.6
    myArray(11, 3) = 13.8
    myArray(11, 4) = 36.8
    myArray(11, 5) = 66.1
    myArray(11, 6) = 50.7
    myArray(11, 7) = 170
    myArray(11, 8) = 373.1
    myArray(11, 9) = 334.7
    myArray(11, 10) = 295.5
    myArray(11, 11) = 54.6
    myArray(11, 12) = 16
    myArray(11, 13) = 11.3
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 20
    myArray(12, 3) = 28.9
    myArray(12, 4) = 8.2
    myArray(12, 5) = 89.3
    myArray(12, 6) = 119.4
    myArray(12, 7) = 115.5
    myArray(12, 8) = 508
    myArray(12, 9) = 52
    myArray(12, 10) = 18.4
    myArray(12, 11) = 21.3
    myArray(12, 12) = 83.4
    myArray(12, 13) = 16.7
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 11.2
    myArray(13, 3) = 33.3
    myArray(13, 4) = 103.2
    myArray(13, 5) = 35.8
    myArray(13, 6) = 145.5
    myArray(13, 7) = 81.2
    myArray(13, 8) = 273.2
    myArray(13, 9) = 385.5
    myArray(13, 10) = 391.4
    myArray(13, 11) = 43.5
    myArray(13, 12) = 8.8
    myArray(13, 13) = 21.9
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 29
    myArray(14, 3) = 7.7
    myArray(14, 4) = 29.4
    myArray(14, 5) = 27
    myArray(14, 6) = 64.5
    myArray(14, 7) = 112
    myArray(14, 8) = 296.6
    myArray(14, 9) = 195.5
    myArray(14, 10) = 92.6
    myArray(14, 11) = 13.1
    myArray(14, 12) = 10.5
    myArray(14, 13) = 14.4
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 17.8
    myArray(15, 3) = 13.1
    myArray(15, 4) = 54.9
    myArray(15, 5) = 30.4
    myArray(15, 6) = 109.6
    myArray(15, 7) = 77.2
    myArray(15, 8) = 345.7
    myArray(15, 9) = 187.5
    myArray(15, 10) = 49.5
    myArray(15, 11) = 49.5
    myArray(15, 12) = 43.9
    myArray(15, 13) = 40.7
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 37.8
    myArray(16, 3) = 69.2
    myArray(16, 4) = 99.8
    myArray(16, 5) = 70.5
    myArray(16, 6) = 110
    myArray(16, 7) = 42.6
    myArray(16, 8) = 224.1
    myArray(16, 9) = 433.2
    myArray(16, 10) = 278.6
    myArray(16, 11) = 17.1
    myArray(16, 12) = 15.7
    myArray(16, 13) = 23.8
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 4.5
    myArray(17, 3) = 43.2
    myArray(17, 4) = 23.5
    myArray(17, 5) = 111.2
    myArray(17, 6) = 116.2
    myArray(17, 7) = 360.7
    myArray(17, 8) = 531.9
    myArray(17, 9) = 290.2
    myArray(17, 10) = 182.5
    myArray(17, 11) = 34.5
    myArray(17, 12) = 92.6
    myArray(17, 13) = 14.6
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 17.8
    myArray(18, 3) = 3.7
    myArray(18, 4) = 65.1
    myArray(18, 5) = 106.8
    myArray(18, 6) = 31.2
    myArray(18, 7) = 93.7
    myArray(18, 8) = 257.4
    myArray(18, 9) = 479.5
    myArray(18, 10) = 162.5
    myArray(18, 11) = 61.2
    myArray(18, 12) = 52.1
    myArray(18, 13) = 56.6
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 30.5
    myArray(19, 3) = 33.2
    myArray(19, 4) = 46.8
    myArray(19, 5) = 65
    myArray(19, 6) = 97.9
    myArray(19, 7) = 229.9
    myArray(19, 8) = 253.6
    myArray(19, 9) = 183.9
    myArray(19, 10) = 162.6
    myArray(19, 11) = 25
    myArray(19, 12) = 75
    myArray(19, 13) = 37.3
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 5.9
    myArray(20, 3) = 6.8
    myArray(20, 4) = 51.1
    myArray(20, 5) = 43.7
    myArray(20, 6) = 35
    myArray(20, 7) = 92.6
    myArray(20, 8) = 125.1
    myArray(20, 9) = 197.5
    myArray(20, 10) = 147.5
    myArray(20, 11) = 151.1
    myArray(20, 12) = 24.8
    myArray(20, 13) = 32.6
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 16
    myArray(21, 3) = 26.5
    myArray(21, 4) = 44.1
    myArray(21, 5) = 109.1
    myArray(21, 6) = 24.4
    myArray(21, 7) = 83.3
    myArray(21, 8) = 141.4
    myArray(21, 9) = 54.3
    myArray(21, 10) = 20.1
    myArray(21, 11) = 90.5
    myArray(21, 12) = 107.5
    myArray(21, 13) = 39.7
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 5.7
    myArray(22, 3) = 45.5
    myArray(22, 4) = 13.2
    myArray(22, 5) = 132.1
    myArray(22, 6) = 84.4
    myArray(22, 7) = 39.9
    myArray(22, 8) = 320
    myArray(22, 9) = 69
    myArray(22, 10) = 78.1
    myArray(22, 11) = 83.6
    myArray(22, 12) = 26.4
    myArray(22, 13) = 40.1
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 12
    myArray(23, 3) = 38.7
    myArray(23, 4) = 8.9
    myArray(23, 5) = 61.7
    myArray(23, 6) = 11.9
    myArray(23, 7) = 17.5
    myArray(23, 8) = 789.1
    myArray(23, 9) = 225.2
    myArray(23, 10) = 78.3
    myArray(23, 11) = 23.1
    myArray(23, 12) = 13.7
    myArray(23, 13) = 21.1
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 17.6
    myArray(24, 3) = 30.6
    myArray(24, 4) = 81.7
    myArray(24, 5) = 133
    myArray(24, 6) = 92
    myArray(24, 7) = 63.3
    myArray(24, 8) = 324.9
    myArray(24, 9) = 247.9
    myArray(24, 10) = 204
    myArray(24, 11) = 112.2
    myArray(24, 12) = 45.9
    myArray(24, 13) = 28.5
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.1
    myArray(25, 3) = 23
    myArray(25, 4) = 20.3
    myArray(25, 5) = 60.8
    myArray(25, 6) = 20.3
    myArray(25, 7) = 82.5
    myArray(25, 8) = 204.8
    myArray(25, 9) = 80.5
    myArray(25, 10) = 155.1
    myArray(25, 11) = 84.3
    myArray(25, 12) = 104.9
    myArray(25, 13) = 20.1
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 62
    myArray(26, 3) = 62.7
    myArray(26, 4) = 22.9
    myArray(26, 5) = 15.7
    myArray(26, 6) = 65.3
    myArray(26, 7) = 145.9
    myArray(26, 8) = 386.6
    myArray(26, 9) = 385.8
    myArray(26, 10) = 160.6
    myArray(26, 11) = 5.8
    myArray(26, 12) = 41
    myArray(26, 13) = 4.3
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 12.7
    myArray(27, 3) = 7.5
    myArray(27, 4) = 76.6
    myArray(27, 5) = 46.4
    myArray(27, 6) = 136.4
    myArray(27, 7) = 75.4
    myArray(27, 8) = 138.1
    myArray(27, 9) = 233.1
    myArray(27, 10) = 185
    myArray(27, 11) = 29.4
    myArray(27, 12) = 57.3
    myArray(27, 13) = 3.7
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 1.4
    myArray(28, 3) = 2.4
    myArray(28, 4) = 59
    myArray(28, 5) = 45.2
    myArray(28, 6) = 9.1
    myArray(28, 7) = 129.6
    myArray(28, 8) = 171.7
    myArray(28, 9) = 519.4
    myArray(28, 10) = 116
    myArray(28, 11) = 105.9
    myArray(28, 12) = 56.7
    myArray(28, 13) = 20
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 28
    myArray(29, 3) = 2.8
    myArray(29, 4) = 18.8
    myArray(29, 5) = 30.1
    myArray(29, 6) = 202.4
    myArray(29, 7) = 100.5
    myArray(29, 8) = 698.5
    myArray(29, 9) = 297.7
    myArray(29, 10) = 270.6
    myArray(29, 11) = 17.4
    myArray(29, 12) = 41.5
    myArray(29, 13) = 97.3
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 37.2
    myArray(30, 3) = 79.1
    myArray(30, 4) = 37.9
    myArray(30, 5) = 49.1
    myArray(30, 6) = 122.7
    myArray(30, 7) = 84.7
    myArray(30, 8) = 520.6
    myArray(30, 9) = 113.1
    myArray(30, 10) = 328
    myArray(30, 11) = 112.7
    myArray(30, 12) = 41.2
    myArray(30, 13) = 8.6

    data_CHEONGJU = myArray

End Function


Function data_CHEONAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 19
    myArray(1, 3) = 8.2
    myArray(1, 4) = 25.3
    myArray(1, 5) = 47
    myArray(1, 6) = 48
    myArray(1, 7) = 14.5
    myArray(1, 8) = 239.9
    myArray(1, 9) = 1082.5
    myArray(1, 10) = 29
    myArray(1, 11) = 23.5
    myArray(1, 12) = 40.2
    myArray(1, 13) = 8.9
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 29.5
    myArray(2, 3) = 10.2
    myArray(2, 4) = 115
    myArray(2, 5) = 54.5
    myArray(2, 6) = 19
    myArray(2, 7) = 237
    myArray(2, 8) = 177.5
    myArray(2, 9) = 116.5
    myArray(2, 10) = 8
    myArray(2, 11) = 102.5
    myArray(2, 12) = 71.6
    myArray(2, 13) = 26.2
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 10.7
    myArray(3, 3) = 44.1
    myArray(3, 4) = 30
    myArray(3, 5) = 66.5
    myArray(3, 6) = 211
    myArray(3, 7) = 191.5
    myArray(3, 8) = 305
    myArray(3, 9) = 175.5
    myArray(3, 10) = 14.5
    myArray(3, 11) = 23
    myArray(3, 12) = 153.5
    myArray(3, 13) = 43.5
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 20.4
    myArray(4, 3) = 27.9
    myArray(4, 4) = 29.5
    myArray(4, 5) = 120.5
    myArray(4, 6) = 85
    myArray(4, 7) = 219.5
    myArray(4, 8) = 277
    myArray(4, 9) = 408.1
    myArray(4, 10) = 283
    myArray(4, 11) = 51.5
    myArray(4, 12) = 52.8
    myArray(4, 13) = 8.5
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 2.7
    myArray(5, 3) = 2.8
    myArray(5, 4) = 46.5
    myArray(5, 5) = 88.5
    myArray(5, 6) = 121.5
    myArray(5, 7) = 163.7
    myArray(5, 8) = 138.5
    myArray(5, 9) = 313.5
    myArray(5, 10) = 324.5
    myArray(5, 11) = 134.5
    myArray(5, 12) = 16.5
    myArray(5, 13) = 11.9
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 52.3
    myArray(6, 3) = 2.7
    myArray(6, 4) = 7.1
    myArray(6, 5) = 36
    myArray(6, 6) = 36
    myArray(6, 7) = 181
    myArray(6, 8) = 83
    myArray(6, 9) = 636
    myArray(6, 10) = 287.5
    myArray(6, 11) = 32
    myArray(6, 12) = 32
    myArray(6, 13) = 22.5
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 43.5
    myArray(7, 3) = 44
    myArray(7, 4) = 16.5
    myArray(7, 5) = 19
    myArray(7, 6) = 15
    myArray(7, 7) = 227.5
    myArray(7, 8) = 178
    myArray(7, 9) = 194.5
    myArray(7, 10) = 12
    myArray(7, 11) = 63.5
    myArray(7, 12) = 6.3
    myArray(7, 13) = 18.4
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 45.3
    myArray(8, 3) = 6
    myArray(8, 4) = 25.5
    myArray(8, 5) = 128
    myArray(8, 6) = 104
    myArray(8, 7) = 54
    myArray(8, 8) = 229.5
    myArray(8, 9) = 481.5
    myArray(8, 10) = 57
    myArray(8, 11) = 91.5
    myArray(8, 12) = 42.1
    myArray(8, 13) = 48.1
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 18.6
    myArray(9, 3) = 44
    myArray(9, 4) = 38.1
    myArray(9, 5) = 172.3
    myArray(9, 6) = 106
    myArray(9, 7) = 178.6
    myArray(9, 8) = 381.2
    myArray(9, 9) = 334.6
    myArray(9, 10) = 264.2
    myArray(9, 11) = 27
    myArray(9, 12) = 46.7
    myArray(9, 13) = 17
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 16.4
    myArray(10, 3) = 21.3
    myArray(10, 4) = 21.5
    myArray(10, 5) = 67.5
    myArray(10, 6) = 127.6
    myArray(10, 7) = 235
    myArray(10, 8) = 365.2
    myArray(10, 9) = 229.3
    myArray(10, 10) = 189
    myArray(10, 11) = 4.5
    myArray(10, 12) = 53
    myArray(10, 13) = 33
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 3
    myArray(11, 3) = 29.8
    myArray(11, 4) = 37
    myArray(11, 5) = 53.7
    myArray(11, 6) = 48
    myArray(11, 7) = 183
    myArray(11, 8) = 313.8
    myArray(11, 9) = 202
    myArray(11, 10) = 377.5
    myArray(11, 11) = 26.7
    myArray(11, 12) = 23.5
    myArray(11, 13) = 11.3
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 25.2
    myArray(12, 3) = 18.5
    myArray(12, 4) = 6.1
    myArray(12, 5) = 78.6
    myArray(12, 6) = 79
    myArray(12, 7) = 120
    myArray(12, 8) = 535.1
    myArray(12, 9) = 63.5
    myArray(12, 10) = 22.2
    myArray(12, 11) = 21.6
    myArray(12, 12) = 56.3
    myArray(12, 13) = 17.2
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 9.4
    myArray(13, 3) = 34.1
    myArray(13, 4) = 108.3
    myArray(13, 5) = 35.3
    myArray(13, 6) = 126.2
    myArray(13, 7) = 106.7
    myArray(13, 8) = 215.6
    myArray(13, 9) = 470.6
    myArray(13, 10) = 353.3
    myArray(13, 11) = 43.4
    myArray(13, 12) = 15.6
    myArray(13, 13) = 43.9
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 17.5
    myArray(14, 3) = 11.1
    myArray(14, 4) = 40.1
    myArray(14, 5) = 34
    myArray(14, 6) = 62.6
    myArray(14, 7) = 126.7
    myArray(14, 8) = 287.2
    myArray(14, 9) = 138.8
    myArray(14, 10) = 89.3
    myArray(14, 11) = 30.4
    myArray(14, 12) = 16.6
    myArray(14, 13) = 15.8
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 13.3
    myArray(15, 3) = 16
    myArray(15, 4) = 51.6
    myArray(15, 5) = 30.6
    myArray(15, 6) = 112.6
    myArray(15, 7) = 55.6
    myArray(15, 8) = 335.8
    myArray(15, 9) = 212.3
    myArray(15, 10) = 30.8
    myArray(15, 11) = 61.1
    myArray(15, 12) = 39.7
    myArray(15, 13) = 40.5
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 40.7
    myArray(16, 3) = 50.4
    myArray(16, 4) = 73.8
    myArray(16, 5) = 61
    myArray(16, 6) = 84
    myArray(16, 7) = 37
    myArray(16, 8) = 171
    myArray(16, 9) = 486.1
    myArray(16, 10) = 316.9
    myArray(16, 11) = 19.4
    myArray(16, 12) = 13.5
    myArray(16, 13) = 24.5
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 7.9
    myArray(17, 3) = 31
    myArray(17, 4) = 26.5
    myArray(17, 5) = 133.2
    myArray(17, 6) = 103.3
    myArray(17, 7) = 374.6
    myArray(17, 8) = 645.1
    myArray(17, 9) = 268.2
    myArray(17, 10) = 153.2
    myArray(17, 11) = 26.5
    myArray(17, 12) = 65.8
    myArray(17, 13) = 10.5
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 14.5
    myArray(18, 3) = 2.3
    myArray(18, 4) = 44.9
    myArray(18, 5) = 81.6
    myArray(18, 6) = 16.8
    myArray(18, 7) = 75.1
    myArray(18, 8) = 252.5
    myArray(18, 9) = 483.7
    myArray(18, 10) = 190.1
    myArray(18, 11) = 66.6
    myArray(18, 12) = 52.6
    myArray(18, 13) = 56
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 28.5
    myArray(19, 3) = 35.2
    myArray(19, 4) = 40
    myArray(19, 5) = 56.3
    myArray(19, 6) = 123.5
    myArray(19, 7) = 102.1
    myArray(19, 8) = 308.2
    myArray(19, 9) = 173.6
    myArray(19, 10) = 117.5
    myArray(19, 11) = 12.2
    myArray(19, 12) = 58.2
    myArray(19, 13) = 40.3
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 4.9
    myArray(20, 3) = 14.7
    myArray(20, 4) = 40.9
    myArray(20, 5) = 62.1
    myArray(20, 6) = 34.6
    myArray(20, 7) = 73.9
    myArray(20, 8) = 239
    myArray(20, 9) = 218.7
    myArray(20, 10) = 144
    myArray(20, 11) = 119.5
    myArray(20, 12) = 28.9
    myArray(20, 13) = 38.9
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 12.7
    myArray(21, 3) = 21.5
    myArray(21, 4) = 23.3
    myArray(21, 5) = 87.6
    myArray(21, 6) = 27.5
    myArray(21, 7) = 86
    myArray(21, 8) = 136.8
    myArray(21, 9) = 64.2
    myArray(21, 10) = 29
    myArray(21, 11) = 69
    myArray(21, 12) = 128.6
    myArray(21, 13) = 41.8
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 8
    myArray(22, 3) = 43.6
    myArray(22, 4) = 16.5
    myArray(22, 5) = 118.3
    myArray(22, 6) = 107.2
    myArray(22, 7) = 36.2
    myArray(22, 8) = 364.3
    myArray(22, 9) = 82
    myArray(22, 10) = 55
    myArray(22, 11) = 95.9
    myArray(22, 12) = 33.5
    myArray(22, 13) = 44.3
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 13.9
    myArray(23, 3) = 32.2
    myArray(23, 4) = 6.5
    myArray(23, 5) = 42.9
    myArray(23, 6) = 14.3
    myArray(23, 7) = 15.6
    myArray(23, 8) = 788.1
    myArray(23, 9) = 291.5
    myArray(23, 10) = 43.3
    myArray(23, 11) = 14.1
    myArray(23, 12) = 23.8
    myArray(23, 13) = 18.8
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 14
    myArray(24, 3) = 31.6
    myArray(24, 4) = 62.2
    myArray(24, 5) = 117
    myArray(24, 6) = 82.7
    myArray(24, 7) = 88.9
    myArray(24, 8) = 185.8
    myArray(24, 9) = 282.7
    myArray(24, 10) = 124.6
    myArray(24, 11) = 99.8
    myArray(24, 12) = 48.3
    myArray(24, 13) = 25.8
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.6
    myArray(25, 3) = 25.5
    myArray(25, 4) = 26.9
    myArray(25, 5) = 43.9
    myArray(25, 6) = 15.1
    myArray(25, 7) = 84.9
    myArray(25, 8) = 234.7
    myArray(25, 9) = 90.7
    myArray(25, 10) = 102.8
    myArray(25, 11) = 81.9
    myArray(25, 12) = 120.6
    myArray(25, 13) = 18
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 59.7
    myArray(26, 3) = 63.1
    myArray(26, 4) = 21.7
    myArray(26, 5) = 15.1
    myArray(26, 6) = 86.4
    myArray(26, 7) = 121.9
    myArray(26, 8) = 356.4
    myArray(26, 9) = 481.7
    myArray(26, 10) = 167.2
    myArray(26, 11) = 18.9
    myArray(26, 12) = 45.9
    myArray(26, 13) = 5.5
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 17.8
    myArray(27, 3) = 9.2
    myArray(27, 4) = 75.3
    myArray(27, 5) = 54.7
    myArray(27, 6) = 135.9
    myArray(27, 7) = 44.8
    myArray(27, 8) = 117.6
    myArray(27, 9) = 230
    myArray(27, 10) = 250.8
    myArray(27, 11) = 49.5
    myArray(27, 12) = 67.9
    myArray(27, 13) = 5.4
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 3.3
    myArray(28, 3) = 3.3
    myArray(28, 4) = 57.6
    myArray(28, 5) = 51.6
    myArray(28, 6) = 9.8
    myArray(28, 7) = 168
    myArray(28, 8) = 117
    myArray(28, 9) = 366.6
    myArray(28, 10) = 133.3
    myArray(28, 11) = 98.2
    myArray(28, 12) = 43.2
    myArray(28, 13) = 28.8
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 31
    myArray(29, 3) = 3.1
    myArray(29, 4) = 16.4
    myArray(29, 5) = 29.6
    myArray(29, 6) = 116.9
    myArray(29, 7) = 178.9
    myArray(29, 8) = 574.9
    myArray(29, 9) = 196.5
    myArray(29, 10) = 180.1
    myArray(29, 11) = 28.7
    myArray(29, 12) = 56.9
    myArray(29, 13) = 89.5
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 34.9
    myArray(30, 3) = 89.7
    myArray(30, 4) = 41.3
    myArray(30, 5) = 48.2
    myArray(30, 6) = 107.7
    myArray(30, 7) = 106.4
    myArray(30, 8) = 509.9
    myArray(30, 9) = 66.4
    myArray(30, 10) = 318
    myArray(30, 11) = 99.7
    myArray(30, 12) = 39.5
    myArray(30, 13) = 15.2
    

    data_CHEONAN = myArray

End Function


Function data_JAECHEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 14
    myArray(1, 3) = 6.1
    myArray(1, 4) = 43.3
    myArray(1, 5) = 60
    myArray(1, 6) = 64.5
    myArray(1, 7) = 72.5
    myArray(1, 8) = 292.5
    myArray(1, 9) = 742.5
    myArray(1, 10) = 66
    myArray(1, 11) = 38.5
    myArray(1, 12) = 42
    myArray(1, 13) = 6
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 33.2
    myArray(2, 3) = 8.8
    myArray(2, 4) = 111.5
    myArray(2, 5) = 46.5
    myArray(2, 6) = 33.5
    myArray(2, 7) = 174
    myArray(2, 8) = 264
    myArray(2, 9) = 121.5
    myArray(2, 10) = 23
    myArray(2, 11) = 90
    myArray(2, 12) = 62.4
    myArray(2, 13) = 18.8
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 9.3
    myArray(3, 3) = 48.4
    myArray(3, 4) = 29.5
    myArray(3, 5) = 65
    myArray(3, 6) = 203
    myArray(3, 7) = 151
    myArray(3, 8) = 423.5
    myArray(3, 9) = 197.3
    myArray(3, 10) = 85.5
    myArray(3, 11) = 16.2
    myArray(3, 12) = 123.2
    myArray(3, 13) = 32.9
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 15.4
    myArray(4, 3) = 27.1
    myArray(4, 4) = 27
    myArray(4, 5) = 136.5
    myArray(4, 6) = 101
    myArray(4, 7) = 195.8
    myArray(4, 8) = 289.5
    myArray(4, 9) = 546.1
    myArray(4, 10) = 132.2
    myArray(4, 11) = 72.5
    myArray(4, 12) = 34.7
    myArray(4, 13) = 3.6
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 4.9
    myArray(5, 3) = 4.2
    myArray(5, 4) = 66.1
    myArray(5, 5) = 129.5
    myArray(5, 6) = 104
    myArray(5, 7) = 136.5
    myArray(5, 8) = 214
    myArray(5, 9) = 323
    myArray(5, 10) = 281.5
    myArray(5, 11) = 159
    myArray(5, 12) = 23.5
    myArray(5, 13) = 7.2
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 39.9
    myArray(6, 3) = 4.1
    myArray(6, 4) = 15.5
    myArray(6, 5) = 41.5
    myArray(6, 6) = 96
    myArray(6, 7) = 196.5
    myArray(6, 8) = 197.2
    myArray(6, 9) = 259
    myArray(6, 10) = 221.5
    myArray(6, 11) = 24
    myArray(6, 12) = 34
    myArray(6, 13) = 19.9
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 46.4
    myArray(7, 3) = 49.6
    myArray(7, 4) = 18.8
    myArray(7, 5) = 18.5
    myArray(7, 6) = 9
    myArray(7, 7) = 269.5
    myArray(7, 8) = 227.5
    myArray(7, 9) = 93.5
    myArray(7, 10) = 19.5
    myArray(7, 11) = 84
    myArray(7, 12) = 3
    myArray(7, 13) = 10
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 42.3
    myArray(8, 3) = 3.2
    myArray(8, 4) = 24
    myArray(8, 5) = 199
    myArray(8, 6) = 92
    myArray(8, 7) = 89
    myArray(8, 8) = 214.5
    myArray(8, 9) = 652
    myArray(8, 10) = 69.5
    myArray(8, 11) = 44.5
    myArray(8, 12) = 13
    myArray(8, 13) = 57.4
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 11.2
    myArray(9, 3) = 48.7
    myArray(9, 4) = 42
    myArray(9, 5) = 198
    myArray(9, 6) = 149.5
    myArray(9, 7) = 196.5
    myArray(9, 8) = 495
    myArray(9, 9) = 346
    myArray(9, 10) = 287.5
    myArray(9, 11) = 24.5
    myArray(9, 12) = 60.5
    myArray(9, 13) = 17.2
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 15.2
    myArray(10, 3) = 31.5
    myArray(10, 4) = 34.5
    myArray(10, 5) = 74
    myArray(10, 6) = 142
    myArray(10, 7) = 395.5
    myArray(10, 8) = 455.5
    myArray(10, 9) = 259
    myArray(10, 10) = 163.5
    myArray(10, 11) = 2
    myArray(10, 12) = 37
    myArray(10, 13) = 21.1
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 3.8
    myArray(11, 3) = 17.9
    myArray(11, 4) = 44
    myArray(11, 5) = 77.5
    myArray(11, 6) = 78.5
    myArray(11, 7) = 171.5
    myArray(11, 8) = 424.5
    myArray(11, 9) = 259
    myArray(11, 10) = 363
    myArray(11, 11) = 57.5
    myArray(11, 12) = 19.5
    myArray(11, 13) = 8.5
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 26.5
    myArray(12, 3) = 33.3
    myArray(12, 4) = 12.1
    myArray(12, 5) = 106
    myArray(12, 6) = 105.5
    myArray(12, 7) = 145
    myArray(12, 8) = 1111
    myArray(12, 9) = 56.5
    myArray(12, 10) = 25
    myArray(12, 11) = 39.5
    myArray(12, 12) = 46.3
    myArray(12, 13) = 13.1
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 11
    myArray(13, 3) = 34
    myArray(13, 4) = 177.2
    myArray(13, 5) = 19.5
    myArray(13, 6) = 151.5
    myArray(13, 7) = 124
    myArray(13, 8) = 442.5
    myArray(13, 9) = 696.5
    myArray(13, 10) = 333
    myArray(13, 11) = 33.5
    myArray(13, 12) = 24.3
    myArray(13, 13) = 20.3
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 19.8
    myArray(14, 3) = 6.5
    myArray(14, 4) = 63.8
    myArray(14, 5) = 41.5
    myArray(14, 6) = 54
    myArray(14, 7) = 86.5
    myArray(14, 8) = 274.6
    myArray(14, 9) = 222
    myArray(14, 10) = 63.6
    myArray(14, 11) = 24
    myArray(14, 12) = 8.4
    myArray(14, 13) = 21.1
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 14.8
    myArray(15, 3) = 27.5
    myArray(15, 4) = 57.1
    myArray(15, 5) = 40.4
    myArray(15, 6) = 121.7
    myArray(15, 7) = 162.1
    myArray(15, 8) = 474.5
    myArray(15, 9) = 213
    myArray(15, 10) = 58.5
    myArray(15, 11) = 31.5
    myArray(15, 12) = 44.3
    myArray(15, 13) = 32
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 61.1
    myArray(16, 3) = 68.6
    myArray(16, 4) = 117.3
    myArray(16, 5) = 71.5
    myArray(16, 6) = 118.1
    myArray(16, 7) = 87.6
    myArray(16, 8) = 180.1
    myArray(16, 9) = 345.7
    myArray(16, 10) = 432.8
    myArray(16, 11) = 22.8
    myArray(16, 12) = 22.4
    myArray(16, 13) = 17.2
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 2.8
    myArray(17, 3) = 52.9
    myArray(17, 4) = 36.5
    myArray(17, 5) = 189.8
    myArray(17, 6) = 121.8
    myArray(17, 7) = 459
    myArray(17, 8) = 665.2
    myArray(17, 9) = 398.5
    myArray(17, 10) = 157.2
    myArray(17, 11) = 55.4
    myArray(17, 12) = 81.1
    myArray(17, 13) = 10.3
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 16
    myArray(18, 3) = 5
    myArray(18, 4) = 83.2
    myArray(18, 5) = 135.5
    myArray(18, 6) = 40.9
    myArray(18, 7) = 108.1
    myArray(18, 8) = 344.8
    myArray(18, 9) = 319.9
    myArray(18, 10) = 144.5
    myArray(18, 11) = 67.1
    myArray(18, 12) = 68.7
    myArray(18, 13) = 47.6
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 40.5
    myArray(19, 3) = 55
    myArray(19, 4) = 48
    myArray(19, 5) = 92.3
    myArray(19, 6) = 118.5
    myArray(19, 7) = 144.6
    myArray(19, 8) = 442.4
    myArray(19, 9) = 274.3
    myArray(19, 10) = 118.9
    myArray(19, 11) = 8.9
    myArray(19, 12) = 63.2
    myArray(19, 13) = 30.5
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 10.5
    myArray(20, 3) = 23.6
    myArray(20, 4) = 44.5
    myArray(20, 5) = 49.5
    myArray(20, 6) = 41.4
    myArray(20, 7) = 62.1
    myArray(20, 8) = 111.4
    myArray(20, 9) = 246.2
    myArray(20, 10) = 131
    myArray(20, 11) = 150.6
    myArray(20, 12) = 24
    myArray(20, 13) = 18.8
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 17.5
    myArray(21, 3) = 32.2
    myArray(21, 4) = 31.7
    myArray(21, 5) = 83.5
    myArray(21, 6) = 31.5
    myArray(21, 7) = 75.4
    myArray(21, 8) = 225.1
    myArray(21, 9) = 63.8
    myArray(21, 10) = 36.6
    myArray(21, 11) = 68.1
    myArray(21, 12) = 110.6
    myArray(21, 13) = 27.4
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 4.5
    myArray(22, 3) = 68.7
    myArray(22, 4) = 21.5
    myArray(22, 5) = 117.1
    myArray(22, 6) = 82.4
    myArray(22, 7) = 42.1
    myArray(22, 8) = 419.7
    myArray(22, 9) = 113.8
    myArray(22, 10) = 47.3
    myArray(22, 11) = 109.6
    myArray(22, 12) = 22.1
    myArray(22, 13) = 59.1
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 10.2
    myArray(23, 3) = 29.5
    myArray(23, 4) = 24.3
    myArray(23, 5) = 70.8
    myArray(23, 6) = 12.5
    myArray(23, 7) = 69.6
    myArray(23, 8) = 464.8
    myArray(23, 9) = 265.1
    myArray(23, 10) = 43.3
    myArray(23, 11) = 22.5
    myArray(23, 12) = 34.1
    myArray(23, 13) = 24
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 6.9
    myArray(24, 3) = 23.8
    myArray(24, 4) = 61.7
    myArray(24, 5) = 112.9
    myArray(24, 6) = 172.5
    myArray(24, 7) = 138.5
    myArray(24, 8) = 161.5
    myArray(24, 9) = 350.3
    myArray(24, 10) = 185.5
    myArray(24, 11) = 105.4
    myArray(24, 12) = 60.3
    myArray(24, 13) = 30
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 1
    myArray(25, 3) = 28.5
    myArray(25, 4) = 49.5
    myArray(25, 5) = 58.1
    myArray(25, 6) = 26.1
    myArray(25, 7) = 90
    myArray(25, 8) = 158.6
    myArray(25, 9) = 99.1
    myArray(25, 10) = 164.5
    myArray(25, 11) = 66.9
    myArray(25, 12) = 81
    myArray(25, 13) = 19.7
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 68.1
    myArray(26, 3) = 57.4
    myArray(26, 4) = 14.4
    myArray(26, 5) = 28.9
    myArray(26, 6) = 130.1
    myArray(26, 7) = 78.6
    myArray(26, 8) = 317.5
    myArray(26, 9) = 685.8
    myArray(26, 10) = 134.1
    myArray(26, 11) = 12.4
    myArray(26, 12) = 13.4
    myArray(26, 13) = 11.1
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 18.2
    myArray(27, 3) = 12.2
    myArray(27, 4) = 100.1
    myArray(27, 5) = 104
    myArray(27, 6) = 157.8
    myArray(27, 7) = 81.6
    myArray(27, 8) = 223.1
    myArray(27, 9) = 174.9
    myArray(27, 10) = 209.2
    myArray(27, 11) = 26.6
    myArray(27, 12) = 48
    myArray(27, 13) = 6.5
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 3
    myArray(28, 3) = 6.8
    myArray(28, 4) = 88.3
    myArray(28, 5) = 40.5
    myArray(28, 6) = 7.5
    myArray(28, 7) = 273.9
    myArray(28, 8) = 310.9
    myArray(28, 9) = 485.3
    myArray(28, 10) = 93.7
    myArray(28, 11) = 87.1
    myArray(28, 12) = 60
    myArray(28, 13) = 19.6
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 32.6
    myArray(29, 3) = 2.2
    myArray(29, 4) = 14
    myArray(29, 5) = 48.6
    myArray(29, 6) = 167.1
    myArray(29, 7) = 231.3
    myArray(29, 8) = 605.3
    myArray(29, 9) = 240.9
    myArray(29, 10) = 234.6
    myArray(29, 11) = 33.8
    myArray(29, 12) = 49.8
    myArray(29, 13) = 105.5
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 24.8
    myArray(30, 3) = 78.9
    myArray(30, 4) = 52.2
    myArray(30, 5) = 54.7
    myArray(30, 6) = 128.7
    myArray(30, 7) = 171.7
    myArray(30, 8) = 467.8
    myArray(30, 9) = 99.1
    myArray(30, 10) = 195.6
    myArray(30, 11) = 114.1
    myArray(30, 12) = 42.1
    myArray(30, 13) = 3.1

    data_JAECHEON = myArray

End Function


Function data_SEOSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 22.7
    myArray(1, 3) = 7.2
    myArray(1, 4) = 37.3
    myArray(1, 5) = 48.2
    myArray(1, 6) = 67.1
    myArray(1, 7) = 24.5
    myArray(1, 8) = 144.1
    myArray(1, 9) = 992.7
    myArray(1, 10) = 20.2
    myArray(1, 11) = 19.3
    myArray(1, 12) = 49.9
    myArray(1, 13) = 15.1
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 29.1
    myArray(2, 3) = 5.7
    myArray(2, 4) = 115.1
    myArray(2, 5) = 48.1
    myArray(2, 6) = 20
    myArray(2, 7) = 179.2
    myArray(2, 8) = 152.8
    myArray(2, 9) = 74.1
    myArray(2, 10) = 6.4
    myArray(2, 11) = 92.2
    myArray(2, 12) = 72.1
    myArray(2, 13) = 35.3
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 20.5
    myArray(3, 3) = 32.5
    myArray(3, 4) = 29.6
    myArray(3, 5) = 69.5
    myArray(3, 6) = 232.8
    myArray(3, 7) = 204.4
    myArray(3, 8) = 298.7
    myArray(3, 9) = 87.2
    myArray(3, 10) = 16.1
    myArray(3, 11) = 8.7
    myArray(3, 12) = 116.7
    myArray(3, 13) = 40.2
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 40.1
    myArray(4, 3) = 54.2
    myArray(4, 4) = 35
    myArray(4, 5) = 160.6
    myArray(4, 6) = 95.5
    myArray(4, 7) = 281.7
    myArray(4, 8) = 295.6
    myArray(4, 9) = 491.8
    myArray(4, 10) = 168
    myArray(4, 11) = 24.3
    myArray(4, 12) = 55.6
    myArray(4, 13) = 9.2
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 8
    myArray(5, 3) = 7.8
    myArray(5, 4) = 59.9
    myArray(5, 5) = 90.1
    myArray(5, 6) = 178.8
    myArray(5, 7) = 105.1
    myArray(5, 8) = 175.6
    myArray(5, 9) = 497.4
    myArray(5, 10) = 532.6
    myArray(5, 11) = 111.3
    myArray(5, 12) = 36.6
    myArray(5, 13) = 23.4
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 63
    myArray(6, 3) = 2.9
    myArray(6, 4) = 3.7
    myArray(6, 5) = 38.1
    myArray(6, 6) = 62.1
    myArray(6, 7) = 204.4
    myArray(6, 8) = 60.8
    myArray(6, 9) = 608.1
    myArray(6, 10) = 298.1
    myArray(6, 11) = 34.4
    myArray(6, 12) = 24.8
    myArray(6, 13) = 24.4
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 66.9
    myArray(7, 3) = 40.4
    myArray(7, 4) = 12.7
    myArray(7, 5) = 18.7
    myArray(7, 6) = 17.8
    myArray(7, 7) = 200.2
    myArray(7, 8) = 402
    myArray(7, 9) = 136.6
    myArray(7, 10) = 15
    myArray(7, 11) = 47.5
    myArray(7, 12) = 8.2
    myArray(7, 13) = 20.8
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 22.5
    myArray(8, 3) = 7
    myArray(8, 4) = 29.3
    myArray(8, 5) = 179.5
    myArray(8, 6) = 177.3
    myArray(8, 7) = 60.8
    myArray(8, 8) = 296.1
    myArray(8, 9) = 428.2
    myArray(8, 10) = 57.5
    myArray(8, 11) = 78.3
    myArray(8, 12) = 36.3
    myArray(8, 13) = 14.8
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 20.9
    myArray(9, 3) = 39
    myArray(9, 4) = 22.5
    myArray(9, 5) = 180
    myArray(9, 6) = 105.5
    myArray(9, 7) = 221.8
    myArray(9, 8) = 290.2
    myArray(9, 9) = 257.9
    myArray(9, 10) = 201.9
    myArray(9, 11) = 23
    myArray(9, 12) = 53.6
    myArray(9, 13) = 17.1
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 27.3
    myArray(10, 3) = 26.3
    myArray(10, 4) = 15.7
    myArray(10, 5) = 80.2
    myArray(10, 6) = 140.3
    myArray(10, 7) = 211.1
    myArray(10, 8) = 321.9
    myArray(10, 9) = 131.2
    myArray(10, 10) = 282.6
    myArray(10, 11) = 1.8
    myArray(10, 12) = 70.5
    myArray(10, 13) = 32
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 10.4
    myArray(11, 3) = 34
    myArray(11, 4) = 36.1
    myArray(11, 5) = 77.2
    myArray(11, 6) = 56.1
    myArray(11, 7) = 147
    myArray(11, 8) = 386.1
    myArray(11, 9) = 270.5
    myArray(11, 10) = 228.7
    myArray(11, 11) = 30.9
    myArray(11, 12) = 19.6
    myArray(11, 13) = 37.6
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 29.7
    myArray(12, 3) = 18.9
    myArray(12, 4) = 5
    myArray(12, 5) = 77.3
    myArray(12, 6) = 133.5
    myArray(12, 7) = 226.8
    myArray(12, 8) = 494.5
    myArray(12, 9) = 58.2
    myArray(12, 10) = 10.1
    myArray(12, 11) = 10.5
    myArray(12, 12) = 55
    myArray(12, 13) = 19.7
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 13
    myArray(13, 3) = 25.5
    myArray(13, 4) = 127.2
    myArray(13, 5) = 28.1
    myArray(13, 6) = 108.5
    myArray(13, 7) = 123.5
    myArray(13, 8) = 257
    myArray(13, 9) = 414.6
    myArray(13, 10) = 305.8
    myArray(13, 11) = 30.7
    myArray(13, 12) = 14.4
    myArray(13, 13) = 22.8
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 15
    myArray(14, 3) = 7
    myArray(14, 4) = 26
    myArray(14, 5) = 46.1
    myArray(14, 6) = 88.5
    myArray(14, 7) = 118.1
    myArray(14, 8) = 335.5
    myArray(14, 9) = 114.2
    myArray(14, 10) = 62.7
    myArray(14, 11) = 34
    myArray(14, 12) = 34.6
    myArray(14, 13) = 27.9
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 15.2
    myArray(15, 3) = 26.5
    myArray(15, 4) = 67
    myArray(15, 5) = 43
    myArray(15, 6) = 117.9
    myArray(15, 7) = 74.9
    myArray(15, 8) = 364.9
    myArray(15, 9) = 196.3
    myArray(15, 10) = 16
    myArray(15, 11) = 49.2
    myArray(15, 12) = 59.1
    myArray(15, 13) = 44.3
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 55.5
    myArray(16, 3) = 58.4
    myArray(16, 4) = 79.2
    myArray(16, 5) = 52.2
    myArray(16, 6) = 168
    myArray(16, 7) = 94.9
    myArray(16, 8) = 447.1
    myArray(16, 9) = 707
    myArray(16, 10) = 402
    myArray(16, 11) = 29.1
    myArray(16, 12) = 12
    myArray(16, 13) = 36.4
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 8.8
    myArray(17, 3) = 55.8
    myArray(17, 4) = 34.5
    myArray(17, 5) = 96.2
    myArray(17, 6) = 107.9
    myArray(17, 7) = 462.6
    myArray(17, 8) = 656.5
    myArray(17, 9) = 151.2
    myArray(17, 10) = 50.3
    myArray(17, 11) = 18.1
    myArray(17, 12) = 48.9
    myArray(17, 13) = 13.6
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 15.1
    myArray(18, 3) = 2.4
    myArray(18, 4) = 41.6
    myArray(18, 5) = 113.5
    myArray(18, 6) = 14.5
    myArray(18, 7) = 91.1
    myArray(18, 8) = 266.8
    myArray(18, 9) = 647.9
    myArray(18, 10) = 201.5
    myArray(18, 11) = 100.7
    myArray(18, 12) = 82.1
    myArray(18, 13) = 65.4
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 36.8
    myArray(19, 3) = 64.8
    myArray(19, 4) = 60.8
    myArray(19, 5) = 61.8
    myArray(19, 6) = 114.9
    myArray(19, 7) = 94.4
    myArray(19, 8) = 213.8
    myArray(19, 9) = 120.6
    myArray(19, 10) = 147.4
    myArray(19, 11) = 5.7
    myArray(19, 12) = 64.9
    myArray(19, 13) = 32.8
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 7
    myArray(20, 3) = 17
    myArray(20, 4) = 31.2
    myArray(20, 5) = 85.6
    myArray(20, 6) = 52.7
    myArray(20, 7) = 69.3
    myArray(20, 8) = 151.7
    myArray(20, 9) = 242.3
    myArray(20, 10) = 106.7
    myArray(20, 11) = 117.2
    myArray(20, 12) = 37.8
    myArray(20, 13) = 81.6
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 20.7
    myArray(21, 3) = 23.1
    myArray(21, 4) = 20.6
    myArray(21, 5) = 116.8
    myArray(21, 6) = 40.6
    myArray(21, 7) = 64.1
    myArray(21, 8) = 158.5
    myArray(21, 9) = 63.1
    myArray(21, 10) = 15.1
    myArray(21, 11) = 73.1
    myArray(21, 12) = 156.6
    myArray(21, 13) = 63.6
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 21.9
    myArray(22, 3) = 61.7
    myArray(22, 4) = 24.3
    myArray(22, 5) = 87
    myArray(22, 6) = 153.7
    myArray(22, 7) = 36.8
    myArray(22, 8) = 295.6
    myArray(22, 9) = 34
    myArray(22, 10) = 53.1
    myArray(22, 11) = 73.8
    myArray(22, 12) = 17.5
    myArray(22, 13) = 62.7
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 21.3
    myArray(23, 3) = 31.4
    myArray(23, 4) = 4.8
    myArray(23, 5) = 38.9
    myArray(23, 6) = 27.9
    myArray(23, 7) = 23.3
    myArray(23, 8) = 327.8
    myArray(23, 9) = 231.3
    myArray(23, 10) = 37.6
    myArray(23, 11) = 25.5
    myArray(23, 12) = 24.7
    myArray(23, 13) = 35.9
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 21
    myArray(24, 3) = 40.5
    myArray(24, 4) = 76.6
    myArray(24, 5) = 132.8
    myArray(24, 6) = 147.7
    myArray(24, 7) = 162.3
    myArray(24, 8) = 152.9
    myArray(24, 9) = 156.8
    myArray(24, 10) = 82.7
    myArray(24, 11) = 153.2
    myArray(24, 12) = 73.9
    myArray(24, 13) = 26.8
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 1.1
    myArray(25, 3) = 30.2
    myArray(25, 4) = 43.7
    myArray(25, 5) = 43.9
    myArray(25, 6) = 20.1
    myArray(25, 7) = 56.3
    myArray(25, 8) = 174.5
    myArray(25, 9) = 121.1
    myArray(25, 10) = 181.1
    myArray(25, 11) = 81
    myArray(25, 12) = 124.6
    myArray(25, 13) = 37.4
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 46
    myArray(26, 3) = 72.3
    myArray(26, 4) = 23
    myArray(26, 5) = 20.7
    myArray(26, 6) = 101.3
    myArray(26, 7) = 144
    myArray(26, 8) = 329.4
    myArray(26, 9) = 400
    myArray(26, 10) = 257.7
    myArray(26, 11) = 12.6
    myArray(26, 12) = 72
    myArray(26, 13) = 9.7
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 25.3
    myArray(27, 3) = 9.6
    myArray(27, 4) = 112.8
    myArray(27, 5) = 110.6
    myArray(27, 6) = 132.3
    myArray(27, 7) = 70.9
    myArray(27, 8) = 121.6
    myArray(27, 9) = 217.8
    myArray(27, 10) = 206
    myArray(27, 11) = 55.9
    myArray(27, 12) = 126.2
    myArray(27, 13) = 18.3
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 8.6
    myArray(28, 3) = 4.7
    myArray(28, 4) = 72.1
    myArray(28, 5) = 52.2
    myArray(28, 6) = 2.9
    myArray(28, 7) = 352.4
    myArray(28, 8) = 178.4
    myArray(28, 9) = 468.7
    myArray(28, 10) = 165.9
    myArray(28, 11) = 160
    myArray(28, 12) = 72.9
    myArray(28, 13) = 31.9
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 30.5
    myArray(29, 3) = 0.1
    myArray(29, 4) = 6.4
    myArray(29, 5) = 54.6
    myArray(29, 6) = 132.9
    myArray(29, 7) = 138.1
    myArray(29, 8) = 507
    myArray(29, 9) = 225
    myArray(29, 10) = 166.1
    myArray(29, 11) = 39.6
    myArray(29, 12) = 122.9
    myArray(29, 13) = 106.5
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 34.4
    myArray(30, 3) = 86.4
    myArray(30, 4) = 26.1
    myArray(30, 5) = 45
    myArray(30, 6) = 157.4
    myArray(30, 7) = 120.4
    myArray(30, 8) = 556.1
    myArray(30, 9) = 197.6
    myArray(30, 10) = 354.2
    myArray(30, 11) = 150.3
    myArray(30, 12) = 62.1
    myArray(30, 13) = 16.1
    
    data_SEOSAN = myArray

End Function

Function data_BUYEO() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 22.6
    myArray(1, 3) = 23.5
    myArray(1, 4) = 24.4
    myArray(1, 5) = 62
    myArray(1, 6) = 59.5
    myArray(1, 7) = 34.5
    myArray(1, 8) = 171.5
    myArray(1, 9) = 839
    myArray(1, 10) = 46.5
    myArray(1, 11) = 22
    myArray(1, 12) = 15
    myArray(1, 13) = 5.7
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 26.4
    myArray(2, 3) = 2.8
    myArray(2, 4) = 131
    myArray(2, 5) = 45
    myArray(2, 6) = 33
    myArray(2, 7) = 289
    myArray(2, 8) = 235
    myArray(2, 9) = 67
    myArray(2, 10) = 16
    myArray(2, 11) = 90.5
    myArray(2, 12) = 76
    myArray(2, 13) = 35
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 9
    myArray(3, 3) = 54.9
    myArray(3, 4) = 44
    myArray(3, 5) = 70
    myArray(3, 6) = 229.5
    myArray(3, 7) = 236.5
    myArray(3, 8) = 404.5
    myArray(3, 9) = 263
    myArray(3, 10) = 24.5
    myArray(3, 11) = 8
    myArray(3, 12) = 219.5
    myArray(3, 13) = 39.5
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 40.6
    myArray(4, 3) = 47
    myArray(4, 4) = 45
    myArray(4, 5) = 200.5
    myArray(4, 6) = 130.5
    myArray(4, 7) = 324
    myArray(4, 8) = 323
    myArray(4, 9) = 451.3
    myArray(4, 10) = 313.1
    myArray(4, 11) = 75.5
    myArray(4, 12) = 46.3
    myArray(4, 13) = 3.5
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 3.5
    myArray(5, 3) = 10
    myArray(5, 4) = 75.7
    myArray(5, 5) = 92.5
    myArray(5, 6) = 127.5
    myArray(5, 7) = 203
    myArray(5, 8) = 149
    myArray(5, 9) = 119.5
    myArray(5, 10) = 426
    myArray(5, 11) = 290
    myArray(5, 12) = 15.5
    myArray(5, 13) = 17.4
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 41.4
    myArray(6, 3) = 2.3
    myArray(6, 4) = 14.1
    myArray(6, 5) = 62
    myArray(6, 6) = 40
    myArray(6, 7) = 248.5
    myArray(6, 8) = 248.5
    myArray(6, 9) = 543
    myArray(6, 10) = 238.5
    myArray(6, 11) = 39
    myArray(6, 12) = 29.5
    myArray(6, 13) = 13.8
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 65
    myArray(7, 3) = 69.5
    myArray(7, 4) = 9.8
    myArray(7, 5) = 25
    myArray(7, 6) = 23.5
    myArray(7, 7) = 132
    myArray(7, 8) = 216
    myArray(7, 9) = 98
    myArray(7, 10) = 10.5
    myArray(7, 11) = 76.5
    myArray(7, 12) = 10.5
    myArray(7, 13) = 16.3
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 72.3
    myArray(8, 3) = 6
    myArray(8, 4) = 32.5
    myArray(8, 5) = 142.5
    myArray(8, 6) = 159
    myArray(8, 7) = 70.5
    myArray(8, 8) = 208
    myArray(8, 9) = 358.5
    myArray(8, 10) = 57.5
    myArray(8, 11) = 78.5
    myArray(8, 12) = 31.5
    myArray(8, 13) = 57.2
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 24.2
    myArray(9, 3) = 59
    myArray(9, 4) = 52
    myArray(9, 5) = 208.5
    myArray(9, 6) = 144.5
    myArray(9, 7) = 228
    myArray(9, 8) = 626.5
    myArray(9, 9) = 202
    myArray(9, 10) = 167.5
    myArray(9, 11) = 24.5
    myArray(9, 12) = 29.5
    myArray(9, 13) = 13.8
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 18.1
    myArray(10, 3) = 26.2
    myArray(10, 4) = 63.1
    myArray(10, 5) = 73.5
    myArray(10, 6) = 109
    myArray(10, 7) = 388
    myArray(10, 8) = 296
    myArray(10, 9) = 249
    myArray(10, 10) = 176.5
    myArray(10, 11) = 1
    myArray(10, 12) = 50.5
    myArray(10, 13) = 43
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 6
    myArray(11, 3) = 39
    myArray(11, 4) = 26.5
    myArray(11, 5) = 75
    myArray(11, 6) = 65.5
    myArray(11, 7) = 186
    myArray(11, 8) = 448.5
    myArray(11, 9) = 381.5
    myArray(11, 10) = 225.5
    myArray(11, 11) = 30.5
    myArray(11, 12) = 21
    myArray(11, 13) = 22
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 30.2
    myArray(12, 3) = 29.5
    myArray(12, 4) = 7.8
    myArray(12, 5) = 99
    myArray(12, 6) = 81.5
    myArray(12, 7) = 111
    myArray(12, 8) = 503
    myArray(12, 9) = 83.5
    myArray(12, 10) = 37.5
    myArray(12, 11) = 15
    myArray(12, 12) = 51
    myArray(12, 13) = 27.5
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 21.8
    myArray(13, 3) = 47.8
    myArray(13, 4) = 159
    myArray(13, 5) = 28
    myArray(13, 6) = 104
    myArray(13, 7) = 101
    myArray(13, 8) = 286
    myArray(13, 9) = 319.5
    myArray(13, 10) = 502.5
    myArray(13, 11) = 37
    myArray(13, 12) = 13
    myArray(13, 13) = 31.7
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 39.6
    myArray(14, 3) = 11.2
    myArray(14, 4) = 42.2
    myArray(14, 5) = 38.8
    myArray(14, 6) = 51.6
    myArray(14, 7) = 260
    myArray(14, 8) = 194.3
    myArray(14, 9) = 154
    myArray(14, 10) = 48.8
    myArray(14, 11) = 24.1
    myArray(14, 12) = 14.1
    myArray(14, 13) = 23.4
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 10.6
    myArray(15, 3) = 23.6
    myArray(15, 4) = 63.9
    myArray(15, 5) = 51
    myArray(15, 6) = 135.5
    myArray(15, 7) = 113.2
    myArray(15, 8) = 408
    myArray(15, 9) = 140.2
    myArray(15, 10) = 30.5
    myArray(15, 11) = 23.7
    myArray(15, 12) = 54.5
    myArray(15, 13) = 34.9
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 37.1
    myArray(16, 3) = 89.5
    myArray(16, 4) = 94.9
    myArray(16, 5) = 69.6
    myArray(16, 6) = 140.7
    myArray(16, 7) = 36.1
    myArray(16, 8) = 262.1
    myArray(16, 9) = 431.1
    myArray(16, 10) = 149.8
    myArray(16, 11) = 17.8
    myArray(16, 12) = 18.6
    myArray(16, 13) = 31
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 3.7
    myArray(17, 3) = 60.7
    myArray(17, 4) = 16
    myArray(17, 5) = 70
    myArray(17, 6) = 111.2
    myArray(17, 7) = 316
    myArray(17, 8) = 599.6
    myArray(17, 9) = 618.1
    myArray(17, 10) = 104.2
    myArray(17, 11) = 26.6
    myArray(17, 12) = 81.6
    myArray(17, 13) = 7
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 16
    myArray(18, 3) = 3.2
    myArray(18, 4) = 60.2
    myArray(18, 5) = 109.3
    myArray(18, 6) = 19.5
    myArray(18, 7) = 71.3
    myArray(18, 8) = 302.9
    myArray(18, 9) = 573.3
    myArray(18, 10) = 186.2
    myArray(18, 11) = 83
    myArray(18, 12) = 60.7
    myArray(18, 13) = 60.2
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 45.4
    myArray(19, 3) = 58.7
    myArray(19, 4) = 50.3
    myArray(19, 5) = 93.7
    myArray(19, 6) = 159
    myArray(19, 7) = 151.7
    myArray(19, 8) = 240.4
    myArray(19, 9) = 119.5
    myArray(19, 10) = 184.8
    myArray(19, 11) = 17.5
    myArray(19, 12) = 79.4
    myArray(19, 13) = 35.9
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 2.2
    myArray(20, 3) = 15.3
    myArray(20, 4) = 69.3
    myArray(20, 5) = 94.1
    myArray(20, 6) = 61.5
    myArray(20, 7) = 77.8
    myArray(20, 8) = 174.7
    myArray(20, 9) = 225.1
    myArray(20, 10) = 157.5
    myArray(20, 11) = 170.5
    myArray(20, 12) = 42.4
    myArray(20, 13) = 51.7
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 35.4
    myArray(21, 3) = 35.6
    myArray(21, 4) = 42.4
    myArray(21, 5) = 99.5
    myArray(21, 6) = 53.5
    myArray(21, 7) = 92.7
    myArray(21, 8) = 119.9
    myArray(21, 9) = 56.9
    myArray(21, 10) = 22
    myArray(21, 11) = 104
    myArray(21, 12) = 130
    myArray(21, 13) = 56.9
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 6.6
    myArray(22, 3) = 59.6
    myArray(22, 4) = 19
    myArray(22, 5) = 164.6
    myArray(22, 6) = 121.6
    myArray(22, 7) = 49.4
    myArray(22, 8) = 341.1
    myArray(22, 9) = 33.4
    myArray(22, 10) = 133.7
    myArray(22, 11) = 120.1
    myArray(22, 12) = 17.1
    myArray(22, 13) = 63.1
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 16
    myArray(23, 3) = 28.5
    myArray(23, 4) = 8.8
    myArray(23, 5) = 78.4
    myArray(23, 6) = 35.8
    myArray(23, 7) = 51.4
    myArray(23, 8) = 326.7
    myArray(23, 9) = 358.5
    myArray(23, 10) = 97.1
    myArray(23, 11) = 51.9
    myArray(23, 12) = 22.8
    myArray(23, 13) = 36.1
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 25
    myArray(24, 3) = 43.1
    myArray(24, 4) = 99.3
    myArray(24, 5) = 156.5
    myArray(24, 6) = 116.1
    myArray(24, 7) = 107.1
    myArray(24, 8) = 278.8
    myArray(24, 9) = 277
    myArray(24, 10) = 98.3
    myArray(24, 11) = 159.2
    myArray(24, 12) = 66
    myArray(24, 13) = 31.5
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.5
    myArray(25, 3) = 37.6
    myArray(25, 4) = 35
    myArray(25, 5) = 73.7
    myArray(25, 6) = 44.3
    myArray(25, 7) = 59.9
    myArray(25, 8) = 216.7
    myArray(25, 9) = 102.1
    myArray(25, 10) = 191.9
    myArray(25, 11) = 85.6
    myArray(25, 12) = 113.5
    myArray(25, 13) = 31.2
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 79.6
    myArray(26, 3) = 92.4
    myArray(26, 4) = 19.3
    myArray(26, 5) = 17.7
    myArray(26, 6) = 108.5
    myArray(26, 7) = 188.4
    myArray(26, 8) = 492.6
    myArray(26, 9) = 367.8
    myArray(26, 10) = 208.9
    myArray(26, 11) = 4.4
    myArray(26, 12) = 41.8
    myArray(26, 13) = 3.4
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 32.1
    myArray(27, 3) = 18.1
    myArray(27, 4) = 95.7
    myArray(27, 5) = 42.3
    myArray(27, 6) = 136.9
    myArray(27, 7) = 76.9
    myArray(27, 8) = 187.7
    myArray(27, 9) = 227.6
    myArray(27, 10) = 187.1
    myArray(27, 11) = 36.9
    myArray(27, 12) = 73.4
    myArray(27, 13) = 8.7
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 3.5
    myArray(28, 3) = 2.5
    myArray(28, 4) = 76.1
    myArray(28, 5) = 62.6
    myArray(28, 6) = 4
    myArray(28, 7) = 123.4
    myArray(28, 8) = 168.5
    myArray(28, 9) = 615.6
    myArray(28, 10) = 87
    myArray(28, 11) = 103.7
    myArray(28, 12) = 36.4
    myArray(28, 13) = 17.8
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 35.7
    myArray(29, 3) = 4.3
    myArray(29, 4) = 13.2
    myArray(29, 5) = 60.6
    myArray(29, 6) = 248.1
    myArray(29, 7) = 122.2
    myArray(29, 8) = 880.3
    myArray(29, 9) = 300.6
    myArray(29, 10) = 303
    myArray(29, 11) = 16.7
    myArray(29, 12) = 58.1
    myArray(29, 13) = 122.2
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 43.9
    myArray(30, 3) = 135.8
    myArray(30, 4) = 49.8
    myArray(30, 5) = 58.7
    myArray(30, 6) = 134
    myArray(30, 7) = 91.2
    myArray(30, 8) = 470.8
    myArray(30, 9) = 40.8
    myArray(30, 10) = 162.6
    myArray(30, 11) = 84.5
    myArray(30, 12) = 46
    myArray(30, 13) = 16.1

    data_BUYEO = myArray

End Function


Function data_BOEUN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1995
    myArray(1, 2) = 17
    myArray(1, 3) = 13.4
    myArray(1, 4) = 46.4
    myArray(1, 5) = 60.5
    myArray(1, 6) = 56
    myArray(1, 7) = 45.5
    myArray(1, 8) = 126.5
    myArray(1, 9) = 508
    myArray(1, 10) = 31
    myArray(1, 11) = 42
    myArray(1, 12) = 32.3
    myArray(1, 13) = 5.3
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 18.4
    myArray(2, 3) = 5.3
    myArray(2, 4) = 100.7
    myArray(2, 5) = 28.5
    myArray(2, 6) = 62.5
    myArray(2, 7) = 385.5
    myArray(2, 8) = 264
    myArray(2, 9) = 97.5
    myArray(2, 10) = 29
    myArray(2, 11) = 80
    myArray(2, 12) = 66.8
    myArray(2, 13) = 25.8
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 14.1
    myArray(3, 3) = 55
    myArray(3, 4) = 32
    myArray(3, 5) = 49
    myArray(3, 6) = 238
    myArray(3, 7) = 226
    myArray(3, 8) = 378.8
    myArray(3, 9) = 402.5
    myArray(3, 10) = 46
    myArray(3, 11) = 9.5
    myArray(3, 12) = 162.7
    myArray(3, 13) = 50.1
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 22.4
    myArray(4, 3) = 28
    myArray(4, 4) = 23.7
    myArray(4, 5) = 173.5
    myArray(4, 6) = 103.5
    myArray(4, 7) = 256
    myArray(4, 8) = 311.5
    myArray(4, 9) = 894
    myArray(4, 10) = 180.5
    myArray(4, 11) = 54.5
    myArray(4, 12) = 33
    myArray(4, 13) = 4.5
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 1.5
    myArray(5, 3) = 6.9
    myArray(5, 4) = 72.5
    myArray(5, 5) = 118
    myArray(5, 6) = 108
    myArray(5, 7) = 206
    myArray(5, 8) = 136.5
    myArray(5, 9) = 249.1
    myArray(5, 10) = 306.8
    myArray(5, 11) = 144
    myArray(5, 12) = 15.6
    myArray(5, 13) = 14.3
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 34.3
    myArray(6, 3) = 3.7
    myArray(6, 4) = 20.7
    myArray(6, 5) = 60
    myArray(6, 6) = 44.5
    myArray(6, 7) = 244.6
    myArray(6, 8) = 384.2
    myArray(6, 9) = 348.5
    myArray(6, 10) = 229.5
    myArray(6, 11) = 25
    myArray(6, 12) = 38
    myArray(6, 13) = 16.2
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 49.4
    myArray(7, 3) = 60.5
    myArray(7, 4) = 12.7
    myArray(7, 5) = 11.5
    myArray(7, 6) = 18
    myArray(7, 7) = 259.5
    myArray(7, 8) = 139.5
    myArray(7, 9) = 127
    myArray(7, 10) = 42.5
    myArray(7, 11) = 81
    myArray(7, 12) = 9
    myArray(7, 13) = 23.8
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 89.1
    myArray(8, 3) = 11
    myArray(8, 4) = 35
    myArray(8, 5) = 178
    myArray(8, 6) = 126
    myArray(8, 7) = 49
    myArray(8, 8) = 156.5
    myArray(8, 9) = 418
    myArray(8, 10) = 114.5
    myArray(8, 11) = 42
    myArray(8, 12) = 20.1
    myArray(8, 13) = 45.7
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 22.7
    myArray(9, 3) = 67
    myArray(9, 4) = 44.5
    myArray(9, 5) = 200.5
    myArray(9, 6) = 159
    myArray(9, 7) = 192
    myArray(9, 8) = 689
    myArray(9, 9) = 310
    myArray(9, 10) = 306.5
    myArray(9, 11) = 32
    myArray(9, 12) = 36.5
    myArray(9, 13) = 19.5
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 15.4
    myArray(10, 3) = 33.8
    myArray(10, 4) = 59.9
    myArray(10, 5) = 84.2
    myArray(10, 6) = 117
    myArray(10, 7) = 327
    myArray(10, 8) = 295.5
    myArray(10, 9) = 203.5
    myArray(10, 10) = 136
    myArray(10, 11) = 5.5
    myArray(10, 12) = 38.5
    myArray(10, 13) = 49.1
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 9.2
    myArray(11, 3) = 20
    myArray(11, 4) = 38
    myArray(11, 5) = 62.5
    myArray(11, 6) = 62
    myArray(11, 7) = 215
    myArray(11, 8) = 400
    myArray(11, 9) = 493.5
    myArray(11, 10) = 177.5
    myArray(11, 11) = 31
    myArray(11, 12) = 15
    myArray(11, 13) = 12.6
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 27
    myArray(12, 3) = 37.8
    myArray(12, 4) = 11.9
    myArray(12, 5) = 92.5
    myArray(12, 6) = 107
    myArray(12, 7) = 113
    myArray(12, 8) = 511.5
    myArray(12, 9) = 143
    myArray(12, 10) = 27.5
    myArray(12, 11) = 25
    myArray(12, 12) = 76.5
    myArray(12, 13) = 23.5
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 11.6
    myArray(13, 3) = 42.5
    myArray(13, 4) = 119.8
    myArray(13, 5) = 40
    myArray(13, 6) = 105
    myArray(13, 7) = 142
    myArray(13, 8) = 282.5
    myArray(13, 9) = 366
    myArray(13, 10) = 351.5
    myArray(13, 11) = 37
    myArray(13, 12) = 10.6
    myArray(13, 13) = 23.6
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 46.9
    myArray(14, 3) = 7
    myArray(14, 4) = 28.7
    myArray(14, 5) = 23.3
    myArray(14, 6) = 83.5
    myArray(14, 7) = 152.1
    myArray(14, 8) = 212.5
    myArray(14, 9) = 311.1
    myArray(14, 10) = 51.8
    myArray(14, 11) = 19.4
    myArray(14, 12) = 11.3
    myArray(14, 13) = 14.3
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 10.5
    myArray(15, 3) = 23.9
    myArray(15, 4) = 53
    myArray(15, 5) = 37
    myArray(15, 6) = 147.5
    myArray(15, 7) = 137
    myArray(15, 8) = 404
    myArray(15, 9) = 124
    myArray(15, 10) = 62.5
    myArray(15, 11) = 27.2
    myArray(15, 12) = 48
    myArray(15, 13) = 37.6
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 32
    myArray(16, 3) = 74.1
    myArray(16, 4) = 82.9
    myArray(16, 5) = 74
    myArray(16, 6) = 108
    myArray(16, 7) = 20.6
    myArray(16, 8) = 199.6
    myArray(16, 9) = 357.7
    myArray(16, 10) = 249.7
    myArray(16, 11) = 26.8
    myArray(16, 12) = 9
    myArray(16, 13) = 28.5
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 2.6
    myArray(17, 3) = 38.9
    myArray(17, 4) = 19.8
    myArray(17, 5) = 94.5
    myArray(17, 6) = 153.8
    myArray(17, 7) = 412
    myArray(17, 8) = 535.1
    myArray(17, 9) = 296.7
    myArray(17, 10) = 105.8
    myArray(17, 11) = 55.5
    myArray(17, 12) = 84.5
    myArray(17, 13) = 11.5
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 16.1
    myArray(18, 3) = 1
    myArray(18, 4) = 79
    myArray(18, 5) = 92.4
    myArray(18, 6) = 44.6
    myArray(18, 7) = 79.4
    myArray(18, 8) = 294.6
    myArray(18, 9) = 488.9
    myArray(18, 10) = 218.5
    myArray(18, 11) = 83.6
    myArray(18, 12) = 68.2
    myArray(18, 13) = 56
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 45.2
    myArray(19, 3) = 37.2
    myArray(19, 4) = 50.9
    myArray(19, 5) = 90.3
    myArray(19, 6) = 107.8
    myArray(19, 7) = 160.9
    myArray(19, 8) = 245.2
    myArray(19, 9) = 114
    myArray(19, 10) = 139.9
    myArray(19, 11) = 32
    myArray(19, 12) = 67.2
    myArray(19, 13) = 35.3
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 8
    myArray(20, 3) = 5.5
    myArray(20, 4) = 72.2
    myArray(20, 5) = 46.7
    myArray(20, 6) = 53
    myArray(20, 7) = 103.4
    myArray(20, 8) = 164.7
    myArray(20, 9) = 288.7
    myArray(20, 10) = 106.7
    myArray(20, 11) = 171.3
    myArray(20, 12) = 43.2
    myArray(20, 13) = 25.8
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 25.1
    myArray(21, 3) = 31.8
    myArray(21, 4) = 45
    myArray(21, 5) = 92.1
    myArray(21, 6) = 31.8
    myArray(21, 7) = 73.4
    myArray(21, 8) = 156.7
    myArray(21, 9) = 85.2
    myArray(21, 10) = 38.8
    myArray(21, 11) = 84.5
    myArray(21, 12) = 109.8
    myArray(21, 13) = 42.8
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 9.7
    myArray(22, 3) = 39.8
    myArray(22, 4) = 45.5
    myArray(22, 5) = 149.2
    myArray(22, 6) = 88.5
    myArray(22, 7) = 48.8
    myArray(22, 8) = 494.9
    myArray(22, 9) = 53
    myArray(22, 10) = 149.4
    myArray(22, 11) = 135.7
    myArray(22, 12) = 33.5
    myArray(22, 13) = 43.6
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 16.4
    myArray(23, 3) = 47
    myArray(23, 4) = 16.9
    myArray(23, 5) = 61
    myArray(23, 6) = 28
    myArray(23, 7) = 87.6
    myArray(23, 8) = 572
    myArray(23, 9) = 315.8
    myArray(23, 10) = 108.6
    myArray(23, 11) = 28.5
    myArray(23, 12) = 13.4
    myArray(23, 13) = 27.6
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 25
    myArray(24, 3) = 28.5
    myArray(24, 4) = 93.5
    myArray(24, 5) = 139.2
    myArray(24, 6) = 103.9
    myArray(24, 7) = 75.3
    myArray(24, 8) = 224.9
    myArray(24, 9) = 386
    myArray(24, 10) = 134.7
    myArray(24, 11) = 108.3
    myArray(24, 12) = 52.2
    myArray(24, 13) = 38
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.9
    myArray(25, 3) = 39.8
    myArray(25, 4) = 29
    myArray(25, 5) = 100.8
    myArray(25, 6) = 42.8
    myArray(25, 7) = 73.9
    myArray(25, 8) = 226.1
    myArray(25, 9) = 132.7
    myArray(25, 10) = 186.2
    myArray(25, 11) = 100.6
    myArray(25, 12) = 86.2
    myArray(25, 13) = 27.2
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 71.8
    myArray(26, 3) = 80.8
    myArray(26, 4) = 23.6
    myArray(26, 5) = 35.9
    myArray(26, 6) = 89.1
    myArray(26, 7) = 171.2
    myArray(26, 8) = 500.8
    myArray(26, 9) = 587.8
    myArray(26, 10) = 162.2
    myArray(26, 11) = 3
    myArray(26, 12) = 37
    myArray(26, 13) = 5.6
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 19.8
    myArray(27, 3) = 14.1
    myArray(27, 4) = 93
    myArray(27, 5) = 52.5
    myArray(27, 6) = 154.5
    myArray(27, 7) = 76.8
    myArray(27, 8) = 163.9
    myArray(27, 9) = 275.8
    myArray(27, 10) = 162.2
    myArray(27, 11) = 33.8
    myArray(27, 12) = 44.9
    myArray(27, 13) = 6.4
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 4.3
    myArray(28, 3) = 4.3
    myArray(28, 4) = 98.2
    myArray(28, 5) = 61.1
    myArray(28, 6) = 5.6
    myArray(28, 7) = 98.3
    myArray(28, 8) = 160.5
    myArray(28, 9) = 391.9
    myArray(28, 10) = 71.7
    myArray(28, 11) = 89.7
    myArray(28, 12) = 39.9
    myArray(28, 13) = 18
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 27.3
    myArray(29, 3) = 5.3
    myArray(29, 4) = 25
    myArray(29, 5) = 51.9
    myArray(29, 6) = 146.4
    myArray(29, 7) = 169.3
    myArray(29, 8) = 796.9
    myArray(29, 9) = 242.9
    myArray(29, 10) = 243.2
    myArray(29, 11) = 14.5
    myArray(29, 12) = 39.6
    myArray(29, 13) = 108.1
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 34.1
    myArray(30, 3) = 79.6
    myArray(30, 4) = 50.7
    myArray(30, 5) = 45.8
    myArray(30, 6) = 128.4
    myArray(30, 7) = 70.1
    myArray(30, 8) = 513.8
    myArray(30, 9) = 47.6
    myArray(30, 10) = 148.1
    myArray(30, 11) = 99.4
    myArray(30, 12) = 29.6
    myArray(30, 13) = 5.6

    data_BOEUN = myArray

End Function

Function data_BORYUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 15.7
    myArray(1, 3) = 11
    myArray(1, 4) = 19.6
    myArray(1, 5) = 65.5
    myArray(1, 6) = 49.5
    myArray(1, 7) = 26.5
    myArray(1, 8) = 144.5
    myArray(1, 9) = 996.5
    myArray(1, 10) = 70.5
    myArray(1, 11) = 24.5
    myArray(1, 12) = 23
    myArray(1, 13) = 12.7
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 33.4
    myArray(2, 3) = 6.8
    myArray(2, 4) = 104.5
    myArray(2, 5) = 34
    myArray(2, 6) = 22.5
    myArray(2, 7) = 235
    myArray(2, 8) = 192.5
    myArray(2, 9) = 44.5
    myArray(2, 10) = 14
    myArray(2, 11) = 106.5
    myArray(2, 12) = 74.2
    myArray(2, 13) = 31.7
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 15.1
    myArray(3, 3) = 38.4
    myArray(3, 4) = 30.5
    myArray(3, 5) = 57.5
    myArray(3, 6) = 203
    myArray(3, 7) = 272
    myArray(3, 8) = 353.5
    myArray(3, 9) = 211.5
    myArray(3, 10) = 23
    myArray(3, 11) = 10
    myArray(3, 12) = 193.5
    myArray(3, 13) = 34.3
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 29.9
    myArray(4, 3) = 40.2
    myArray(4, 4) = 30.5
    myArray(4, 5) = 138
    myArray(4, 6) = 100
    myArray(4, 7) = 209.5
    myArray(4, 8) = 263
    myArray(4, 9) = 341.7
    myArray(4, 10) = 150.3
    myArray(4, 11) = 61
    myArray(4, 12) = 29.3
    myArray(4, 13) = 3.8
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 7.9
    myArray(5, 3) = 9.5
    myArray(5, 4) = 71
    myArray(5, 5) = 88.5
    myArray(5, 6) = 124.5
    myArray(5, 7) = 192.5
    myArray(5, 8) = 98
    myArray(5, 9) = 180
    myArray(5, 10) = 292.5
    myArray(5, 11) = 169
    myArray(5, 12) = 24.9
    myArray(5, 13) = 25.8
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 42.1
    myArray(6, 3) = 3.2
    myArray(6, 4) = 7
    myArray(6, 5) = 35
    myArray(6, 6) = 53.5
    myArray(6, 7) = 159.5
    myArray(6, 8) = 155
    myArray(6, 9) = 701.5
    myArray(6, 10) = 241
    myArray(6, 11) = 46
    myArray(6, 12) = 39.5
    myArray(6, 13) = 32.1
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 73.3
    myArray(7, 3) = 46
    myArray(7, 4) = 15.9
    myArray(7, 5) = 26
    myArray(7, 6) = 17
    myArray(7, 7) = 129
    myArray(7, 8) = 286.5
    myArray(7, 9) = 170
    myArray(7, 10) = 10
    myArray(7, 11) = 85
    myArray(7, 12) = 13
    myArray(7, 13) = 32
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 50.8
    myArray(8, 3) = 5.5
    myArray(8, 4) = 32
    myArray(8, 5) = 169
    myArray(8, 6) = 155.5
    myArray(8, 7) = 72
    myArray(8, 8) = 217.5
    myArray(8, 9) = 477
    myArray(8, 10) = 27
    myArray(8, 11) = 134
    myArray(8, 12) = 61.1
    myArray(8, 13) = 51.8
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 30.7
    myArray(9, 3) = 44.5
    myArray(9, 4) = 39.5
    myArray(9, 5) = 168.5
    myArray(9, 6) = 78.5
    myArray(9, 7) = 153
    myArray(9, 8) = 309.5
    myArray(9, 9) = 310
    myArray(9, 10) = 128
    myArray(9, 11) = 23
    myArray(9, 12) = 45.5
    myArray(9, 13) = 13
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 22.1
    myArray(10, 3) = 28.5
    myArray(10, 4) = 45.7
    myArray(10, 5) = 58
    myArray(10, 6) = 105.5
    myArray(10, 7) = 234.5
    myArray(10, 8) = 263.5
    myArray(10, 9) = 164
    myArray(10, 10) = 195
    myArray(10, 11) = 4
    myArray(10, 12) = 56.5
    myArray(10, 13) = 38.9
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 5.8
    myArray(11, 3) = 35.8
    myArray(11, 4) = 30
    myArray(11, 5) = 73.5
    myArray(11, 6) = 48.5
    myArray(11, 7) = 156
    myArray(11, 8) = 260.5
    myArray(11, 9) = 291.5
    myArray(11, 10) = 282.5
    myArray(11, 11) = 21
    myArray(11, 12) = 18
    myArray(11, 13) = 43.4
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 27
    myArray(12, 3) = 25.9
    myArray(12, 4) = 10.6
    myArray(12, 5) = 81.5
    myArray(12, 6) = 94.5
    myArray(12, 7) = 114.5
    myArray(12, 8) = 321
    myArray(12, 9) = 21.5
    myArray(12, 10) = 23.5
    myArray(12, 11) = 24.5
    myArray(12, 12) = 61.5
    myArray(12, 13) = 25.4
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 23.4
    myArray(13, 3) = 29.8
    myArray(13, 4) = 102
    myArray(13, 5) = 29.5
    myArray(13, 6) = 79
    myArray(13, 7) = 85
    myArray(13, 8) = 214
    myArray(13, 9) = 239.5
    myArray(13, 10) = 384
    myArray(13, 11) = 59
    myArray(13, 12) = 17.5
    myArray(13, 13) = 33.1
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 20.9
    myArray(14, 3) = 10.8
    myArray(14, 4) = 48.2
    myArray(14, 5) = 40.5
    myArray(14, 6) = 78.9
    myArray(14, 7) = 101.3
    myArray(14, 8) = 257.2
    myArray(14, 9) = 119.5
    myArray(14, 10) = 46.9
    myArray(14, 11) = 26.7
    myArray(14, 12) = 37.6
    myArray(14, 13) = 25
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 18.5
    myArray(15, 3) = 23.3
    myArray(15, 4) = 55.1
    myArray(15, 5) = 41.5
    myArray(15, 6) = 154.5
    myArray(15, 7) = 115.1
    myArray(15, 8) = 320.9
    myArray(15, 9) = 176.6
    myArray(15, 10) = 25.5
    myArray(15, 11) = 39.5
    myArray(15, 12) = 52.9
    myArray(15, 13) = 58
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 30.1
    myArray(16, 3) = 73.5
    myArray(16, 4) = 75.9
    myArray(16, 5) = 58.5
    myArray(16, 6) = 122.8
    myArray(16, 7) = 60.8
    myArray(16, 8) = 396.5
    myArray(16, 9) = 402.7
    myArray(16, 10) = 213.1
    myArray(16, 11) = 19.2
    myArray(16, 12) = 16.3
    myArray(16, 13) = 32.9
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 11.1
    myArray(17, 3) = 37.5
    myArray(17, 4) = 18
    myArray(17, 5) = 72.1
    myArray(17, 6) = 115.3
    myArray(17, 7) = 318
    myArray(17, 8) = 723.1
    myArray(17, 9) = 289.4
    myArray(17, 10) = 70.8
    myArray(17, 11) = 13.9
    myArray(17, 12) = 61.3
    myArray(17, 13) = 12.5
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 24.2
    myArray(18, 3) = 9.2
    myArray(18, 4) = 45
    myArray(18, 5) = 68.9
    myArray(18, 6) = 14.6
    myArray(18, 7) = 76.8
    myArray(18, 8) = 231.1
    myArray(18, 9) = 450.2
    myArray(18, 10) = 207.7
    myArray(18, 11) = 65
    myArray(18, 12) = 61.1
    myArray(18, 13) = 65.2
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 28.4
    myArray(19, 3) = 40.7
    myArray(19, 4) = 53.4
    myArray(19, 5) = 68.2
    myArray(19, 6) = 116.6
    myArray(19, 7) = 159.9
    myArray(19, 8) = 267.5
    myArray(19, 9) = 214.6
    myArray(19, 10) = 320
    myArray(19, 11) = 10.9
    myArray(19, 12) = 81.1
    myArray(19, 13) = 26.4
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 3.4
    myArray(20, 3) = 20.5
    myArray(20, 4) = 56.3
    myArray(20, 5) = 70
    myArray(20, 6) = 47.1
    myArray(20, 7) = 125.8
    myArray(20, 8) = 104
    myArray(20, 9) = 168.5
    myArray(20, 10) = 152
    myArray(20, 11) = 156
    myArray(20, 12) = 39.9
    myArray(20, 13) = 66.6
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 29.9
    myArray(21, 3) = 23.4
    myArray(21, 4) = 30.9
    myArray(21, 5) = 129.7
    myArray(21, 6) = 38.8
    myArray(21, 7) = 83.9
    myArray(21, 8) = 94.7
    myArray(21, 9) = 30.2
    myArray(21, 10) = 13.3
    myArray(21, 11) = 90
    myArray(21, 12) = 155.6
    myArray(21, 13) = 65
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 7.8
    myArray(22, 3) = 54.2
    myArray(22, 4) = 18.7
    myArray(22, 5) = 105.1
    myArray(22, 6) = 146.5
    myArray(22, 7) = 23.7
    myArray(22, 8) = 200.2
    myArray(22, 9) = 5.1
    myArray(22, 10) = 73.4
    myArray(22, 11) = 108
    myArray(22, 12) = 5.6
    myArray(22, 13) = 44.5
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 14.8
    myArray(23, 3) = 30.2
    myArray(23, 4) = 14.4
    myArray(23, 5) = 57.6
    myArray(23, 6) = 58.9
    myArray(23, 7) = 21.1
    myArray(23, 8) = 278.1
    myArray(23, 9) = 210
    myArray(23, 10) = 90
    myArray(23, 11) = 26.6
    myArray(23, 12) = 15.9
    myArray(23, 13) = 38.6
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 15
    myArray(24, 3) = 33.6
    myArray(24, 4) = 92
    myArray(24, 5) = 128.1
    myArray(24, 6) = 104.5
    myArray(24, 7) = 71
    myArray(24, 8) = 262.7
    myArray(24, 9) = 239.6
    myArray(24, 10) = 158.2
    myArray(24, 11) = 154.7
    myArray(24, 12) = 46.7
    myArray(24, 13) = 31.1
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 1.9
    myArray(25, 3) = 17.8
    myArray(25, 4) = 18.2
    myArray(25, 5) = 71.9
    myArray(25, 6) = 31.3
    myArray(25, 7) = 56
    myArray(25, 8) = 149
    myArray(25, 9) = 131.3
    myArray(25, 10) = 118.7
    myArray(25, 11) = 63.9
    myArray(25, 12) = 130.6
    myArray(25, 13) = 31.3
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 49.4
    myArray(26, 3) = 75.3
    myArray(26, 4) = 22.8
    myArray(26, 5) = 16.5
    myArray(26, 6) = 92.4
    myArray(26, 7) = 139.7
    myArray(26, 8) = 345.9
    myArray(26, 9) = 321.5
    myArray(26, 10) = 177.1
    myArray(26, 11) = 16.2
    myArray(26, 12) = 35.4
    myArray(26, 13) = 9.7
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 32
    myArray(27, 3) = 18.7
    myArray(27, 4) = 76.1
    myArray(27, 5) = 43.4
    myArray(27, 6) = 110
    myArray(27, 7) = 55
    myArray(27, 8) = 131.3
    myArray(27, 9) = 253.7
    myArray(27, 10) = 215.9
    myArray(27, 11) = 39.6
    myArray(27, 12) = 117.8
    myArray(27, 13) = 14.4
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 8.4
    myArray(28, 3) = 5.3
    myArray(28, 4) = 60.9
    myArray(28, 5) = 34.8
    myArray(28, 6) = 5.7
    myArray(28, 7) = 225
    myArray(28, 8) = 119.7
    myArray(28, 9) = 637.1
    myArray(28, 10) = 102
    myArray(28, 11) = 112
    myArray(28, 12) = 23.3
    myArray(28, 13) = 14
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 19.2
    myArray(29, 3) = 0.4
    myArray(29, 4) = 7.2
    myArray(29, 5) = 42.4
    myArray(29, 6) = 190.8
    myArray(29, 7) = 95.1
    myArray(29, 8) = 772.2
    myArray(29, 9) = 107.8
    myArray(29, 10) = 288.4
    myArray(29, 11) = 25.5
    myArray(29, 12) = 65.2
    myArray(29, 13) = 112.5
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 29.5
    myArray(30, 3) = 92.6
    myArray(30, 4) = 35.8
    myArray(30, 5) = 48.8
    myArray(30, 6) = 130.9
    myArray(30, 7) = 89
    myArray(30, 8) = 557.3
    myArray(30, 9) = 149.7
    myArray(30, 10) = 234.9
    myArray(30, 11) = 58.3
    myArray(30, 12) = 49
    myArray(30, 13) = 18.5
    
    data_BORYUNG = myArray

End Function

Function data_DAEJEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1995
    myArray(1, 2) = 23.5
    myArray(1, 3) = 16.9
    myArray(1, 4) = 33.8
    myArray(1, 5) = 54.7
    myArray(1, 6) = 62.2
    myArray(1, 7) = 33.6
    myArray(1, 8) = 155.4
    myArray(1, 9) = 641.9
    myArray(1, 10) = 53.4
    myArray(1, 11) = 36
    myArray(1, 12) = 17.5
    myArray(1, 13) = 7.3
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 32.7
    myArray(2, 3) = 4.4
    myArray(2, 4) = 138
    myArray(2, 5) = 49.8
    myArray(2, 6) = 62.9
    myArray(2, 7) = 411.4
    myArray(2, 8) = 257.7
    myArray(2, 9) = 114.4
    myArray(2, 10) = 11.4
    myArray(2, 11) = 90.8
    myArray(2, 12) = 77.1
    myArray(2, 13) = 28.6
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 15.6
    myArray(3, 3) = 51.1
    myArray(3, 4) = 37.1
    myArray(3, 5) = 55.4
    myArray(3, 6) = 200.9
    myArray(3, 7) = 267.5
    myArray(3, 8) = 424.2
    myArray(3, 9) = 463.5
    myArray(3, 10) = 30.2
    myArray(3, 11) = 7.7
    myArray(3, 12) = 168.2
    myArray(3, 13) = 44.5
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 33.3
    myArray(4, 3) = 36.3
    myArray(4, 4) = 31.1
    myArray(4, 5) = 154.3
    myArray(4, 6) = 119.5
    myArray(4, 7) = 297.2
    myArray(4, 8) = 256.1
    myArray(4, 9) = 781.7
    myArray(4, 10) = 254.7
    myArray(4, 11) = 71.5
    myArray(4, 12) = 31.6
    myArray(4, 13) = 2.7
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 1.8
    myArray(5, 3) = 12.2
    myArray(5, 4) = 79.4
    myArray(5, 5) = 103
    myArray(5, 6) = 116.8
    myArray(5, 7) = 245.7
    myArray(5, 8) = 137.8
    myArray(5, 9) = 203
    myArray(5, 10) = 359.5
    myArray(5, 11) = 171.6
    myArray(5, 12) = 16.5
    myArray(5, 13) = 7.9
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 27.5
    myArray(6, 3) = 4.1
    myArray(6, 4) = 17.8
    myArray(6, 5) = 67.8
    myArray(6, 6) = 54.3
    myArray(6, 7) = 238.3
    myArray(6, 8) = 470.1
    myArray(6, 9) = 473.6
    myArray(6, 10) = 263.2
    myArray(6, 11) = 24.6
    myArray(6, 12) = 44.6
    myArray(6, 13) = 21.6
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 61.2
    myArray(7, 3) = 70
    myArray(7, 4) = 16
    myArray(7, 5) = 20.4
    myArray(7, 6) = 30.2
    myArray(7, 7) = 234.2
    myArray(7, 8) = 171
    myArray(7, 9) = 78.1
    myArray(7, 10) = 25.2
    myArray(7, 11) = 91.2
    myArray(7, 12) = 10.8
    myArray(7, 13) = 20.4
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 92.1
    myArray(8, 3) = 12
    myArray(8, 4) = 33.5
    myArray(8, 5) = 155.5
    myArray(8, 6) = 130.5
    myArray(8, 7) = 55.4
    myArray(8, 8) = 149.1
    myArray(8, 9) = 538.8
    myArray(8, 10) = 77
    myArray(8, 11) = 67.8
    myArray(8, 12) = 24
    myArray(8, 13) = 43
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 11.2
    myArray(9, 3) = 59.2
    myArray(9, 4) = 44.2
    myArray(9, 5) = 217.5
    myArray(9, 6) = 119.5
    myArray(9, 7) = 186.4
    myArray(9, 8) = 576.3
    myArray(9, 9) = 254.9
    myArray(9, 10) = 208.5
    myArray(9, 11) = 21.5
    myArray(9, 12) = 32.6
    myArray(9, 13) = 17.1
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 10.9
    myArray(10, 3) = 30.6
    myArray(10, 4) = 83.2
    myArray(10, 5) = 73.1
    myArray(10, 6) = 109
    myArray(10, 7) = 383.5
    myArray(10, 8) = 391
    myArray(10, 9) = 198.3
    myArray(10, 10) = 133.7
    myArray(10, 11) = 5
    myArray(10, 12) = 37.1
    myArray(10, 13) = 41.1
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 6
    myArray(11, 3) = 37.5
    myArray(11, 4) = 38.8
    myArray(11, 5) = 48.5
    myArray(11, 6) = 60.5
    myArray(11, 7) = 209.6
    myArray(11, 8) = 463.3
    myArray(11, 9) = 499.5
    myArray(11, 10) = 226.4
    myArray(11, 11) = 30.5
    myArray(11, 12) = 20.3
    myArray(11, 13) = 15.2
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 31.2
    myArray(12, 3) = 33.1
    myArray(12, 4) = 8.1
    myArray(12, 5) = 94.2
    myArray(12, 6) = 119.7
    myArray(12, 7) = 131
    myArray(12, 8) = 531
    myArray(12, 9) = 113.6
    myArray(12, 10) = 24.1
    myArray(12, 11) = 19.3
    myArray(12, 12) = 60
    myArray(12, 13) = 29.9
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 14
    myArray(13, 3) = 45
    myArray(13, 4) = 117.5
    myArray(13, 5) = 28.6
    myArray(13, 6) = 130.1
    myArray(13, 7) = 133
    myArray(13, 8) = 275.7
    myArray(13, 9) = 373
    myArray(13, 10) = 549.9
    myArray(13, 11) = 47.4
    myArray(13, 12) = 9.8
    myArray(13, 13) = 26.9
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 45.3
    myArray(14, 3) = 9.1
    myArray(14, 4) = 29.1
    myArray(14, 5) = 34.4
    myArray(14, 6) = 59.2
    myArray(14, 7) = 148.3
    myArray(14, 8) = 253.4
    myArray(14, 9) = 325.2
    myArray(14, 10) = 85.5
    myArray(14, 11) = 19.6
    myArray(14, 12) = 12.1
    myArray(14, 13) = 16.4
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 15.4
    myArray(15, 3) = 27.5
    myArray(15, 4) = 60.3
    myArray(15, 5) = 34.5
    myArray(15, 6) = 124.3
    myArray(15, 7) = 87.3
    myArray(15, 8) = 429.2
    myArray(15, 9) = 148.3
    myArray(15, 10) = 46
    myArray(15, 11) = 24.7
    myArray(15, 12) = 54.7
    myArray(15, 13) = 38.2
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 46.4
    myArray(16, 3) = 81.5
    myArray(16, 4) = 100.1
    myArray(16, 5) = 88.5
    myArray(16, 6) = 117.6
    myArray(16, 7) = 65.6
    myArray(16, 8) = 223.1
    myArray(16, 9) = 376.4
    myArray(16, 10) = 250.5
    myArray(16, 11) = 17.9
    myArray(16, 12) = 16.4
    myArray(16, 13) = 35.7
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 4
    myArray(17, 3) = 44.8
    myArray(17, 4) = 19
    myArray(17, 5) = 71
    myArray(17, 6) = 162
    myArray(17, 7) = 391.6
    myArray(17, 8) = 587.3
    myArray(17, 9) = 420.3
    myArray(17, 10) = 91.7
    myArray(17, 11) = 37
    myArray(17, 12) = 103.2
    myArray(17, 13) = 11.5
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 16.4
    myArray(18, 3) = 2.5
    myArray(18, 4) = 54.6
    myArray(18, 5) = 66.2
    myArray(18, 6) = 24
    myArray(18, 7) = 57.8
    myArray(18, 8) = 277.6
    myArray(18, 9) = 463.6
    myArray(18, 10) = 242.5
    myArray(18, 11) = 81.3
    myArray(18, 12) = 58.4
    myArray(18, 13) = 64.6
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 46.2
    myArray(19, 3) = 54.2
    myArray(19, 4) = 52.8
    myArray(19, 5) = 86.8
    myArray(19, 6) = 110.4
    myArray(19, 7) = 162.6
    myArray(19, 8) = 218.7
    myArray(19, 9) = 126.6
    myArray(19, 10) = 146.4
    myArray(19, 11) = 19.6
    myArray(19, 12) = 63.1
    myArray(19, 13) = 32.8
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 6.5
    myArray(20, 3) = 8.5
    myArray(20, 4) = 67.2
    myArray(20, 5) = 59.4
    myArray(20, 6) = 49.7
    myArray(20, 7) = 143.7
    myArray(20, 8) = 177.2
    myArray(20, 9) = 240.9
    myArray(20, 10) = 118
    myArray(20, 11) = 169.4
    myArray(20, 12) = 40.7
    myArray(20, 13) = 36.5
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 31.5
    myArray(21, 3) = 27
    myArray(21, 4) = 44.7
    myArray(21, 5) = 95.2
    myArray(21, 6) = 28.9
    myArray(21, 7) = 119.8
    myArray(21, 8) = 145.6
    myArray(21, 9) = 51.6
    myArray(21, 10) = 18.5
    myArray(21, 11) = 94.1
    myArray(21, 12) = 126.1
    myArray(21, 13) = 39.7
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 11.6
    myArray(22, 3) = 45.8
    myArray(22, 4) = 40.3
    myArray(22, 5) = 154.9
    myArray(22, 6) = 85.1
    myArray(22, 7) = 62.5
    myArray(22, 8) = 367.9
    myArray(22, 9) = 57.4
    myArray(22, 10) = 196
    myArray(22, 11) = 122.6
    myArray(22, 12) = 29.5
    myArray(22, 13) = 54.8
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 15
    myArray(23, 3) = 42
    myArray(23, 4) = 11.6
    myArray(23, 5) = 77.7
    myArray(23, 6) = 29.3
    myArray(23, 7) = 35.3
    myArray(23, 8) = 434.5
    myArray(23, 9) = 293.8
    myArray(23, 10) = 111.4
    myArray(23, 11) = 28.3
    myArray(23, 12) = 15.1
    myArray(23, 13) = 33.5
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 23.9
    myArray(24, 3) = 40.5
    myArray(24, 4) = 108.4
    myArray(24, 5) = 155.3
    myArray(24, 6) = 95.9
    myArray(24, 7) = 115.8
    myArray(24, 8) = 226.9
    myArray(24, 9) = 408.6
    myArray(24, 10) = 149.4
    myArray(24, 11) = 133.9
    myArray(24, 12) = 49.8
    myArray(24, 13) = 33.7
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 1.7
    myArray(25, 3) = 46.3
    myArray(25, 4) = 33.7
    myArray(25, 5) = 91.6
    myArray(25, 6) = 35.6
    myArray(25, 7) = 77.9
    myArray(25, 8) = 199
    myArray(25, 9) = 104.3
    myArray(25, 10) = 167
    myArray(25, 11) = 106.1
    myArray(25, 12) = 94
    myArray(25, 13) = 27
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 78.5
    myArray(26, 3) = 91.2
    myArray(26, 4) = 24.4
    myArray(26, 5) = 17.8
    myArray(26, 6) = 80.4
    myArray(26, 7) = 192.5
    myArray(26, 8) = 544.9
    myArray(26, 9) = 361.6
    myArray(26, 10) = 173.6
    myArray(26, 11) = 3.2
    myArray(26, 12) = 41.8
    myArray(26, 13) = 4.1
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 23.6
    myArray(27, 3) = 14.1
    myArray(27, 4) = 95.2
    myArray(27, 5) = 47.4
    myArray(27, 6) = 134.2
    myArray(27, 7) = 105.9
    myArray(27, 8) = 151.8
    myArray(27, 9) = 289.2
    myArray(27, 10) = 161.2
    myArray(27, 11) = 40.8
    myArray(27, 12) = 41.7
    myArray(27, 13) = 4.4
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 1.2
    myArray(28, 3) = 1.4
    myArray(28, 4) = 74
    myArray(28, 5) = 69.7
    myArray(28, 6) = 8.1
    myArray(28, 7) = 117.6
    myArray(28, 8) = 195
    myArray(28, 9) = 496.1
    myArray(28, 10) = 90.2
    myArray(28, 11) = 89.3
    myArray(28, 12) = 45.8
    myArray(28, 13) = 14.7
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 28.4
    myArray(29, 3) = 5.4
    myArray(29, 4) = 23.8
    myArray(29, 5) = 54.5
    myArray(29, 6) = 192.9
    myArray(29, 7) = 147.5
    myArray(29, 8) = 776.3
    myArray(29, 9) = 326.9
    myArray(29, 10) = 310.2
    myArray(29, 11) = 12.2
    myArray(29, 12) = 40.3
    myArray(29, 13) = 124.1
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 47.8
    myArray(30, 3) = 93.9
    myArray(30, 4) = 57.9
    myArray(30, 5) = 32.7
    myArray(30, 6) = 126.8
    myArray(30, 7) = 76.9
    myArray(30, 8) = 485.1
    myArray(30, 9) = 87.3
    myArray(30, 10) = 204.6
    myArray(30, 11) = 109.2
    myArray(30, 12) = 34.6
    myArray(30, 13) = 3.7
    
    data_DAEJEON = myArray
End Function



Function data_GEUMSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1995
    myArray(1, 2) = 23.2
    myArray(1, 3) = 17.1
    myArray(1, 4) = 46.9
    myArray(1, 5) = 65.5
    myArray(1, 6) = 35.5
    myArray(1, 7) = 54
    myArray(1, 8) = 83.5
    myArray(1, 9) = 579.5
    myArray(1, 10) = 47.5
    myArray(1, 11) = 23.5
    myArray(1, 12) = 31
    myArray(1, 13) = 4.6
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 25.4
    myArray(2, 3) = 2.9
    myArray(2, 4) = 123
    myArray(2, 5) = 42.5
    myArray(2, 6) = 37.5
    myArray(2, 7) = 546
    myArray(2, 8) = 174
    myArray(2, 9) = 130
    myArray(2, 10) = 12.5
    myArray(2, 11) = 75.5
    myArray(2, 12) = 89.8
    myArray(2, 13) = 43
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 21.3
    myArray(3, 3) = 48.2
    myArray(3, 4) = 34
    myArray(3, 5) = 58
    myArray(3, 6) = 170.5
    myArray(3, 7) = 238.5
    myArray(3, 8) = 444.5
    myArray(3, 9) = 246.5
    myArray(3, 10) = 89
    myArray(3, 11) = 9
    myArray(3, 12) = 160
    myArray(3, 13) = 49
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 38.4
    myArray(4, 3) = 53.9
    myArray(4, 4) = 25.6
    myArray(4, 5) = 177.5
    myArray(4, 6) = 98.5
    myArray(4, 7) = 278.5
    myArray(4, 8) = 184
    myArray(4, 9) = 520
    myArray(4, 10) = 237.3
    myArray(4, 11) = 49
    myArray(4, 12) = 46.1
    myArray(4, 13) = 6.8
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 5.3
    myArray(5, 3) = 22.9
    myArray(5, 4) = 73
    myArray(5, 5) = 91.5
    myArray(5, 6) = 117.5
    myArray(5, 7) = 198
    myArray(5, 8) = 114.5
    myArray(5, 9) = 167.5
    myArray(5, 10) = 289.5
    myArray(5, 11) = 125
    myArray(5, 12) = 16.4
    myArray(5, 13) = 10.3
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 36.2
    myArray(6, 3) = 2.9
    myArray(6, 4) = 24.5
    myArray(6, 5) = 73.7
    myArray(6, 6) = 29
    myArray(6, 7) = 244.5
    myArray(6, 8) = 344
    myArray(6, 9) = 372
    myArray(6, 10) = 223
    myArray(6, 11) = 34.5
    myArray(6, 12) = 42
    myArray(6, 13) = 6.5
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 63.2
    myArray(7, 3) = 76.5
    myArray(7, 4) = 17
    myArray(7, 5) = 22.5
    myArray(7, 6) = 22.5
    myArray(7, 7) = 212.5
    myArray(7, 8) = 203
    myArray(7, 9) = 43
    myArray(7, 10) = 87
    myArray(7, 11) = 96
    myArray(7, 12) = 12
    myArray(7, 13) = 24.1
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 71.5
    myArray(8, 3) = 7.7
    myArray(8, 4) = 52
    myArray(8, 5) = 149.5
    myArray(8, 6) = 127.5
    myArray(8, 7) = 57
    myArray(8, 8) = 139.5
    myArray(8, 9) = 551
    myArray(8, 10) = 98.5
    myArray(8, 11) = 55.5
    myArray(8, 12) = 23.2
    myArray(8, 13) = 57.8
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 22.4
    myArray(9, 3) = 66
    myArray(9, 4) = 44
    myArray(9, 5) = 202.5
    myArray(9, 6) = 164
    myArray(9, 7) = 138
    myArray(9, 8) = 575
    myArray(9, 9) = 280.5
    myArray(9, 10) = 192
    myArray(9, 11) = 22.5
    myArray(9, 12) = 42.5
    myArray(9, 13) = 17
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 11.2
    myArray(10, 3) = 27.3
    myArray(10, 4) = 33
    myArray(10, 5) = 75.5
    myArray(10, 6) = 90.5
    myArray(10, 7) = 323.5
    myArray(10, 8) = 406
    myArray(10, 9) = 330.5
    myArray(10, 10) = 126
    myArray(10, 11) = 2.5
    myArray(10, 12) = 43
    myArray(10, 13) = 34.5
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 9.4
    myArray(11, 3) = 34
    myArray(11, 4) = 51
    myArray(11, 5) = 31.5
    myArray(11, 6) = 65.5
    myArray(11, 7) = 191
    myArray(11, 8) = 411.5
    myArray(11, 9) = 387
    myArray(11, 10) = 118
    myArray(11, 11) = 23
    myArray(11, 12) = 30.5
    myArray(11, 13) = 22.6
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 28
    myArray(12, 3) = 41.1
    myArray(12, 4) = 8.4
    myArray(12, 5) = 112
    myArray(12, 6) = 93.5
    myArray(12, 7) = 73
    myArray(12, 8) = 681.5
    myArray(12, 9) = 118
    myArray(12, 10) = 40.5
    myArray(12, 11) = 54
    myArray(12, 12) = 71
    myArray(12, 13) = 28.9
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 13.7
    myArray(13, 3) = 57
    myArray(13, 4) = 129
    myArray(13, 5) = 27.5
    myArray(13, 6) = 104
    myArray(13, 7) = 180
    myArray(13, 8) = 252
    myArray(13, 9) = 343.5
    myArray(13, 10) = 398.5
    myArray(13, 11) = 32
    myArray(13, 12) = 13.5
    myArray(13, 13) = 35.4
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 32.4
    myArray(14, 3) = 6.1
    myArray(14, 4) = 28.3
    myArray(14, 5) = 37.6
    myArray(14, 6) = 84.5
    myArray(14, 7) = 190.5
    myArray(14, 8) = 202
    myArray(14, 9) = 210
    myArray(14, 10) = 35.9
    myArray(14, 11) = 40.1
    myArray(14, 12) = 15.1
    myArray(14, 13) = 19.7
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 13.2
    myArray(15, 3) = 41.5
    myArray(15, 4) = 43
    myArray(15, 5) = 36
    myArray(15, 6) = 120.3
    myArray(15, 7) = 116.3
    myArray(15, 8) = 515.5
    myArray(15, 9) = 97
    myArray(15, 10) = 54.5
    myArray(15, 11) = 24
    myArray(15, 12) = 29
    myArray(15, 13) = 38.3
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 33.6
    myArray(16, 3) = 74.5
    myArray(16, 4) = 83.8
    myArray(16, 5) = 73.7
    myArray(16, 6) = 114.5
    myArray(16, 7) = 62.5
    myArray(16, 8) = 278.5
    myArray(16, 9) = 495.6
    myArray(16, 10) = 110.3
    myArray(16, 11) = 20.2
    myArray(16, 12) = 20
    myArray(16, 13) = 36.5
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 2.2
    myArray(17, 3) = 63.5
    myArray(17, 4) = 21.5
    myArray(17, 5) = 132.9
    myArray(17, 6) = 130.6
    myArray(17, 7) = 237.8
    myArray(17, 8) = 571.2
    myArray(17, 9) = 403
    myArray(17, 10) = 77.8
    myArray(17, 11) = 52.2
    myArray(17, 12) = 98
    myArray(17, 13) = 7.8
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 23.7
    myArray(18, 3) = 1.1
    myArray(18, 4) = 83.6
    myArray(18, 5) = 75.9
    myArray(18, 6) = 21.7
    myArray(18, 7) = 115.7
    myArray(18, 8) = 239.2
    myArray(18, 9) = 497.5
    myArray(18, 10) = 219.5
    myArray(18, 11) = 46.6
    myArray(18, 12) = 47.3
    myArray(18, 13) = 62.7
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 37
    myArray(19, 3) = 43.8
    myArray(19, 4) = 64.6
    myArray(19, 5) = 86.4
    myArray(19, 6) = 79.5
    myArray(19, 7) = 117.7
    myArray(19, 8) = 216.9
    myArray(19, 9) = 159.5
    myArray(19, 10) = 80.8
    myArray(19, 11) = 32.6
    myArray(19, 12) = 53.9
    myArray(19, 13) = 24.1
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 4.1
    myArray(20, 3) = 2.7
    myArray(20, 4) = 97.9
    myArray(20, 5) = 88.7
    myArray(20, 6) = 26
    myArray(20, 7) = 45.6
    myArray(20, 8) = 105.8
    myArray(20, 9) = 426.4
    myArray(20, 10) = 91.2
    myArray(20, 11) = 141.2
    myArray(20, 12) = 70.1
    myArray(20, 13) = 31.3
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 37.6
    myArray(21, 3) = 23.4
    myArray(21, 4) = 46.6
    myArray(21, 5) = 93.5
    myArray(21, 6) = 29.5
    myArray(21, 7) = 143.7
    myArray(21, 8) = 162.3
    myArray(21, 9) = 83.6
    myArray(21, 10) = 18.6
    myArray(21, 11) = 93.5
    myArray(21, 12) = 109.6
    myArray(21, 13) = 35.7
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 11.1
    myArray(22, 3) = 46
    myArray(22, 4) = 54.5
    myArray(22, 5) = 171.7
    myArray(22, 6) = 70.5
    myArray(22, 7) = 87.4
    myArray(22, 8) = 377.9
    myArray(22, 9) = 105.6
    myArray(22, 10) = 160.9
    myArray(22, 11) = 157.2
    myArray(22, 12) = 33.2
    myArray(22, 13) = 49.6
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 13.6
    myArray(23, 3) = 54.6
    myArray(23, 4) = 29.8
    myArray(23, 5) = 76.1
    myArray(23, 6) = 31.8
    myArray(23, 7) = 48.3
    myArray(23, 8) = 305.5
    myArray(23, 9) = 222.3
    myArray(23, 10) = 105.6
    myArray(23, 11) = 35.1
    myArray(23, 12) = 15.6
    myArray(23, 13) = 29.3
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 25.7
    myArray(24, 3) = 28.1
    myArray(24, 4) = 91.5
    myArray(24, 5) = 142.4
    myArray(24, 6) = 110.4
    myArray(24, 7) = 104.3
    myArray(24, 8) = 163.5
    myArray(24, 9) = 410.4
    myArray(24, 10) = 135.2
    myArray(24, 11) = 112.6
    myArray(24, 12) = 45.5
    myArray(24, 13) = 27.6
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 6.4
    myArray(25, 3) = 41.5
    myArray(25, 4) = 33
    myArray(25, 5) = 93
    myArray(25, 6) = 44.2
    myArray(25, 7) = 101
    myArray(25, 8) = 141.1
    myArray(25, 9) = 105.8
    myArray(25, 10) = 236.4
    myArray(25, 11) = 99.3
    myArray(25, 12) = 47.9
    myArray(25, 13) = 33
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 80.8
    myArray(26, 3) = 83.9
    myArray(26, 4) = 20.5
    myArray(26, 5) = 35.6
    myArray(26, 6) = 80.5
    myArray(26, 7) = 234
    myArray(26, 8) = 628
    myArray(26, 9) = 373.4
    myArray(26, 10) = 167.2
    myArray(26, 11) = 4.1
    myArray(26, 12) = 41.9
    myArray(26, 13) = 8.3
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 23.5
    myArray(27, 3) = 19.3
    myArray(27, 4) = 88
    myArray(27, 5) = 39.3
    myArray(27, 6) = 162.7
    myArray(27, 7) = 105.6
    myArray(27, 8) = 300.8
    myArray(27, 9) = 297.2
    myArray(27, 10) = 151.9
    myArray(27, 11) = 44
    myArray(27, 12) = 50.7
    myArray(27, 13) = 7.1
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 1.5
    myArray(28, 3) = 4.1
    myArray(28, 4) = 80.6
    myArray(28, 5) = 63.3
    myArray(28, 6) = 4.7
    myArray(28, 7) = 145.4
    myArray(28, 8) = 183.7
    myArray(28, 9) = 265.7
    myArray(28, 10) = 68.2
    myArray(28, 11) = 59.3
    myArray(28, 12) = 54.2
    myArray(28, 13) = 18
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 28.7
    myArray(29, 3) = 8.7
    myArray(29, 4) = 22
    myArray(29, 5) = 46.1
    myArray(29, 6) = 211.5
    myArray(29, 7) = 196.9
    myArray(29, 8) = 624.5
    myArray(29, 9) = 248.7
    myArray(29, 10) = 218.5
    myArray(29, 11) = 7.6
    myArray(29, 12) = 60
    myArray(29, 13) = 127.7
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 54.3
    myArray(30, 3) = 95.7
    myArray(30, 4) = 72.7
    myArray(30, 5) = 46.4
    myArray(30, 6) = 83.2
    myArray(30, 7) = 119.7
    myArray(30, 8) = 501.5
    myArray(30, 9) = 56.7
    myArray(30, 10) = 264.8
    myArray(30, 11) = 81
    myArray(30, 12) = 32.7
    myArray(30, 13) = 7.5
    
    data_GEUMSAN = myArray

End Function












Option Explicit

Option Explicit


Function data_GWANGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 42.3
    myArray(1, 3) = 34.9
    myArray(1, 4) = 28.1
    myArray(1, 5) = 111.7
    myArray(1, 6) = 75.5
    myArray(1, 7) = 96.4
    myArray(1, 8) = 110
    myArray(1, 9) = 151.4
    myArray(1, 10) = 40.5
    myArray(1, 11) = 17.2
    myArray(1, 12) = 35.1
    myArray(1, 13) = 21.3
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 32.9
    myArray(2, 3) = 11.8
    myArray(2, 4) = 127.4
    myArray(2, 5) = 38.4
    myArray(2, 6) = 37.4
    myArray(2, 7) = 302.9
    myArray(2, 8) = 186.3
    myArray(2, 9) = 261.7
    myArray(2, 10) = 66.1
    myArray(2, 11) = 60.7
    myArray(2, 12) = 112.4
    myArray(2, 13) = 30.8
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 19.3
    myArray(3, 3) = 43.6
    myArray(3, 4) = 68.3
    myArray(3, 5) = 82.1
    myArray(3, 6) = 101.6
    myArray(3, 7) = 177.2
    myArray(3, 8) = 358.3
    myArray(3, 9) = 381.9
    myArray(3, 10) = 22.8
    myArray(3, 11) = 14
    myArray(3, 12) = 136.4
    myArray(3, 13) = 73.7
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 41.3
    myArray(4, 3) = 44.4
    myArray(4, 4) = 78.2
    myArray(4, 5) = 124.2
    myArray(4, 6) = 136.9
    myArray(4, 7) = 369.6
    myArray(4, 8) = 210.9
    myArray(4, 9) = 531.2
    myArray(4, 10) = 315.7
    myArray(4, 11) = 57.4
    myArray(4, 12) = 30.9
    myArray(4, 13) = 2.3
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 41
    myArray(5, 3) = 34.9
    myArray(5, 4) = 111.9
    myArray(5, 5) = 44.9
    myArray(5, 6) = 106.9
    myArray(5, 7) = 135.4
    myArray(5, 8) = 220.7
    myArray(5, 9) = 287.1
    myArray(5, 10) = 281.7
    myArray(5, 11) = 126.2
    myArray(5, 12) = 22.1
    myArray(5, 13) = 16.9
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 33.7
    myArray(6, 3) = 15.4
    myArray(6, 4) = 19.9
    myArray(6, 5) = 16.3
    myArray(6, 6) = 47.7
    myArray(6, 7) = 197.7
    myArray(6, 8) = 321.4
    myArray(6, 9) = 538.3
    myArray(6, 10) = 228.7
    myArray(6, 11) = 31.7
    myArray(6, 12) = 53.8
    myArray(6, 13) = 6.4
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 50.4
    myArray(7, 3) = 83.9
    myArray(7, 4) = 26.9
    myArray(7, 5) = 43.7
    myArray(7, 6) = 21.4
    myArray(7, 7) = 245.8
    myArray(7, 8) = 270.8
    myArray(7, 9) = 113.2
    myArray(7, 10) = 135.3
    myArray(7, 11) = 69.7
    myArray(7, 12) = 15.3
    myArray(7, 13) = 53.5
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 95.7
    myArray(8, 3) = 13.5
    myArray(8, 4) = 46.5
    myArray(8, 5) = 117.8
    myArray(8, 6) = 100.5
    myArray(8, 7) = 108.5
    myArray(8, 8) = 180
    myArray(8, 9) = 584
    myArray(8, 10) = 109
    myArray(8, 11) = 38.6
    myArray(8, 12) = 34.1
    myArray(8, 13) = 30.5
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 45
    myArray(9, 3) = 52.2
    myArray(9, 4) = 70.6
    myArray(9, 5) = 243.9
    myArray(9, 6) = 141
    myArray(9, 7) = 150.2
    myArray(9, 8) = 564.9
    myArray(9, 9) = 411
    myArray(9, 10) = 206.3
    myArray(9, 11) = 36.2
    myArray(9, 12) = 38.4
    myArray(9, 13) = 34.4
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 14
    myArray(10, 3) = 54
    myArray(10, 4) = 38.4
    myArray(10, 5) = 63.7
    myArray(10, 6) = 101.1
    myArray(10, 7) = 153.8
    myArray(10, 8) = 409.5
    myArray(10, 9) = 570.5
    myArray(10, 10) = 242
    myArray(10, 11) = 3.4
    myArray(10, 12) = 63.1
    myArray(10, 13) = 28.8
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 10.6
    myArray(11, 3) = 48.3
    myArray(11, 4) = 66.7
    myArray(11, 5) = 92.5
    myArray(11, 6) = 74.1
    myArray(11, 7) = 185
    myArray(11, 8) = 273.8
    myArray(11, 9) = 303.3
    myArray(11, 10) = 108.5
    myArray(11, 11) = 17.4
    myArray(11, 12) = 42.8
    myArray(11, 13) = 66.6
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 16.1
    myArray(12, 3) = 51.9
    myArray(12, 4) = 15
    myArray(12, 5) = 87.8
    myArray(12, 6) = 204.3
    myArray(12, 7) = 226.7
    myArray(12, 8) = 478.3
    myArray(12, 9) = 295.5
    myArray(12, 10) = 46.8
    myArray(12, 11) = 18.2
    myArray(12, 12) = 41.7
    myArray(12, 13) = 37.9
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 12.1
    myArray(13, 3) = 79.1
    myArray(13, 4) = 100.3
    myArray(13, 5) = 38.7
    myArray(13, 6) = 116
    myArray(13, 7) = 52
    myArray(13, 8) = 232
    myArray(13, 9) = 339.3
    myArray(13, 10) = 490.7
    myArray(13, 11) = 95.5
    myArray(13, 12) = 3.3
    myArray(13, 13) = 61.6
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 56.9
    myArray(14, 3) = 11.1
    myArray(14, 4) = 36.7
    myArray(14, 5) = 54.2
    myArray(14, 6) = 150.6
    myArray(14, 7) = 273.2
    myArray(14, 8) = 139.2
    myArray(14, 9) = 157.5
    myArray(14, 10) = 58.9
    myArray(14, 11) = 15.3
    myArray(14, 12) = 39.1
    myArray(14, 13) = 14.5
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 36
    myArray(15, 3) = 71.9
    myArray(15, 4) = 42.1
    myArray(15, 5) = 35.7
    myArray(15, 6) = 114.9
    myArray(15, 7) = 181.1
    myArray(15, 8) = 607.4
    myArray(15, 9) = 263.1
    myArray(15, 10) = 22.6
    myArray(15, 11) = 36.2
    myArray(15, 12) = 26.7
    myArray(15, 13) = 50.5
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 37.1
    myArray(16, 3) = 135.8
    myArray(16, 4) = 76
    myArray(16, 5) = 133
    myArray(16, 6) = 99
    myArray(16, 7) = 70.6
    myArray(16, 8) = 453
    myArray(16, 9) = 337.6
    myArray(16, 10) = 139.7
    myArray(16, 11) = 42
    myArray(16, 12) = 7.4
    myArray(16, 13) = 41.9
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 7.4
    myArray(17, 3) = 59.8
    myArray(17, 4) = 23.1
    myArray(17, 5) = 103
    myArray(17, 6) = 142.9
    myArray(17, 7) = 120
    myArray(17, 8) = 277.5
    myArray(17, 9) = 382.5
    myArray(17, 10) = 13.5
    myArray(17, 11) = 20.5
    myArray(17, 12) = 136.8
    myArray(17, 13) = 13.3
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 15.2
    myArray(18, 3) = 18.6
    myArray(18, 4) = 100.4
    myArray(18, 5) = 82.5
    myArray(18, 6) = 42.6
    myArray(18, 7) = 83.1
    myArray(18, 8) = 330.6
    myArray(18, 9) = 473.5
    myArray(18, 10) = 272
    myArray(18, 11) = 82.8
    myArray(18, 12) = 45.9
    myArray(18, 13) = 79.6
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 20.6
    myArray(19, 3) = 48
    myArray(19, 4) = 76.9
    myArray(19, 5) = 54.9
    myArray(19, 6) = 86.5
    myArray(19, 7) = 83.7
    myArray(19, 8) = 349.1
    myArray(19, 9) = 293.2
    myArray(19, 10) = 88.5
    myArray(19, 11) = 30.8
    myArray(19, 12) = 95
    myArray(19, 13) = 18.2
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 12.9
    myArray(20, 3) = 8.6
    myArray(20, 4) = 101.7
    myArray(20, 5) = 62.5
    myArray(20, 6) = 57
    myArray(20, 7) = 72
    myArray(20, 8) = 240.9
    myArray(20, 9) = 370.2
    myArray(20, 10) = 116.5
    myArray(20, 11) = 105
    myArray(20, 12) = 95.5
    myArray(20, 13) = 47.5
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 47.2
    myArray(21, 3) = 23.9
    myArray(21, 4) = 36.5
    myArray(21, 5) = 145.5
    myArray(21, 6) = 48.6
    myArray(21, 7) = 96.1
    myArray(21, 8) = 164.3
    myArray(21, 9) = 148.9
    myArray(21, 10) = 66.6
    myArray(21, 11) = 90.9
    myArray(21, 12) = 121.9
    myArray(21, 13) = 59.2
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 49.7
    myArray(22, 3) = 45.2
    myArray(22, 4) = 55.2
    myArray(22, 5) = 185
    myArray(22, 6) = 104.5
    myArray(22, 7) = 116.1
    myArray(22, 8) = 301.3
    myArray(22, 9) = 81
    myArray(22, 10) = 251.2
    myArray(22, 11) = 216.7
    myArray(22, 12) = 31.5
    myArray(22, 13) = 44.9
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 11.7
    myArray(23, 3) = 41.9
    myArray(23, 4) = 33.2
    myArray(23, 5) = 60.6
    myArray(23, 6) = 30.2
    myArray(23, 7) = 42.1
    myArray(23, 8) = 211.6
    myArray(23, 9) = 280.5
    myArray(23, 10) = 108.8
    myArray(23, 11) = 85.4
    myArray(23, 12) = 2.1
    myArray(23, 13) = 28.5
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 38.8
    myArray(24, 3) = 22
    myArray(24, 4) = 115.8
    myArray(24, 5) = 127.4
    myArray(24, 6) = 85.4
    myArray(24, 7) = 222.4
    myArray(24, 8) = 84.5
    myArray(24, 9) = 397.1
    myArray(24, 10) = 129.7
    myArray(24, 11) = 125.2
    myArray(24, 12) = 47.2
    myArray(24, 13) = 32.4
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 16.4
    myArray(25, 3) = 37
    myArray(25, 4) = 33.9
    myArray(25, 5) = 84.7
    myArray(25, 6) = 78.8
    myArray(25, 7) = 158
    myArray(25, 8) = 242.2
    myArray(25, 9) = 64.8
    myArray(25, 10) = 165.8
    myArray(25, 11) = 149.9
    myArray(25, 12) = 22.8
    myArray(25, 13) = 31.6
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 74.9
    myArray(26, 3) = 47.8
    myArray(26, 4) = 43.5
    myArray(26, 5) = 55.3
    myArray(26, 6) = 96.8
    myArray(26, 7) = 199.9
    myArray(26, 8) = 533.3
    myArray(26, 9) = 738.1
    myArray(26, 10) = 178.3
    myArray(26, 11) = 12.1
    myArray(26, 12) = 28.3
    myArray(26, 13) = 18.7
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 33
    myArray(27, 3) = 31.1
    myArray(27, 4) = 122.4
    myArray(27, 5) = 34.2
    myArray(27, 6) = 139.4
    myArray(27, 7) = 118.1
    myArray(27, 8) = 227.6
    myArray(27, 9) = 338.7
    myArray(27, 10) = 131.1
    myArray(27, 11) = 35.3
    myArray(27, 12) = 85.8
    myArray(27, 13) = 7.1
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 0.3
    myArray(28, 3) = 2.7
    myArray(28, 4) = 105.2
    myArray(28, 5) = 49.4
    myArray(28, 6) = 0.4
    myArray(28, 7) = 131.7
    myArray(28, 8) = 169
    myArray(28, 9) = 106.4
    myArray(28, 10) = 89.2
    myArray(28, 11) = 38.3
    myArray(28, 12) = 46.7
    myArray(28, 13) = 30.6
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 26.8
    myArray(29, 3) = 14.8
    myArray(29, 4) = 34.8
    myArray(29, 5) = 66.5
    myArray(29, 6) = 190.1
    myArray(29, 7) = 441.2
    myArray(29, 8) = 684.6
    myArray(29, 9) = 341.2
    myArray(29, 10) = 152.5
    myArray(29, 11) = 11.5
    myArray(29, 12) = 71.9
    myArray(29, 13) = 80.2
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 31.1
    myArray(30, 3) = 111.3
    myArray(30, 4) = 79.5
    myArray(30, 5) = 79.6
    myArray(30, 6) = 102
    myArray(30, 7) = 149.1
    myArray(30, 8) = 250.4
    myArray(30, 9) = 132.1
    myArray(30, 10) = 181.4
    myArray(30, 11) = 98.2
    myArray(30, 12) = 66.5
    myArray(30, 13) = 9.8


    data_GWANGJU = myArray

End Function



Function data_SEOUL() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 11.6
    myArray(1, 3) = 5.2
    myArray(1, 4) = 60.6
    myArray(1, 5) = 44.4
    myArray(1, 6) = 60.6
    myArray(1, 7) = 70.7
    myArray(1, 8) = 436.1
    myArray(1, 9) = 786.6
    myArray(1, 10) = 47.2
    myArray(1, 11) = 39.3
    myArray(1, 12) = 32.9
    myArray(1, 13) = 3.4
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 16.3
    myArray(2, 3) = 1
    myArray(2, 4) = 77.9
    myArray(2, 5) = 62
    myArray(2, 6) = 29.3
    myArray(2, 7) = 249.7
    myArray(2, 8) = 512.8
    myArray(2, 9) = 132.4
    myArray(2, 10) = 11
    myArray(2, 11) = 90.3
    myArray(2, 12) = 62.9
    myArray(2, 13) = 11
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 16.8
    myArray(3, 3) = 39.6
    myArray(3, 4) = 25.3
    myArray(3, 5) = 56.1
    myArray(3, 6) = 291.3
    myArray(3, 7) = 110
    myArray(3, 8) = 299.6
    myArray(3, 9) = 117.2
    myArray(3, 10) = 76.9
    myArray(3, 11) = 45.5
    myArray(3, 12) = 93.8
    myArray(3, 13) = 38.1
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 10.4
    myArray(4, 3) = 32.3
    myArray(4, 4) = 45.1
    myArray(4, 5) = 120.2
    myArray(4, 6) = 121.5
    myArray(4, 7) = 234.1
    myArray(4, 8) = 311.8
    myArray(4, 9) = 1237.8
    myArray(4, 10) = 177.9
    myArray(4, 11) = 27.4
    myArray(4, 12) = 26.9
    myArray(4, 13) = 3.7
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 10.2
    myArray(5, 3) = 2.9
    myArray(5, 4) = 55
    myArray(5, 5) = 97.2
    myArray(5, 6) = 109.7
    myArray(5, 7) = 131.8
    myArray(5, 8) = 230.4
    myArray(5, 9) = 600.5
    myArray(5, 10) = 377.3
    myArray(5, 11) = 81.6
    myArray(5, 12) = 19.5
    myArray(5, 13) = 17
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 42.8
    myArray(6, 3) = 2.1
    myArray(6, 4) = 3.1
    myArray(6, 5) = 30.7
    myArray(6, 6) = 75.2
    myArray(6, 7) = 68.1
    myArray(6, 8) = 114.7
    myArray(6, 9) = 599.4
    myArray(6, 10) = 178.5
    myArray(6, 11) = 18.1
    myArray(6, 12) = 27.1
    myArray(6, 13) = 27
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 39.4
    myArray(7, 3) = 45.7
    myArray(7, 4) = 18.1
    myArray(7, 5) = 12.3
    myArray(7, 6) = 16.5
    myArray(7, 7) = 157.4
    myArray(7, 8) = 698.4
    myArray(7, 9) = 252
    myArray(7, 10) = 49.3
    myArray(7, 11) = 68.2
    myArray(7, 12) = 13
    myArray(7, 13) = 15.7
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 37.4
    myArray(8, 3) = 2.4
    myArray(8, 4) = 31.5
    myArray(8, 5) = 155.1
    myArray(8, 6) = 58
    myArray(8, 7) = 61.4
    myArray(8, 8) = 220.6
    myArray(8, 9) = 688
    myArray(8, 10) = 61.1
    myArray(8, 11) = 45
    myArray(8, 12) = 12.5
    myArray(8, 13) = 15
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 14.1
    myArray(9, 3) = 39.6
    myArray(9, 4) = 26.8
    myArray(9, 5) = 139.6
    myArray(9, 6) = 106
    myArray(9, 7) = 156
    myArray(9, 8) = 469.8
    myArray(9, 9) = 684.2
    myArray(9, 10) = 258.2
    myArray(9, 11) = 41.5
    myArray(9, 12) = 69.3
    myArray(9, 13) = 6.9
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 19.8
    myArray(10, 3) = 54.6
    myArray(10, 4) = 27.6
    myArray(10, 5) = 74.1
    myArray(10, 6) = 168.5
    myArray(10, 7) = 138.1
    myArray(10, 8) = 510.7
    myArray(10, 9) = 193.3
    myArray(10, 10) = 198.7
    myArray(10, 11) = 6.5
    myArray(10, 12) = 80
    myArray(10, 13) = 27.2
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 4.5
    myArray(11, 3) = 17.2
    myArray(11, 4) = 12.5
    myArray(11, 5) = 94.7
    myArray(11, 6) = 85.8
    myArray(11, 7) = 168.5
    myArray(11, 8) = 269.4
    myArray(11, 9) = 285
    myArray(11, 10) = 313.3
    myArray(11, 11) = 52.6
    myArray(11, 12) = 44.6
    myArray(11, 13) = 10.3
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 34.3
    myArray(12, 3) = 15.7
    myArray(12, 4) = 14
    myArray(12, 5) = 51.8
    myArray(12, 6) = 156.2
    myArray(12, 7) = 168.5
    myArray(12, 8) = 1014
    myArray(12, 9) = 121.2
    myArray(12, 10) = 11.1
    myArray(12, 11) = 30.2
    myArray(12, 12) = 47.6
    myArray(12, 13) = 17.3
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 10.8
    myArray(13, 3) = 12.6
    myArray(13, 4) = 123.5
    myArray(13, 5) = 41.1
    myArray(13, 6) = 137.6
    myArray(13, 7) = 54.5
    myArray(13, 8) = 274.1
    myArray(13, 9) = 237.6
    myArray(13, 10) = 241.9
    myArray(13, 11) = 39.5
    myArray(13, 12) = 26.4
    myArray(13, 13) = 12.7
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 17.7
    myArray(14, 3) = 15
    myArray(14, 4) = 53.9
    myArray(14, 5) = 38.5
    myArray(14, 6) = 97.7
    myArray(14, 7) = 165
    myArray(14, 8) = 530.8
    myArray(14, 9) = 251.2
    myArray(14, 10) = 99.2
    myArray(14, 11) = 41.8
    myArray(14, 12) = 19.6
    myArray(14, 13) = 25.9
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 5.7
    myArray(15, 3) = 36.9
    myArray(15, 4) = 63.9
    myArray(15, 5) = 66.5
    myArray(15, 6) = 109
    myArray(15, 7) = 132
    myArray(15, 8) = 659.4
    myArray(15, 9) = 285.3
    myArray(15, 10) = 64.5
    myArray(15, 11) = 66.9
    myArray(15, 12) = 52.4
    myArray(15, 13) = 21.5
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 29.3
    myArray(16, 3) = 55.3
    myArray(16, 4) = 82.5
    myArray(16, 5) = 62.8
    myArray(16, 6) = 124
    myArray(16, 7) = 127.6
    myArray(16, 8) = 239.2
    myArray(16, 9) = 598.7
    myArray(16, 10) = 671.5
    myArray(16, 11) = 25.6
    myArray(16, 12) = 10.9
    myArray(16, 13) = 16.1
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 8.9
    myArray(17, 3) = 29.1
    myArray(17, 4) = 14.6
    myArray(17, 5) = 110.1
    myArray(17, 6) = 53.4
    myArray(17, 7) = 404.5
    myArray(17, 8) = 1131
    myArray(17, 9) = 166.8
    myArray(17, 10) = 25.6
    myArray(17, 11) = 32
    myArray(17, 12) = 56.2
    myArray(17, 13) = 7.1
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 6.7
    myArray(18, 3) = 0.8
    myArray(18, 4) = 47.4
    myArray(18, 5) = 157
    myArray(18, 6) = 8.2
    myArray(18, 7) = 91.9
    myArray(18, 8) = 448.9
    myArray(18, 9) = 464.9
    myArray(18, 10) = 212
    myArray(18, 11) = 99.3
    myArray(18, 12) = 67.8
    myArray(18, 13) = 41.4
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 22.1
    myArray(19, 3) = 74.1
    myArray(19, 4) = 27.3
    myArray(19, 5) = 71.7
    myArray(19, 6) = 132
    myArray(19, 7) = 28.3
    myArray(19, 8) = 676.2
    myArray(19, 9) = 148.6
    myArray(19, 10) = 138.5
    myArray(19, 11) = 13.5
    myArray(19, 12) = 46.8
    myArray(19, 13) = 24.7
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 13
    myArray(20, 3) = 16.2
    myArray(20, 4) = 7.2
    myArray(20, 5) = 31
    myArray(20, 6) = 63
    myArray(20, 7) = 98.1
    myArray(20, 8) = 207.9
    myArray(20, 9) = 172.8
    myArray(20, 10) = 88.1
    myArray(20, 11) = 52.2
    myArray(20, 12) = 41.5
    myArray(20, 13) = 17.9
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 11.3
    myArray(21, 3) = 22.7
    myArray(21, 4) = 9.6
    myArray(21, 5) = 80.5
    myArray(21, 6) = 28.9
    myArray(21, 7) = 99
    myArray(21, 8) = 226
    myArray(21, 9) = 72.9
    myArray(21, 10) = 26
    myArray(21, 11) = 81.5
    myArray(21, 12) = 104.6
    myArray(21, 13) = 29.1
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 1
    myArray(22, 3) = 47.6
    myArray(22, 4) = 40.5
    myArray(22, 5) = 76.8
    myArray(22, 6) = 160.5
    myArray(22, 7) = 54.4
    myArray(22, 8) = 358.2
    myArray(22, 9) = 67.1
    myArray(22, 10) = 33
    myArray(22, 11) = 74.8
    myArray(22, 12) = 16.7
    myArray(22, 13) = 61.1
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 14.9
    myArray(23, 3) = 11.1
    myArray(23, 4) = 7.9
    myArray(23, 5) = 61.6
    myArray(23, 6) = 16.1
    myArray(23, 7) = 66.6
    myArray(23, 8) = 621
    myArray(23, 9) = 297
    myArray(23, 10) = 35
    myArray(23, 11) = 26.5
    myArray(23, 12) = 40.7
    myArray(23, 13) = 34.8
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 8.5
    myArray(24, 3) = 29.6
    myArray(24, 4) = 49.5
    myArray(24, 5) = 130.3
    myArray(24, 6) = 222
    myArray(24, 7) = 171.5
    myArray(24, 8) = 185.6
    myArray(24, 9) = 202.6
    myArray(24, 10) = 68.5
    myArray(24, 11) = 120.5
    myArray(24, 12) = 79.1
    myArray(24, 13) = 16.4
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0
    myArray(25, 3) = 23.8
    myArray(25, 4) = 26.8
    myArray(25, 5) = 47.3
    myArray(25, 6) = 37.8
    myArray(25, 7) = 74
    myArray(25, 8) = 194.4
    myArray(25, 9) = 190.5
    myArray(25, 10) = 139.8
    myArray(25, 11) = 55.5
    myArray(25, 12) = 78.8
    myArray(25, 13) = 22.6
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 60.5
    myArray(26, 3) = 53.1
    myArray(26, 4) = 16.3
    myArray(26, 5) = 16.9
    myArray(26, 6) = 112.4
    myArray(26, 7) = 139.6
    myArray(26, 8) = 270.4
    myArray(26, 9) = 675.7
    myArray(26, 10) = 181.5
    myArray(26, 11) = 0
    myArray(26, 12) = 120.1
    myArray(26, 13) = 4.6
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 18.9
    myArray(27, 3) = 7.1
    myArray(27, 4) = 110.9
    myArray(27, 5) = 124.1
    myArray(27, 6) = 183.1
    myArray(27, 7) = 104.6
    myArray(27, 8) = 168.3
    myArray(27, 9) = 211.2
    myArray(27, 10) = 131
    myArray(27, 11) = 57
    myArray(27, 12) = 62.4
    myArray(27, 13) = 7.9
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 5.5
    myArray(28, 3) = 4.7
    myArray(28, 4) = 102.6
    myArray(28, 5) = 20.4
    myArray(28, 6) = 7.5
    myArray(28, 7) = 393.8
    myArray(28, 8) = 252.3
    myArray(28, 9) = 564.8
    myArray(28, 10) = 201.5
    myArray(28, 11) = 124.1
    myArray(28, 12) = 84.5
    myArray(28, 13) = 13.6
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 47.9
    myArray(29, 3) = 1
    myArray(29, 4) = 10.5
    myArray(29, 5) = 96.9
    myArray(29, 6) = 155.6
    myArray(29, 7) = 195.6
    myArray(29, 8) = 459.9
    myArray(29, 9) = 298.1
    myArray(29, 10) = 134.5
    myArray(29, 11) = 31
    myArray(29, 12) = 81.9
    myArray(29, 13) = 85.9
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 18.9
    myArray(30, 3) = 74.7
    myArray(30, 4) = 29.9
    myArray(30, 5) = 33.2
    myArray(30, 6) = 125.1
    myArray(30, 7) = 115.9
    myArray(30, 8) = 557.3
    myArray(30, 9) = 72.8
    myArray(30, 10) = 143.9
    myArray(30, 11) = 74
    myArray(30, 12) = 60
    myArray(30, 13) = 5.7
    

    data_SEOUL = myArray

End Function


Function data_SUWON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1995
    myArray(1, 2) = 13.4
    myArray(1, 3) = 11.2
    myArray(1, 4) = 46.2
    myArray(1, 5) = 33.7
    myArray(1, 6) = 59
    myArray(1, 7) = 67.7
    myArray(1, 8) = 372.9
    myArray(1, 9) = 967.9
    myArray(1, 10) = 24.2
    myArray(1, 11) = 29.2
    myArray(1, 12) = 24.8
    myArray(1, 13) = 3.1
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 20.4
    myArray(2, 3) = 4.1
    myArray(2, 4) = 100.8
    myArray(2, 5) = 51.1
    myArray(2, 6) = 26.5
    myArray(2, 7) = 286.4
    myArray(2, 8) = 241.1
    myArray(2, 9) = 77.5
    myArray(2, 10) = 9.2
    myArray(2, 11) = 70
    myArray(2, 12) = 49
    myArray(2, 13) = 16
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 14.4
    myArray(3, 3) = 41.4
    myArray(3, 4) = 30.4
    myArray(3, 5) = 60.7
    myArray(3, 6) = 260.3
    myArray(3, 7) = 150.4
    myArray(3, 8) = 331.7
    myArray(3, 9) = 299.2
    myArray(3, 10) = 25
    myArray(3, 11) = 52.3
    myArray(3, 12) = 82
    myArray(3, 13) = 46.5
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 23.7
    myArray(4, 3) = 34.3
    myArray(4, 4) = 44
    myArray(4, 5) = 105.9
    myArray(4, 6) = 86.4
    myArray(4, 7) = 213.7
    myArray(4, 8) = 306
    myArray(4, 9) = 591.6
    myArray(4, 10) = 141.2
    myArray(4, 11) = 25
    myArray(4, 12) = 51.6
    myArray(4, 13) = 3.5
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 7.3
    myArray(5, 3) = 1.8
    myArray(5, 4) = 54
    myArray(5, 5) = 73.6
    myArray(5, 6) = 121.3
    myArray(5, 7) = 76.7
    myArray(5, 8) = 345
    myArray(5, 9) = 338.4
    myArray(5, 10) = 402.2
    myArray(5, 11) = 92.3
    myArray(5, 12) = 25.3
    myArray(5, 13) = 18.2
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 57.7
    myArray(6, 3) = 1.4
    myArray(6, 4) = 3.1
    myArray(6, 5) = 20.4
    myArray(6, 6) = 43.7
    myArray(6, 7) = 118.2
    myArray(6, 8) = 375.8
    myArray(6, 9) = 448.8
    myArray(6, 10) = 182.2
    myArray(6, 11) = 21.6
    myArray(6, 12) = 27.5
    myArray(6, 13) = 28.4
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 42.2
    myArray(7, 3) = 47.4
    myArray(7, 4) = 15.1
    myArray(7, 5) = 12.9
    myArray(7, 6) = 13.8
    myArray(7, 7) = 222.3
    myArray(7, 8) = 469.7
    myArray(7, 9) = 144.7
    myArray(7, 10) = 12.1
    myArray(7, 11) = 58.1
    myArray(7, 12) = 14.2
    myArray(7, 13) = 14.7
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 31.8
    myArray(8, 3) = 3.2
    myArray(8, 4) = 35.7
    myArray(8, 5) = 152.4
    myArray(8, 6) = 77
    myArray(8, 7) = 52
    myArray(8, 8) = 257.8
    myArray(8, 9) = 487.3
    myArray(8, 10) = 31.3
    myArray(8, 11) = 73.8
    myArray(8, 12) = 12.2
    myArray(8, 13) = 17.2
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 10.4
    myArray(9, 3) = 46.2
    myArray(9, 4) = 28.3
    myArray(9, 5) = 182
    myArray(9, 6) = 85.5
    myArray(9, 7) = 159
    myArray(9, 8) = 341.9
    myArray(9, 9) = 293.7
    myArray(9, 10) = 271.5
    myArray(9, 11) = 30.6
    myArray(9, 12) = 51.6
    myArray(9, 13) = 14.1
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 17.8
    myArray(10, 3) = 42.6
    myArray(10, 4) = 14.4
    myArray(10, 5) = 63.8
    myArray(10, 6) = 125.2
    myArray(10, 7) = 135.7
    myArray(10, 8) = 382
    myArray(10, 9) = 157.4
    myArray(10, 10) = 183.4
    myArray(10, 11) = 2
    myArray(10, 12) = 67.5
    myArray(10, 13) = 25.2
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 5.7
    myArray(11, 3) = 15
    myArray(11, 4) = 25.6
    myArray(11, 5) = 85.6
    myArray(11, 6) = 89.6
    myArray(11, 7) = 160.8
    myArray(11, 8) = 251.7
    myArray(11, 9) = 357.5
    myArray(11, 10) = 315.2
    myArray(11, 11) = 70.2
    myArray(11, 12) = 38.8
    myArray(11, 13) = 12
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 38.6
    myArray(12, 3) = 19.5
    myArray(12, 4) = 6.9
    myArray(12, 5) = 59.9
    myArray(12, 6) = 133.2
    myArray(12, 7) = 156.7
    myArray(12, 8) = 754.7
    myArray(12, 9) = 66.4
    myArray(12, 10) = 21.9
    myArray(12, 11) = 18
    myArray(12, 12) = 61.6
    myArray(12, 13) = 25.3
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 9.3
    myArray(13, 3) = 15.1
    myArray(13, 4) = 135.3
    myArray(13, 5) = 24.2
    myArray(13, 6) = 146.7
    myArray(13, 7) = 74.2
    myArray(13, 8) = 269.7
    myArray(13, 9) = 295
    myArray(13, 10) = 268.8
    myArray(13, 11) = 18.3
    myArray(13, 12) = 57.1
    myArray(13, 13) = 11.3
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 13.6
    myArray(14, 3) = 8.7
    myArray(14, 4) = 55.9
    myArray(14, 5) = 42.2
    myArray(14, 6) = 92.7
    myArray(14, 7) = 198.4
    myArray(14, 8) = 540.8
    myArray(14, 9) = 217.2
    myArray(14, 10) = 101.9
    myArray(14, 11) = 35.6
    myArray(14, 12) = 18.5
    myArray(14, 13) = 17.4
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 7.9
    myArray(15, 3) = 26.8
    myArray(15, 4) = 59.5
    myArray(15, 5) = 45
    myArray(15, 6) = 102.4
    myArray(15, 7) = 118.8
    myArray(15, 8) = 766
    myArray(15, 9) = 207.1
    myArray(15, 10) = 56.3
    myArray(15, 11) = 64.5
    myArray(15, 12) = 68.2
    myArray(15, 13) = 18.7
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 26.9
    myArray(16, 3) = 56.7
    myArray(16, 4) = 78.7
    myArray(16, 5) = 58.6
    myArray(16, 6) = 100.7
    myArray(16, 7) = 116.1
    myArray(16, 8) = 206.8
    myArray(16, 9) = 372.8
    myArray(16, 10) = 375.9
    myArray(16, 11) = 30
    myArray(16, 12) = 18.1
    myArray(16, 13) = 29.3
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 11.3
    myArray(17, 3) = 49.8
    myArray(17, 4) = 23.4
    myArray(17, 5) = 186.4
    myArray(17, 6) = 74.2
    myArray(17, 7) = 391.5
    myArray(17, 8) = 794.3
    myArray(17, 9) = 315.1
    myArray(17, 10) = 32.8
    myArray(17, 11) = 38.4
    myArray(17, 12) = 46.3
    myArray(17, 13) = 12.4
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 9.9
    myArray(18, 3) = 0.7
    myArray(18, 4) = 43.1
    myArray(18, 5) = 125.5
    myArray(18, 6) = 16.5
    myArray(18, 7) = 100.8
    myArray(18, 8) = 572.3
    myArray(18, 9) = 426.2
    myArray(18, 10) = 241
    myArray(18, 11) = 98.6
    myArray(18, 12) = 66.5
    myArray(18, 13) = 47.2
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 23.1
    myArray(19, 3) = 53.4
    myArray(19, 4) = 48.2
    myArray(19, 5) = 69.5
    myArray(19, 6) = 129
    myArray(19, 7) = 69.6
    myArray(19, 8) = 405.9
    myArray(19, 9) = 157
    myArray(19, 10) = 183.6
    myArray(19, 11) = 5.4
    myArray(19, 12) = 61.9
    myArray(19, 13) = 33.5
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 9.9
    myArray(20, 3) = 16.5
    myArray(20, 4) = 10.9
    myArray(20, 5) = 55.7
    myArray(20, 6) = 64.4
    myArray(20, 7) = 68.1
    myArray(20, 8) = 264
    myArray(20, 9) = 290.9
    myArray(20, 10) = 92
    myArray(20, 11) = 85.4
    myArray(20, 12) = 44.2
    myArray(20, 13) = 27.1
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 17.5
    myArray(21, 3) = 22.7
    myArray(21, 4) = 12.5
    myArray(21, 5) = 99.2
    myArray(21, 6) = 32.6
    myArray(21, 7) = 30.2
    myArray(21, 8) = 225.8
    myArray(21, 9) = 71
    myArray(21, 10) = 6.9
    myArray(21, 11) = 67.4
    myArray(21, 12) = 116
    myArray(21, 13) = 49.3
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 4.6
    myArray(22, 3) = 52.6
    myArray(22, 4) = 54.8
    myArray(22, 5) = 79.2
    myArray(22, 6) = 156.4
    myArray(22, 7) = 37.4
    myArray(22, 8) = 317.7
    myArray(22, 9) = 73
    myArray(22, 10) = 67.8
    myArray(22, 11) = 99.1
    myArray(22, 12) = 17.4
    myArray(22, 13) = 63.4
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 19.4
    myArray(23, 3) = 22.5
    myArray(23, 4) = 9
    myArray(23, 5) = 52.5
    myArray(23, 6) = 22.5
    myArray(23, 7) = 27.6
    myArray(23, 8) = 684.5
    myArray(23, 9) = 359.7
    myArray(23, 10) = 26.1
    myArray(23, 11) = 28.7
    myArray(23, 12) = 37.6
    myArray(23, 13) = 38.5
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 7.9
    myArray(24, 3) = 29.1
    myArray(24, 4) = 86
    myArray(24, 5) = 128.8
    myArray(24, 6) = 196.4
    myArray(24, 7) = 107
    myArray(24, 8) = 222.7
    myArray(24, 9) = 218.6
    myArray(24, 10) = 61.7
    myArray(24, 11) = 132.7
    myArray(24, 12) = 78.1
    myArray(24, 13) = 24.1
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.5
    myArray(25, 3) = 33.4
    myArray(25, 4) = 39.6
    myArray(25, 5) = 43.6
    myArray(25, 6) = 26.7
    myArray(25, 7) = 68.8
    myArray(25, 8) = 190.9
    myArray(25, 9) = 117.9
    myArray(25, 10) = 201.9
    myArray(25, 11) = 73.1
    myArray(25, 12) = 93.2
    myArray(25, 13) = 26.2
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 61
    myArray(26, 3) = 59.2
    myArray(26, 4) = 19.5
    myArray(26, 5) = 15.8
    myArray(26, 6) = 97.9
    myArray(26, 7) = 91.1
    myArray(26, 8) = 384.7
    myArray(26, 9) = 659.3
    myArray(26, 10) = 163.4
    myArray(26, 11) = 22.2
    myArray(26, 12) = 56.8
    myArray(26, 13) = 4.6
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 24.1
    myArray(27, 3) = 9.7
    myArray(27, 4) = 102.8
    myArray(27, 5) = 106.4
    myArray(27, 6) = 179.2
    myArray(27, 7) = 50.1
    myArray(27, 8) = 134.4
    myArray(27, 9) = 161.1
    myArray(27, 10) = 164.1
    myArray(27, 11) = 36
    myArray(27, 12) = 104.2
    myArray(27, 13) = 12.3
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 3.9
    myArray(28, 3) = 4
    myArray(28, 4) = 74.7
    myArray(28, 5) = 42.4
    myArray(28, 6) = 13.4
    myArray(28, 7) = 472.1
    myArray(28, 8) = 243
    myArray(28, 9) = 640.8
    myArray(28, 10) = 190.1
    myArray(28, 11) = 108.5
    myArray(28, 12) = 51.4
    myArray(28, 13) = 19.7
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 38.9
    myArray(29, 3) = 0.7
    myArray(29, 4) = 5.5
    myArray(29, 5) = 67.4
    myArray(29, 6) = 120.3
    myArray(29, 7) = 131.5
    myArray(29, 8) = 490.5
    myArray(29, 9) = 200.7
    myArray(29, 10) = 135.6
    myArray(29, 11) = 35.8
    myArray(29, 12) = 99
    myArray(29, 13) = 83.5
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 21.2
    myArray(30, 3) = 74.8
    myArray(30, 4) = 30.2
    myArray(30, 5) = 55
    myArray(30, 6) = 109.4
    myArray(30, 7) = 137.5
    myArray(30, 8) = 403.3
    myArray(30, 9) = 127.6
    myArray(30, 10) = 219.4
    myArray(30, 11) = 97.2
    myArray(30, 12) = 98.1
    myArray(30, 13) = 4.5


    data_SUWON = myArray

End Function


Function data_INCHEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    myArray(1, 1) = 1995
    myArray(1, 2) = 14.2
    myArray(1, 3) = 2.3
    myArray(1, 4) = 49.1
    myArray(1, 5) = 35.7
    myArray(1, 6) = 43.6
    myArray(1, 7) = 83.4
    myArray(1, 8) = 367.8
    myArray(1, 9) = 621.4
    myArray(1, 10) = 57.2
    myArray(1, 11) = 27
    myArray(1, 12) = 23.2
    myArray(1, 13) = 1.3
    
    myArray(2, 1) = 1996
    myArray(2, 2) = 11.1
    myArray(2, 3) = 0.7
    myArray(2, 4) = 85.4
    myArray(2, 5) = 64
    myArray(2, 6) = 19.9
    myArray(2, 7) = 236.8
    myArray(2, 8) = 276.1
    myArray(2, 9) = 69.6
    myArray(2, 10) = 7.5
    myArray(2, 11) = 76.8
    myArray(2, 12) = 70.7
    myArray(2, 13) = 10
    
    myArray(3, 1) = 1997
    myArray(3, 2) = 14.3
    myArray(3, 3) = 33
    myArray(3, 4) = 20.5
    myArray(3, 5) = 49.3
    myArray(3, 6) = 285.5
    myArray(3, 7) = 75.6
    myArray(3, 8) = 229.9
    myArray(3, 9) = 343.5
    myArray(3, 10) = 52.9
    myArray(3, 11) = 17.6
    myArray(3, 12) = 95.8
    myArray(3, 13) = 40
    
    myArray(4, 1) = 1998
    myArray(4, 2) = 17
    myArray(4, 3) = 40.4
    myArray(4, 4) = 42.4
    myArray(4, 5) = 110.6
    myArray(4, 6) = 103.2
    myArray(4, 7) = 187.4
    myArray(4, 8) = 326.3
    myArray(4, 9) = 568.1
    myArray(4, 10) = 190.6
    myArray(4, 11) = 20.6
    myArray(4, 12) = 29.4
    myArray(4, 13) = 2.1
    
    myArray(5, 1) = 1999
    myArray(5, 2) = 13.6
    myArray(5, 3) = 1.9
    myArray(5, 4) = 55.5
    myArray(5, 5) = 82.6
    myArray(5, 6) = 109.6
    myArray(5, 7) = 67.4
    myArray(5, 8) = 187.4
    myArray(5, 9) = 565.1
    myArray(5, 10) = 247.6
    myArray(5, 11) = 91.4
    myArray(5, 12) = 31.7
    myArray(5, 13) = 18.7
    
    myArray(6, 1) = 2000
    myArray(6, 2) = 47.1
    myArray(6, 3) = 1.5
    myArray(6, 4) = 2.7
    myArray(6, 5) = 27.1
    myArray(6, 6) = 63.5
    myArray(6, 7) = 62.3
    myArray(6, 8) = 78.7
    myArray(6, 9) = 591
    myArray(6, 10) = 210.2
    myArray(6, 11) = 20.9
    myArray(6, 12) = 32.3
    myArray(6, 13) = 22.1
    
    myArray(7, 1) = 2001
    myArray(7, 2) = 40.2
    myArray(7, 3) = 37.3
    myArray(7, 4) = 12.1
    myArray(7, 5) = 6.7
    myArray(7, 6) = 19.7
    myArray(7, 7) = 153.4
    myArray(7, 8) = 591.7
    myArray(7, 9) = 168.4
    myArray(7, 10) = 3.4
    myArray(7, 11) = 77
    myArray(7, 12) = 14
    myArray(7, 13) = 20.6
    
    myArray(8, 1) = 2002
    myArray(8, 2) = 35.3
    myArray(8, 3) = 2.3
    myArray(8, 4) = 28.3
    myArray(8, 5) = 139.5
    myArray(8, 6) = 66.9
    myArray(8, 7) = 45.2
    myArray(8, 8) = 214.7
    myArray(8, 9) = 411.4
    myArray(8, 10) = 24.9
    myArray(8, 11) = 40.9
    myArray(8, 12) = 14.2
    myArray(8, 13) = 10.1
    
    myArray(9, 1) = 2003
    myArray(9, 2) = 9.5
    myArray(9, 3) = 31.7
    myArray(9, 4) = 26.3
    myArray(9, 5) = 129
    myArray(9, 6) = 99.6
    myArray(9, 7) = 148.8
    myArray(9, 8) = 351.4
    myArray(9, 9) = 588.3
    myArray(9, 10) = 210.5
    myArray(9, 11) = 35.5
    myArray(9, 12) = 61.2
    myArray(9, 13) = 10.4
    
    myArray(10, 1) = 2004
    myArray(10, 2) = 18.5
    myArray(10, 3) = 48
    myArray(10, 4) = 14.2
    myArray(10, 5) = 64.4
    myArray(10, 6) = 169
    myArray(10, 7) = 97.5
    myArray(10, 8) = 410.4
    myArray(10, 9) = 111
    myArray(10, 10) = 259.1
    myArray(10, 11) = 5
    myArray(10, 12) = 87.4
    myArray(10, 13) = 23
    
    myArray(11, 1) = 2005
    myArray(11, 2) = 2.5
    myArray(11, 3) = 17
    myArray(11, 4) = 12.9
    myArray(11, 5) = 76.4
    myArray(11, 6) = 82
    myArray(11, 7) = 145.3
    myArray(11, 8) = 228.9
    myArray(11, 9) = 201.3
    myArray(11, 10) = 302.6
    myArray(11, 11) = 41.3
    myArray(11, 12) = 38.7
    myArray(11, 13) = 6.9
    
    myArray(12, 1) = 2006
    myArray(12, 2) = 33.2
    myArray(12, 3) = 14.9
    myArray(12, 4) = 5
    myArray(12, 5) = 35
    myArray(12, 6) = 186.5
    myArray(12, 7) = 134
    myArray(12, 8) = 765
    myArray(12, 9) = 36.2
    myArray(12, 10) = 11.2
    myArray(12, 11) = 19.5
    myArray(12, 12) = 36.2
    myArray(12, 13) = 23.4
    
    myArray(13, 1) = 2007
    myArray(13, 2) = 3.5
    myArray(13, 3) = 10
    myArray(13, 4) = 108.3
    myArray(13, 5) = 28
    myArray(13, 6) = 134.1
    myArray(13, 7) = 59.5
    myArray(13, 8) = 228.5
    myArray(13, 9) = 239
    myArray(13, 10) = 224.4
    myArray(13, 11) = 52.7
    myArray(13, 12) = 24.9
    myArray(13, 13) = 7.1
    
    myArray(14, 1) = 2008
    myArray(14, 2) = 13.7
    myArray(14, 3) = 4.9
    myArray(14, 4) = 45.8
    myArray(14, 5) = 48.5
    myArray(14, 6) = 63.9
    myArray(14, 7) = 102.3
    myArray(14, 8) = 522.8
    myArray(14, 9) = 165.7
    myArray(14, 10) = 76
    myArray(14, 11) = 55.7
    myArray(14, 12) = 18
    myArray(14, 13) = 20.1
    
    myArray(15, 1) = 2009
    myArray(15, 2) = 6.1
    myArray(15, 3) = 25.1
    myArray(15, 4) = 64.1
    myArray(15, 5) = 40.4
    myArray(15, 6) = 134.5
    myArray(15, 7) = 91.4
    myArray(15, 8) = 470.6
    myArray(15, 9) = 316
    myArray(15, 10) = 51
    myArray(15, 11) = 88.5
    myArray(15, 12) = 77
    myArray(15, 13) = 17.4
    
    myArray(16, 1) = 2010
    myArray(16, 2) = 29.2
    myArray(16, 3) = 50.1
    myArray(16, 4) = 62.9
    myArray(16, 5) = 57.3
    myArray(16, 6) = 104.9
    myArray(16, 7) = 200.2
    myArray(16, 8) = 275.4
    myArray(16, 9) = 485.1
    myArray(16, 10) = 454.1
    myArray(16, 11) = 29
    myArray(16, 12) = 13.4
    myArray(16, 13) = 16.1
    
    myArray(17, 1) = 2011
    myArray(17, 2) = 7
    myArray(17, 3) = 32.5
    myArray(17, 4) = 14.5
    myArray(17, 5) = 127.6
    myArray(17, 6) = 44.2
    myArray(17, 7) = 307.6
    myArray(17, 8) = 864.2
    myArray(17, 9) = 208.4
    myArray(17, 10) = 26.4
    myArray(17, 11) = 30.3
    myArray(17, 12) = 55.7
    myArray(17, 13) = 7.1
    
    myArray(18, 1) = 2012
    myArray(18, 2) = 4
    myArray(18, 3) = 0
    myArray(18, 4) = 26.6
    myArray(18, 5) = 104.5
    myArray(18, 6) = 14
    myArray(18, 7) = 90.8
    myArray(18, 8) = 425.3
    myArray(18, 9) = 365.2
    myArray(18, 10) = 161.1
    myArray(18, 11) = 77.7
    myArray(18, 12) = 95.5
    myArray(18, 13) = 50.4
    
    myArray(19, 1) = 2013
    myArray(19, 2) = 25.5
    myArray(19, 3) = 68.9
    myArray(19, 4) = 30.7
    myArray(19, 5) = 52.5
    myArray(19, 6) = 126.2
    myArray(19, 7) = 36.9
    myArray(19, 8) = 447.4
    myArray(19, 9) = 120.7
    myArray(19, 10) = 204.5
    myArray(19, 11) = 5.1
    myArray(19, 12) = 44.8
    myArray(19, 13) = 23.4
    
    myArray(20, 1) = 2014
    myArray(20, 2) = 7.9
    myArray(20, 3) = 16.7
    myArray(20, 4) = 8.5
    myArray(20, 5) = 44.1
    myArray(20, 6) = 74.3
    myArray(20, 7) = 62.8
    myArray(20, 8) = 227.2
    myArray(20, 9) = 122.1
    myArray(20, 10) = 93.9
    myArray(20, 11) = 60.3
    myArray(20, 12) = 46
    myArray(20, 13) = 24.3
    
    myArray(21, 1) = 2015
    myArray(21, 2) = 10.3
    myArray(21, 3) = 22.8
    myArray(21, 4) = 10
    myArray(21, 5) = 53.7
    myArray(21, 6) = 30.1
    myArray(21, 7) = 63.6
    myArray(21, 8) = 175.8
    myArray(21, 9) = 45.5
    myArray(21, 10) = 13.5
    myArray(21, 11) = 87
    myArray(21, 12) = 107.7
    myArray(21, 13) = 32
    
    myArray(22, 1) = 2016
    myArray(22, 2) = 2.8
    myArray(22, 3) = 43.1
    myArray(22, 4) = 44.1
    myArray(22, 5) = 80.8
    myArray(22, 6) = 148.5
    myArray(22, 7) = 19.5
    myArray(22, 8) = 300.5
    myArray(22, 9) = 26.5
    myArray(22, 10) = 34.5
    myArray(22, 11) = 78.7
    myArray(22, 12) = 18.2
    myArray(22, 13) = 67.1
    
    myArray(23, 1) = 2017
    myArray(23, 2) = 20.6
    myArray(23, 3) = 16.2
    myArray(23, 4) = 9.7
    myArray(23, 5) = 57
    myArray(23, 6) = 21.6
    myArray(23, 7) = 49.3
    myArray(23, 8) = 478.3
    myArray(23, 9) = 250.5
    myArray(23, 10) = 21.6
    myArray(23, 11) = 21.8
    myArray(23, 12) = 42.2
    myArray(23, 13) = 40.1
    
    myArray(24, 1) = 2018
    myArray(24, 2) = 4.5
    myArray(24, 3) = 27
    myArray(24, 4) = 78.6
    myArray(24, 5) = 112.2
    myArray(24, 6) = 158.9
    myArray(24, 7) = 144
    myArray(24, 8) = 148.8
    myArray(24, 9) = 214.9
    myArray(24, 10) = 46.1
    myArray(24, 11) = 116.2
    myArray(24, 12) = 66.6
    myArray(24, 13) = 16.6
    
    myArray(25, 1) = 2019
    myArray(25, 2) = 0.7
    myArray(25, 3) = 26.6
    myArray(25, 4) = 28.4
    myArray(25, 5) = 41.1
    myArray(25, 6) = 33.2
    myArray(25, 7) = 49.7
    myArray(25, 8) = 220.8
    myArray(25, 9) = 143.7
    myArray(25, 10) = 229
    myArray(25, 11) = 31.9
    myArray(25, 12) = 89.4
    myArray(25, 13) = 25
    
    myArray(26, 1) = 2020
    myArray(26, 2) = 48.6
    myArray(26, 3) = 49.7
    myArray(26, 4) = 10.7
    myArray(26, 5) = 13.8
    myArray(26, 6) = 101.2
    myArray(26, 7) = 100.2
    myArray(26, 8) = 243.5
    myArray(26, 9) = 486.4
    myArray(26, 10) = 167.7
    myArray(26, 11) = 1.9
    myArray(26, 12) = 81.6
    myArray(26, 13) = 6.6
    
    myArray(27, 1) = 2021
    myArray(27, 2) = 20.8
    myArray(27, 3) = 6.6
    myArray(27, 4) = 93.4
    myArray(27, 5) = 116.9
    myArray(27, 6) = 191.5
    myArray(27, 7) = 45.4
    myArray(27, 8) = 84.2
    myArray(27, 9) = 269
    myArray(27, 10) = 125.8
    myArray(27, 11) = 31.6
    myArray(27, 12) = 79
    myArray(27, 13) = 7.6
    
    myArray(28, 1) = 2022
    myArray(28, 2) = 4.4
    myArray(28, 3) = 2.7
    myArray(28, 4) = 84.6
    myArray(28, 5) = 21.1
    myArray(28, 6) = 5.4
    myArray(28, 7) = 286
    myArray(28, 8) = 215.2
    myArray(28, 9) = 635.9
    myArray(28, 10) = 167.8
    myArray(28, 11) = 102.1
    myArray(28, 12) = 81.3
    myArray(28, 13) = 14
    
    myArray(29, 1) = 2023
    myArray(29, 2) = 47.2
    myArray(29, 3) = 0.5
    myArray(29, 4) = 10
    myArray(29, 5) = 70.4
    myArray(29, 6) = 133.8
    myArray(29, 7) = 116.2
    myArray(29, 8) = 370.8
    myArray(29, 9) = 297.5
    myArray(29, 10) = 129.2
    myArray(29, 11) = 16.8
    myArray(29, 12) = 63.1
    myArray(29, 13) = 75.9
    
    myArray(30, 1) = 2024
    myArray(30, 2) = 16.3
    myArray(30, 3) = 63.8
    myArray(30, 4) = 19.2
    myArray(30, 5) = 35.3
    myArray(30, 6) = 122.6
    myArray(30, 7) = 101.8
    myArray(30, 8) = 342.9
    myArray(30, 9) = 86.7
    myArray(30, 10) = 122.9
    myArray(30, 11) = 82.1
    myArray(30, 12) = 62.2
    myArray(30, 13) = 5


    data_INCHEON = myArray

End Function



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
        
        
        fName = item.CodeModule.name
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
        lineToPrint = item.name
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
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.name
       
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
' Class Module: Class_ReturnTrueFalse
Private mValue As Boolean

Private Sub Class_Initialize()
    ' Initialize default values
    mValue = False
End Sub

Public Property Let result(val As Boolean)
    mValue = val
End Property

Public Property Get result() As Boolean
    result = mValue
End Property

Sub DeleteAllSummaryPage(ByVal well_str As String)
' delete all summary page

    Dim nof_p, i As Integer
    nof_p = GetNumberOf_P
    
    For i = 1 To nof_p
        Application.DisplayAlerts = False
        On Error Resume Next
        
        Worksheets("p" & i).Delete
        
        On Error GoTo 0
        Application.DisplayAlerts = True
    Next i
    
    Sheets(well_str).Activate
End Sub



Sub GetWaterSpecFromYangSoo_Q1()
  Dim thisname, fName As String
  Dim cell  As String
  Dim Time As Date
  Dim bTemp, ec1, ph1 As Double
  
  
  cell = Range("d12").value
  
  thisname = ActiveWorkbook.name
  fName = "A" & GetNumeric2(cell) & "_ge_OriginalSaveFile.xlsm"
 
  If Not IsWorkBookOpen(fName) Then
    MsgBox "Please open the yangsoo data ! " & fName
    Exit Sub
  End If
  
  ' Range("k2") = fname
   
  '------------------------------------------------------------------------
  Time = Workbooks(fName).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fName).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fName).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fName).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  Range("c6").value = Time
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Call TurnOffStuff
  Call getModDataFromYangSooSingle(thisname, fName)
  Call TurnOnStuff
End Sub


Sub GetWaterSpecFromYangSoo_Q2()
  Dim thisname, fname1, fname2 As String
  Dim cell1, cell2 As String
  Dim time1 As Date
  Dim bTemp1, bTemp2, ec1, ec2, ph1, ph2 As Double
  
  
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  ' Range("k2") = fname1
  ' Range("k3") = fname2
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  bTemp1 = Workbooks(fname1).Worksheets("w1").Range("c7").value
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp1
  Range("d7").value = bTemp2
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Call TurnOffStuff
  Call getModDataFromYangSooDual(thisname, fname1)
  Call getModDataFromYangSooDual(thisname, fname2)
  Call TurnOnStuff
End Sub



Sub GetWaterSpecFromYangSoo_Q3()
  Dim thisname, fname1, fname2, fname3 As String
  Dim cell1, cell2, cell3 As String
  Dim time1 As Date
  Dim bTemp, bTemp2, bTemp3, ec1, ec2, ec3, ph1, ph2, ph3 As Double
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  cell3 = Range("j12").value
  
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  fname3 = "A" & GetNumeric2(cell3) & "_ge_OriginalSaveFile.xlsm"
   
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname3) Then
    MsgBox "Please open the yangsoo data ! " & fname3
    Exit Sub
  End If
  
  'Range("k2") = fname1
  'Range("k3") = fname2
  'Range("k4") = fname3
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  bTemp3 = Workbooks(fname3).Worksheets("w1").Range("c7").value
  ec3 = Workbooks(fname3).Worksheets("w1").Range("c8").value
  ph3 = Workbooks(fname3).Worksheets("w1").Range("c9").value
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("d7").value = bTemp2
  Range("e7").value = bTemp3
  
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Range("e8").value = ec3
  Range("e9").value = ph3
  
  Call TurnOffStuff
  Call getModDataFromYangSooTripple(thisname, fname1)
  Call getModDataFromYangSooTripple(thisname, fname2)
  Call getModDataFromYangSooTripple(thisname, fname3)
  Call TurnOnStuff

End Sub




'******************************************************************************************************************************




Sub getModDataFromYangSooSingle(ByVal thisname As String, ByVal fName As String)
    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    Range("h14").Select
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub getModDataFromYangSooDual(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 2

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("h31").Select
    Else
        Range("d31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub



Sub getModDataFromYangSooTripple(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 3

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("l31").Select
    ElseIf f = 1 Then
        Range("d31").Select
    Else
        Range("h31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


'******************************************************************************************************************************



Function getEC_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & well).Activate
    
    If LOWHI = cellLOW Then
        getEC_Q1 = Sheets("p" & CStr(well)).Range("e25").value
    Else
        getEC_Q1 = Sheets("p" & CStr(well)).Range("e24").value
    End If
End Function

Function getPH_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate
    
    If LOWHI = cellLOW Then
        getPH_Q1 = Sheets("p" & CStr(well)).Range("f25").value
    Else
        getPH_Q1 = Sheets("p" & CStr(well)).Range("f24").value
    End If
    
End Function

Function getTEMP_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate

    If LOWHI = cellLOW Then
        getTEMP_Q1 = Sheets("p" & CStr(well)).Range("d25").value
    Else
        getTEMP_Q1 = Sheets("p" & CStr(well)).Range("d24").value
    End If
End Function


Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function



Sub DuplicateQ1Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & n
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
    
    Call SetWellPropertyQ1(n)
    
End Sub

Sub SetWellPropertyQ1(ByVal i As Integer)
    ActiveSheet.Range("C4") = "W-" & CStr(i)
    ActiveSheet.Range("D12") = "W-" & CStr(i)
    ActiveSheet.Range("H12") = "W-" & CStr(i)
    
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


'******************************************************************************************************************************


Function getEC_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getEC_Q2 = Sheets("p" & CStr(page)).Range("e25").value
        Else
            getEC_Q2 = Sheets("p" & CStr(page)).Range("e24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getEC_Q2 = Sheets("p" & CStr(page)).Range("h25").value
        Else
            getEC_Q2 = Sheets("p" & CStr(page)).Range("h24").value
        End If
    End If
End Function

Function getPH_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getPH_Q2 = Sheets("p" & CStr(page)).Range("f25").value
        Else
            getPH_Q2 = Sheets("p" & CStr(page)).Range("f24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getPH_Q2 = Sheets("p" & CStr(page)).Range("i25").value
        Else
            getPH_Q2 = Sheets("p" & CStr(page)).Range("i24").value
        End If
    End If
End Function

Function getTEMP_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("d25").value
        Else
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("d24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("g25").value
        Else
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("g24").value
        End If
    End If
End Function


Sub DuplicateQ2Page(ByVal n As Integer)
' n : Q2 page ������ ȸ��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .themeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ2(i)
    Next i
End Sub


Sub SetWellPropertyQ2(ByVal i As Integer)
' i : index of well

    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 2 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 2 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 2 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 2 + 2)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
End Sub


Sub SetWellPropertyRestQ2(ByVal w2page As Integer)
    Dim firstwell As Integer
      
    firstwell = 2 * w2page + 1
    
    ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
    ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub



Function DivideWellsBy2(ByVal numberOfWells As Integer) As Integer()
    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = (numberOfWells - 1) \ 2
    remainder = numberOfWells Mod 2
    
    
    If remainder = 0 Then
        result(0) = quotient + 1
    Else
        result(0) = quotient
    End If
    
    result(1) = remainder
    
    DivideWellsBy2 = result
End Function

Sub DuplicateRestQ2(ByVal w2page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w2page + 1)
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRestQ2(w2page)
    
End Sub


'**********************************************************************************************************




' 1, 2, 3 --> p1
' 4, 5, 6 --> p2

Function getEC_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("k25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("k24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("e25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("e24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("h25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("h24").value
            End If
    End Select
End Function

Function getPH_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("l25").value
            Else
                getPH_Q3 = Sheets("p" & CStr(page)).Range("l24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("f25").value
            Else
                getPH_Q3 = Sheets("p" & CStr(page)).Range("f24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("i25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("i24").value
            End If
    End Select
End Function

Function getTEMP_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("J25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("J24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("d25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("d24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("g25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("g24").value
            End If
    End Select
End Function




Function DivideWellsBy3(ByVal numberOfWells As Integer) As Integer()

    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = numberOfWells \ 3
    remainder = numberOfWells Mod 3
    
    result(0) = quotient
    result(1) = remainder
    
    DivideWellsBy3 = result
    
End Function




Sub SetWellPropertyQ3(ByVal i As Integer)
    
    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 3 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 3 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 3 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 3 + 2)
    
    ActiveSheet.Range("J12") = "W-" & CStr((i - 1) * 3 + 3)
    ActiveSheet.Range("L29") = "W-" & CStr((i - 1) * 3 + 3)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    
End Sub

Sub SetWellPropertyRestQ3(ByVal wselect As Integer, ByVal w3page As Integer)
    Dim firstwell As Integer
      
    firstwell = 3 * w3page + 1
    
    If wselect = 2 Then
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        
        
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("D29") = "W-" & CStr(firstwell)
        
        ActiveSheet.Range("G12") = "W-" & CStr(firstwell + 1)
        ActiveSheet.Range("H29") = "W-" & CStr(firstwell + 1)
    Else
    
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
        Selection.Delete
    
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    End If
End Sub

Sub DuplicateQ3Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q3")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .themeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ3(i)
    Next i
        
End Sub

Sub DuplicateRestQ3(ByVal wselect As Integer, ByVal w3page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q" & CStr(wselect))
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w3page + 1)
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRestQ3(wselect, w3page)
    
End Sub



'******************************************************************************************************************************



Function CheckWorkbookNameWithRegex(ByVal WB_NAME As String) As Boolean
    Dim regex As Object
    Dim pattern As String
    Dim match As Object

    ' Create the regex object
    Set regex = CreateObject("VBScript.RegExp")

    ' Define the pattern
    ' \bA(1|[2-9]|1[0-9]|2[0-9]|30)_ge_OriginalSaveFile
    pattern = "\bA(1|[2-9]|1[0-9]|2[0-9]|30)_ge_OriginalSaveFile.xlsm"

    ' Configure the regex object
    With regex
        .pattern = pattern
        .IgnoreCase = True
        .Global = False
    End With

    ' Check for the pattern
    If regex.test(WB_NAME) Then
        Set match = regex.Execute(WB_NAME)
        Debug.Print "The workbook name contains the pattern: " & match(0).value
        CheckWorkbookNameWithRegex = True
    Else
        Debug.Print "The workbook name does not contain the pattern."
        CheckWorkbookNameWithRegex = False
    End If
End Function

Function IsOpenedYangSooFiles() As Boolean
'
' ����Ϻ�����, A1_ge_OriginalSaveFile �� �����־
' ����Ϻ��� ������, ������ ������ ������ True
' �׷��� ������ False
'
    Dim fileName, WBNAME As String
    Dim nof_yangsoo As Integer
    Dim nofwell As Integer
    
    nof_yangsoo = 0
    nofwell = sheets_count()
    
    For Each Workbook In Application.Workbooks
        WBNAME = Workbook.name
        If StrComp(ThisWorkbook.name, WBNAME, vbTextCompare) = 0 Then
        ' �̸��� thisworkbook.name �� ���ٸ� , �����б��
            GoTo NEXT_ITERATION
        End If
        
        If CheckWorkbookNameWithRegex(WBNAME) Then
            nof_yangsoo = nof_yangsoo + 1
        End If
        
NEXT_ITERATION:
    Next
    
    If nof_yangsoo = nofwell Then
        IsOpenedYangSooFiles = True
    Else
        IsOpenedYangSooFiles = False
    End If

End Function


Sub PressAll_Button()
' Push All Button
' Fx - Collect Data
' Fx - Formula
' ImportAll, Collect Each Well
' Agg2
' Agg1
' AggStep
' AggChart
' AggWhpa
'
    If Not IsOpenedYangSooFiles() Then
        Popup_MessageBox ("YangSoo File is Does not match with number of well")
        Exit Sub
    End If

    Call Popup_MessageBox("YangSoo, modAggFX - get Data from YangSoo ilbo ...")
        
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
    Call GetBaseDataFromYangSoo(999, False)
    Sheets("YangSoo").Visible = False
    
    Call Popup_MessageBox("YangSoo, Aggregate2 - ImportWellSpec ...")
    

    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
    Call modAgg2.GROK_ImportWellSpec(999, False)
    Sheets("Aggregate2").Visible = False
    
    Call Popup_MessageBox("YangSoo, Aggregate1 - AggregateOne_Import ...")
    

    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
    Call modAgg1.ImportAggregateData(999, False)
    Sheets("Aggregate1").Visible = False
    
    Call Popup_MessageBox("YangSoo, AggStep - Import StepTest Data ...")
     
    Sheets("AggStep").Visible = True
    Sheets("AggStep").Select
    Call modAggStep.WriteStepTestData(999, False)
    Sheets("AggStep").Visible = False
    
    Call Popup_MessageBox("YangSoo, AggChart - Chart Import...")
   
    Sheets("AggChart").Visible = True
    Sheets("AggChart").Select
    Call modAggChart.WriteAllCharts(999, False)
    Sheets("AggChart").Visible = False
        

    Call Popup_MessageBox("Import All QT ...")
    Call modWell.ImportAll_QT
    
    Call Popup_MessageBox("ImportAll Each Well Spec ...")
    Call modWell.ImportAll_EachWellSpec
    
    Call Popup_MessageBox("ImportWell MainWellPage ...")
    Call modWell.ImportWell_MainWellPage
    
    Call Popup_MessageBox("Push Drastic Index ...")
    Call modWell.PushDrasticIndex
    
    

End Sub



Function RemoveSheetIfExists(shname As String) As Boolean
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shname)
    If Not ws Is Nothing Then sheetExists = True
    On Error GoTo 0

    If sheetExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        RemoveSheetIfExists = True
        Exit Function
    Else
        RemoveSheetIfExists = False
        Exit Function
    End If
End Function


Public Sub AddWell_CopyOneSheet()
    Dim n_sheets    As Integer
    
    n_sheets = sheets_count()
    
    '2020/5/30 ��������Ʈ�� ��ϻ������ִ� �κ� �߰�
    InsertOneRow (n_sheets)
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        Call DeleteCommandButton
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
    
    '2022/6/9 ��
    Range("i2") = "A" & CStr(n_sheets + 1) & "_ge_OriginalSaveFile.xlsm"
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
    Sheets("Well").Select
End Sub



' --------------------------------------------------------------------------------------------------------------


Sub DeleteCommandButton()
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Delete
End Sub



Sub InsertOneRow(ByVal n_sheets As Integer)
    n_sheets = n_sheets + 4
    Rows(n_sheets & ":" & n_sheets).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Rows(CStr(n_sheets - 1) & ":" & CStr(n_sheets - 1)).Select
    Selection.Copy
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
End Sub

Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
    ' change sheet data direct to well sheet data value
    ' https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    
    
    ' minhwasoo, 2023-10-13
    ' block, this code ....
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub



' --------------------------------------------------------------------------------------------------------------


Sub JojungButton()
    Dim nofwell As Integer

    TurnOffStuff

    nofwell = sheets_count()
    Call JojungSheetData
    Call make_wellstyle
    Call DecorateWellBorder(nofwell)
    
    Worksheets("1").Range("E21") = "=Well!" & Cells(5 + GetNumberOfWell(), "I").Address
    
    TurnOnStuff
End Sub

Sub Make_OneButton()
    Dim i, nofwell As Integer
    Dim response As VbMsgBoxResult
    
    nofwell = GetNumberOfWell()
    
    If nofwell = 1 Then Exit Sub
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
         For i = 2 To nofwell
             RemoveSheetIfExists (CStr(i))
        Next i
        
        Sheets("Well").Activate
        Rows("5:" & CStr(nofwell + 3)).Select
        Selection.Delete Shift:=xlUp
        
        For i = 1 To 12
            If Not RemoveSheetIfExists("p" & CStr(i)) Then Exit For
        Next i
        
        Call DecorateWellBorder(1)
        Range("A1").Select
    End If
End Sub


Sub DeleteLast()
' delete last

    Dim nofwell As Integer
    'nofwell = GetNumberOfWell()
    nofwell = sheets_count()
    
    If nofwell = 1 Then
        MsgBox "Last is not delete ... ", vbOK
        Exit Sub
    End If
    
    Rows(nofwell + 3).Delete
    Call DeleteWorksheet(nofwell)
    Call DecorateWellBorder(nofwell - 1)
End Sub



Sub DeleteWorksheet(well As Integer)
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(CStr(well)).Delete
    Application.DisplayAlerts = True
End Sub


Sub DecorateWellBorder(ByVal nofwell As Integer)
    Sheets("Well").Activate
    Range("A2:R" & CStr(nofwell + 3)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("D15").Select
End Sub




Sub getDuoSolo(ByVal nofwell As Integer, ByRef nDuo As Integer, ByRef nSolo As Integer)
    Dim page, quotient, remainder As Integer
    
    quotient = WorksheetFunction.quotient(nofwell, 2)
    remainder = nofwell Mod 2
    
    If remainder = 0 Then
        nDuo = quotient
        nSolo = 0
    Else
        nDuo = quotient
        nSolo = 1
    End If

End Sub


Sub ImportAll_EachWellSpec()
'
' �������� ��ȸ�ϸ鼭, ��������Ÿ�� �������� ���ش�.
'
    Dim nofwell, i  As Integer
    ' Dim obj As New Class_Boolean

    nofwell = sheets_count()
    
    BaseData_ETC_02.TurnOffStuff
    
    For i = 1 To nofwell
        Sheets(CStr(i)).Activate
        Call modWell_Each.ImportWellSpecFX(i)
    Next i
    
    Sheets("Well").Activate
    
    BaseData_ETC_02.TurnOnStuff
End Sub


Sub ImportAll_EachWellSpec_OLD()
'
' �������� ��ȸ�ϸ鼭, ��������Ÿ�� �������� ���ش�.
'
    Dim nofwell, i  As Integer
    Dim obj As New Class_Boolean

    nofwell = sheets_count()
    
    BaseData_ETC_02.TurnOffStuff
    
    For i = 1 To nofwell
        Sheets(CStr(i)).Activate
        Call Module_ImportWellSpec.ImportWellSpec_OLD(i, obj)
        If obj.result Then Exit For
    Next i
    
    Sheets("Well").Activate
    
    BaseData_ETC_02.TurnOnStuff
End Sub




Sub ImportWell_MainWellPage()
'
' import Sheets("Well") Page
'
    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Address, Company As String
    Dim simdo, diameter, Q, Hp As Double
    
    nofwell = sheets_count()
    Sheets("Well").Select
    
    Dim wsYangSoo, wsWell, wsRecharge As Worksheet
    Set wsYangSoo = Worksheets("YangSoo")
    Set wsWell = Worksheets("Well")
    Set wsRecharge = Worksheets("Recharge")
    
    '2024,12,25 - Add Title
    wsWell.Range("D1").value = wsYangSoo.Cells(5, "AR").value
    
    Call TurnOffStuff
           
    For i = 1 To nofwell
        '2025/3/5
        Address = Replace(wsYangSoo.Cells(4 + i, "ao").value, "��û���� ", "")
        Address = Replace(Address, "����", "")
        
        simdo = wsYangSoo.Cells(4 + i, "i").value
        diameter = wsYangSoo.Cells(4 + i, "g").value
        Q = wsYangSoo.Cells(4 + i, "k").value
        Hp = wsYangSoo.Cells(4 + i, "m").value
        
        wsWell.Cells(3 + i, "d").value = Address
        wsWell.Cells(3 + i, "g").value = diameter
        wsWell.Cells(3 + i, "h").value = simdo
        wsWell.Cells(3 + i, "i").value = Q
        wsWell.Cells(3 + i, "j").value = Q
        wsWell.Cells(3 + i, "l").value = Hp
    Next i

    
    Company = wsYangSoo.Range("AP5").value
    wsRecharge.Range("B32").value = Company
    
    Application.CutCopyMode = False
    Call TurnOnStuff
End Sub





Sub DuplicateBasicWellData()
' 2024/6/24 - dupl, duplicate basic well data ...
' �⺻��������Ÿ �����ϴ°�
' ������ ��ȸ�ϸ鼭, �ű⿡�� �����͸� ������ ���µ� ��
' ���� , �����, �����
' ����, �Ÿ�, ��������, ��ǥ��ǥ�� �̷��� ������ ���� �ɵ��ϴ�.

' k6 - ����� / long axis
' k7 - ����� / short axis
' k12 - degree of flow
' k13 - well distance
' k14 - well height
' k15 - surfacewater height

    Dim nofwell, i  As Integer
    Dim obj As New Class_Boolean
    Dim WB_NAME As String
    Dim weather_station, river_section As String
    

    nofwell = sheets_count()
     
    WB_NAME = mod_DuplicatetWellSpec.GetOtherFileName
    
    If WB_NAME = "NOTHING" Then
        MsgBox "�⺻��������Ÿ�� �����ؾ� �ϹǷ�, �⺻���������͸� ����νñ� �ٶ��ϴ�. ", vbOK
        Exit Sub
    Else
        BaseData_ETC_02.TurnOffStuff
        
        Call mod_DuplicatetWellSpec.Duplicate_WATER(ThisWorkbook.name, WB_NAME)
        Call mod_DuplicatetWellSpec.Duplicate_WELL_MAIN(ThisWorkbook.name, WB_NAME, nofwell)
        weather_station = Replace(Sheets("Well").Range("F4").value, "���û", "")
        river_section = Sheets("Well").Range("E4").value
        
        ' 2024/6/27 ��, ���� �߰����� ������� �������� ...
        ThisWorkbook.Sheets("Recharge").Range("b32") = Range("B4").value
        
        ' �� ������ ������ ����
        For i = 1 To nofwell
            Sheets(CStr(i)).Activate
            Call mod_DuplicatetWellSpec.DuplicateWellSpec(ThisWorkbook.name, WB_NAME, i, obj)
            
            If obj.result Then Exit For
        Next i
        
        Worksheets("Well").Activate
        
        'WSet Button, CommandButton14
        For i = 1 To nofwell
            Cells(i + 3, "E").formula = "=Recharge!$I$24"
            Cells(i + 3, "F").formula = "=All!$B$2"
            Cells(i + 3, "O").formula = "=ROUND(water!$F$7, 1)"
            
            Cells(i + 3, "B").formula = "=Recharge!$B$32"
        Next i
        
        Sheets("Well").Activate
        BaseData_ETC_02.TurnOnStuff
    End If
    
     ' ��ǿ�, �߱ǿ� ����
     Sheets("Recharge").Range("I24") = river_section
     
     ' 2024/7/9 Add, Company Name Setting
     Sheets("Recharge").Range("B32") = Sheets("YangSoo").Range("AP5")
     
     
    ' ���û ����Ÿ, �ٽ� �ҷ�����
    If Not BaseData_ETC.CheckSubstring(Sheets("All").Range("T5").value, weather_station) Then
         Call modProvince.ResetWeatherData(weather_station)
    End If
    
    Call modWell.PushDrasticIndex

End Sub


Sub ImportAll_QT()
'
' ������� ������ȭ���
'
    Dim i, nof_p As Integer
    Dim qt As String
    
    nof_p = GetNumberOf_P
    
    For i = 1 To nof_p
        Sheets("p" & i).Activate
        qt = determin_Q_Type
        
        Application.Run "modWaterQualityTest.GetWaterSpecFromYangSoo_" & qt
    Next i
End Sub


Function determin_Q_Type() As String
' �̰���, p1, p2, p3 �� � Ÿ������ üũ�ϴºκ�
' �� Q1, Q2, Q3 ���� �˾Ƴ��°�
' D12 --- q1
' G12 --- q2
' J12 --- q3

    If Range("J12").value <> "" Then
        determin_Q_Type = "Q3"
    ElseIf Range("G12").value <> "" Then
        determin_Q_Type = "Q2"
    Else
        determin_Q_Type = "Q1"
    End If

End Function

Function GetNumberOf_P()
    Dim nofwell, i, nof_p As Integer

    nofwell = sheets_count()
    nof_p = 0
    
    For Each sheet In Worksheets
        If Left(sheet.name, 1) = "p" And ConvertToLongInteger(Right(sheet.name, 1)) <> 0 Then
            nof_p = nof_p + 1
        End If
    Next

    GetNumberOf_P = nof_p
End Function


Sub PushDrasticIndex()

    Call BaseData_DrasticIndex.main_drasticindex
    Call BaseData_DrasticIndex.print_drastic_string
    
End Sub


'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>
' 2024,12,25 Add import "Title "
'
Sub FXSAVE_GetBaseDataFromYangSoo(ByVal WellNumber As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String

    ' Arrays to store data
    Dim dataArrays As Variant
    dataArrays = Array("natural", "stable", "recover", "delta_h", "Sw", "radius", _
                       "Rw", "well_depth", "casing", "Q", "delta_s", "hp", _
                       "daeSoo", "T1", "T2", "TA", "S1", "S2", "K", "time_", _
                       "shultze", "webber", "jacob", "skin", "er", "ER1", _
                       "ER2", "ER3", "qh", "qg", "sd1", "sd2", "q1", "C", _
                       "B", "ratio", "T0", "S0", "ER_MODE", "Address", "Company", "S3", "Title")

    ' Check if all well data should be imported
    nofwell = GetNumberOfWell()
    If Not isSingleWellImport And WellNumber = 999 Then
        rngString = "A5:AR37"
    Else
       rngString = "A" & (WellNumber + 5 - 1) & ":AR" & (WellNumber + 5 - 1)
    End If

    Call EraseCellData(rngString)

    ' Loop through each well
    For i = 1 To nofwell
        ' Import data for all wells or only for the specified single well
        If Not isSingleWellImport Or (isSingleWellImport And i = WellNumber) Then
            ImportDataForWell i, dataArrays
        End If
    Next i
End Sub

Sub FXSAVE_ImportDataForWell(ByVal wellIndex As Integer, ByVal dataArrays As Variant)
    Dim fName As String
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataIdx As Integer
    Dim cellOffset As Integer
    Dim dataCell As Range

    ' Open the workbook
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Sub
    End If
    Set wb = Workbooks(fName)

    ' Loop through data arrays and import values
    For dataIdx = LBound(dataArrays) To UBound(dataArrays)
        SetDataArrayValues wb, wellIndex, dataArrays(dataIdx)
    Next dataIdx

    ' Close workbook
    ' wb.Close SaveChanges:=False
End Sub


Sub FXSAVE_SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range


    Dim dataRanges() As Variant
    Dim addresses() As Variant
    Dim i As Integer

    ' Set references to worksheets
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    ' Define data ranges for each dataArrayName
    dataRanges = Array(wsInput.Range("m51"), wsInput.Range("i48"), _
                        wsInput.Range("m48"), wsInput.Range("m49"), _
                        wsInput.Range("m44"), wsSkinFactor.Range("e4"), _
                        wsInput.Range("m45"), wsInput.Range("i52"), _
                        wsInput.Range("A31"), wsInput.Range("B31"), _
                        wsSkinFactor.Range("c10"), wsSkinFactor.Range("c11"), _
                        wsSkinFactor.Range("b16"), wsSkinFactor.Range("b4"), _
                        wsSkinFactor.Range("c16"), wsSkinFactor.Range("d4"), _
                        wsSkinFactor.Range("f4"), wsSkinFactor.Range("h10"), _
                        wsSkinFactor.Range("d5"), wsSkinFactor.Range("h13"), _
                        wsSkinFactor.Range("d16"), wsSkinFactor.Range("e10"), _
                        wsSkinFactor.Range("i16"), wsSkinFactor.Range("e16"), _
                        wsSkinFactor.Range("h16"), wsSkinFactor.Range("c13"), _
                        wsSkinFactor.Range("c18"), wsSkinFactor.Range("c23"), _
                        wsSkinFactor.Range("g6"), wsSkinFactor.Range("c8"), _
                        wsSkinFactor.Range("k8"), wsSkinFactor.Range("k9"), _
                        wsSkinFactor.Range("k10"), wsSafeYield.Range("b13"), _
                        wsSafeYield.Range("b7"), wsSafeYield.Range("b3"), _
                        wsSafeYield.Range("b4"), wsSafeYield.Range("b2"), _
                        wsSafeYield.Range("b11"), wsInput.Range("i46"), _
                        wsInput.Range("i47"), wsSkinFactor.Range("i13"), _
                        wsInput.Range("i44"))

    ' Array of data addresses
    addresses = Array("Q", "hp", "natural", "stable", "radius", "Rw", _
                        "well_depth", "casing", "C", "B", "recover", "Sw", _
                        "delta_h", "delta_s", "daeSoo", "T0", "S0", "ER_MODE", _
                        "T1", "T2", "TA", "S1", "S2", "K", "time_", "shultze", _
                        "webber", "jacob", "skin", "er", "ER1", "ER2", "ER3", _
                        "qh", "qg", "sd1", "sd2", "q1", "ratio", "Address", _
                        "Company", "S3", "Title")

    ' Find index of dataArrayName in addresses array
    For i = LBound(addresses) To UBound(addresses)
        If addresses(i) = dataArrayName Then
            Set dataCell = dataRanges(i)
            Exit For
        End If
    Next i

    ' Check if dataArrayName is found
    If Not dataCell Is Nothing Then
        SetCellValueForWell wellIndex, dataCell, dataArrayName
    Else
        MsgBox "Data array name not found: " & dataArrayName
    End If
End Sub


' 2024-8-22 : �����������޽ð�, time_ : 0.0000 �� ����

Sub FXSAVE_SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim WellData As Variant
    Dim numberFormats As Object
    Set numberFormats = CreateObject("Scripting.Dictionary")

    ' Define number formats for each dataArrayName
    With numberFormats
        .Add "recover", "0.00"
        .Add "Sw", "0.00"
        .Add "S2", "0.0000000"
        .Add "S3", "0.00"
        .Add "T1", "0.0000"
        .Add "T2", "0.0000"
        .Add "TA", "0.0000"
        .Add "qh", "0."
        .Add "qg", "0.00"
        .Add "q1", "0.00"
        .Add "sd1", "0.00"
        .Add "sd2", "0.00"
        .Add "skin", "0.0000"
        .Add "er", "0.0000"
        .Add "ratio", "0.0%"
        .Add "T0", "0.0000"
        .Add "S0", "0.0000"
        .Add "delta_s", "0.00"
        .Add "time_", "0.0000"
        .Add "shultze", "0.00"
        .Add "webber", "0.00"
        .Add "jacob", "0.00"

    End With

    ' Get value from dataCell
    WellData = dataCell.value

    Cells(4 + wellIndex, 1).value = "W-" & wellIndex

    ' Set value and number format based on dataArrayName
    With Cells(4 + wellIndex, GetColumnIndex(dataArrayName))
        .value = WellData
        If numberFormats.Exists(dataArrayName) Then
            .numberFormat = numberFormats(dataArrayName)
        End If
    End With
End Sub



Function FXSAVE_GetColumnIndex(ByVal columnName As String) As Integer
    ' Define array to store column indices
    Dim columnIndices As Variant
    columnIndices = Array( _
        11, 13, 2, 3, 7, 8, 9, 10, _
        32, 33, 4, 5, 6, 12, 14, _
        35, 36, 37, 15, 16, 17, 18, _
        19, 20, 21, 22, 23, 24, 25, _
        26, 38, 39, 40, 27, 28, 30, _
        31, 29, 34, 41, 42, 43, 44 _
    )

    ' Define array to store column names
    Dim columnNames As Variant
    columnNames = Array( _
        "Q", "hp", "natural", "stable", "radius", "Rw", "well_depth", "casing", _
        "C", "B", "recover", "Sw", "delta_h", "delta_s", "daeSoo", _
        "T0", "S0", "ER_MODE", "T1", "T2", "TA", "S1", _
        "S2", "K", "time_", "shultze", "webber", "jacob", "skin", _
        "er", "ER1", "ER2", "ER3", "qh", "qg", "sd1", _
        "sd2", "q1", "ratio", "Address", "Company", "S3", "Title" _
    )

    ' Find index of columnName in columnNames array
    Dim index As Integer
    index = Application.match(columnName, columnNames, 0)

    ' Check if columnName exists in columnNames array
    If IsNumeric(index) Then
        GetColumnIndex = columnIndices(index - 1)
    Else
        ' Return -1 if columnName is not found
        GetColumnIndex = -1
    End If
End Function



' in here by refctor by  openai
' replace GetBaseDataFromYangSoo Module
'
'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>

Public Sub FXSAVE_MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True

      DoEvents '  !!! DoEvents is VERY IMPORTANT here !!!

      Debug.Print String(30, vbCrLf)
   End If

   Debug.Print sPrintStr
End Sub


'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10

Function FXSAVE_DetermineEffectiveRadius(ERMode As String) As Integer
    Dim Er, R As String

    Er = ERMode
    'MsgBox er
    R = Mid(Er, 5, 1)

    If R = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(R)
    End If
End Function


Function FXSAVE_CheckFileExistence(filePath As String) As Boolean

    If Dir(filePath) <> "" Then
        CheckFileExistence = True
    Else
        CheckFileExistence = False
    End If

End Function


Sub FXSAVE_WriteFormula()

    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double

    ' Save array to a file
    Dim filePath As String
    Dim FileNum As Integer


    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select

    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    DeleteFileIfExists filePath
    FileNum = FreeFile

    Open filePath For Output As FileNum

    Call MyDebug("Formula SkinFactor ... ", True)

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

    ' ��Ų���
    Call FormulaSkinFactorAndER("SKIN", FileNum)

    ' ��ȿ�칰�ݰ�
    Call FormulaSkinFactorAndER("ER", FileNum)


    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    Call FormulaChwiSoo(FileNum)
    ' 3-7, ���������

    Call FormulaRadiusOfInfluence(FileNum)
    ' �������ݰ�

    Close FileNum

End Sub



Sub FXSAVE_FormulaSkinFactorAndER(ByVal Mode As String, ByVal FileNum As Integer)
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double


    Call MyDebug("Formula SkinFactor ... ", True)

    nofwell = GetNumberOfWell()

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    For i = 1 To nofwell
        T = format(Cells(4 + i, "o").value, "0.0000")
        Q = Cells(4 + i, "k").value

        T0 = format(Cells(4 + i, "AI").value, "0.0000")
        S0 = format(Cells(4 + i, "AJ").value, "0.0000")
        S1 = Cells(4 + i, "R").value

        delta_s = format(Cells(4 + i, "l").value, "0.00")
        Radius = format(Cells(4 + i, "h").value, "0.000")
        Skin = format(Cells(4 + i, "y").value, "0.0000")
        Er = format(Cells(4 + i, "z").value, "0.0000")


        B = format(Cells(4 + i, "AG").value, "0.0000")
        ER1 = format(Cells(4 + i, "AL").value, "0.0000")
        ER2 = format(Cells(4 + i, "AM").value, "0.0000")
        ER3 = format(Cells(4 + i, "AN").value, "0.0000")


        Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
        ' ����� 1��
        Case erRE1

            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
            formula2 = "erRE1, ����� 1��"

        ' ����� 2��
        Case erRE2
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
            formula2 = "erRE2, ����� 2��"
        ' ����� 3��
        Case erRE3
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~" & Radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
            formula2 = "erRE3, ����� 3��"

        Case Else
            ' ��Ų���
            formula1 = "W-" & i & "ȣ��~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {" & S0 & " TIMES  (" & Radius & " TIMES  " & Radius & ")} =`" & Skin
            ' ��ȿ�칰�ݰ�
            formula2 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & Radius & " TIMES e ^{-(" & Skin & ")} =" & Er & "m"
        End Select


        If Mode = "SKIN" Then
            Debug.Print formula1
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

            Print #FileNum, formula1
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Else
            Debug.Print formula2
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

            Print #FileNum, formula2
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    Next i

End Sub


Sub FXSAVE_DeleteFileIfExists(filePath As String)
    If Len(Dir(filePath)) > 0 Then ' Check if file exists
        On Error Resume Next
        Kill filePath ' Attempt to delete the file

        On Error GoTo 0
        If Len(Dir(filePath)) > 0 Then
            MsgBox "Unable to delete file '" & filePath & "'. The file may be in use.", vbExclamation
        Else
            ' MsgBox "File '" & filePath & "' has been deleted.", vbInformation

        End If
    Else
        ' MsgBox "File '" & filePath & "' does not exist.", vbExclamation
    End If
End Sub



Sub FXSAVE_FormulaChwiSoo(FileNum As Integer)
' 3-7, ���������

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim Q1, S1, S2, res As Double

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select


    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

    For i = 1 To nofwell
        Q1 = Cells(4 + i, "ac").value

        S1 = format(Cells(4 + i, "ad").value, "0.00")
        S2 = format(Cells(4 + i, "ae").value, "0.00")
        res = format(Cells(4 + i, "ab").value, "0.00")

        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =165.52��/��
        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =`165.52��/��

        formula = "W-" & i & "ȣ��~~Q_{& 2} = " & Q1 & " TIMES ({" & S2 & "} over {" & S1 & "}) ^{2/3} = `" & res & " ��/��"

        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

End Sub


Sub FXSAVE_FormulaRadiusOfInfluence(FileNum As Integer)
' �������ݰ�

    Call FormulaSUB_ROI("SCHULTZE", FileNum)
    Call FormulaSUB_ROI("WEBBER", FileNum)
    Call FormulaSUB_ROI("JCOB", FileNum)

End Sub




Sub FXSAVE_FormulaSUB_ROI(Mode As String, FileNum As Integer)
  Dim formula1, formula2, formula3 As String
    ' ����, ����, ������

    Dim nofwell As String
    Dim i As Integer
    Dim Shultze, Webber, Jacob, T, K, S, time_, delta_h As String

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select


    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    For i = 1 To nofwell
        schultze = CStr(format(Cells(4 + i, "v").value, "0.0"))
        Webber = CStr(format(Cells(4 + i, "w").value, "0.0"))
        Jacob = CStr(format(Cells(4 + i, "x").value, "0.0"))

        T = CStr(format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(format(Cells(4 + i, "t").value, "0.0000"))

        delta_h = CStr(format(Cells(4 + i, "f").value, "0.00"))
        time_ = CStr(format(Cells(4 + i, "u").value, "0.0000"))


        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")

        formula1 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & schultze & "m"
        formula2 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & Webber & "`m"
        formula3 = "W-" & i & "ȣ��~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & Jacob & "m"


        Select Case Mode
            Case "SCHULTZE"
                Debug.Print formula1
                Print #FileNum, formula1

            Case "WEBBER"
                Debug.Print formula2
                Print #FileNum, formula2

            Case "JCOB"
                Debug.Print formula3
                Print #FileNum, formula3
        End Select

        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

End Sub




Sub WriteStepTestData(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
'
' SingleWell --> ImportWell Number
' 999 & False --> �������� ����Ʈ
'

    Dim nofwell, i As Integer
    Dim a1, a2, a3, Q, h, delta_h, qsw, swq As String
    Dim fName As String
    
    nofwell = GetNumberOfWell()
    
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim rngString As String
    
    
    Call TurnOffStuff
    
    If ActiveSheet.name <> "AggStep" Then Sheets("AggStep").Select
    
    
    If isSingleWellImport Then
        rngString = "C" & (singleWell + 5 - 1) & ":K" & (singleWell + 5 - 1)
        Call EraseCellData(rngString)
    Else
        rngString = "C5:K36"
        
        fName = "A1_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Call EraseCellData(rngString)
    End If
        
    
    For i = 1 To nofwell
    
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
    
SINGLE_ITERATION:

        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Set wb = Workbooks(fName)
        Set wsInput = wb.Worksheets("Input")
        
        Q = wsInput.Range("q64").value
        h = wsInput.Range("r64").value
        delta_h = wsInput.Range("s64").value
        qsw = wsInput.Range("t64").value
        swq = wsInput.Range("u64").value

        a1 = wsInput.Range("v64").value
        a2 = wsInput.Range("w64").value
        a3 = wsInput.Range("x64").value
        
        Call Write31_StepTestData_Single(a1, a2, a3, Q, h, delta_h, qsw, swq, i)

NEXT_ITERATION:

    Next i
    
    Call TurnOnStuff
    'Call Write31_StepTestData(a1, a2, a3, Q, h, delta_h, qsw, swq, nofwell)
End Sub


Sub Write31_StepTestData_Single(a1 As Variant, a2 As Variant, a3 As Variant, Q As Variant, h As Variant, delta_h As Variant, qsw As Variant, swq As Variant, i As Integer)
' i : well_index
    
    ' Call EraseCellData("C5:K36")
    
    Cells(4 + i, "c").value = "W-" & CStr(i)
    
    Cells(4 + i, "d").value = a1
    Cells(4 + i, "e").value = a2
    Cells(4 + i, "f").value = a3

    Cells(4 + i, "g").value = Q
    Cells(4 + i, "h").value = h
    Cells(4 + i, "i").value = delta_h
    Cells(4 + i, "j").value = qsw
    Cells(4 + i, "k").value = swq

End Sub


Function getDirectionFromWell(i) As Integer

    If Sheets(CStr(i)).Range("k12").Font.Bold Then
        getDirectionFromWell = Sheets(CStr(i)).Range("k12").value
    Else
        getDirectionFromWell = Sheets(CStr(i)).Range("l12").value
    End If

End Function

Sub WriteWellData_Single(Q As Variant, DaeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, ByVal i As Integer)
    
    Call UnmergeAllCells
        
    Cells(3 + i, "c").value = "W-" & CStr(i)
    Cells(3 + i, "e").value = Q
    Cells(3 + i, "f").value = T1
    Cells(3 + i, "i").value = DaeSoo
    Cells(3 + i, "k").value = direction
    
    ' 2025/03/10 --> ABS Gradient
    Cells(3 + i, "m").value = format(Abs(gradient), "###0.0000")
    Cells(4, "d").value = "5��"
    
End Sub


Sub MakeAverageAndMergeCells(ByVal nofwell As Integer)
    Dim t_sum, daesoo_sum, gradient_sum, direction_sum As Double
    Dim i As Integer

    For i = 1 To nofwell
        t_sum = t_sum + Range("F" & (i + 3)).value
        daesoo_sum = daesoo_sum + Range("I" & (i + 3)).value
        direction_sum = direction_sum + Range("K" & (i + 3)).value
        
        ' 2025/03/10 --> ABS Gradient
        gradient_sum = gradient_sum + Abs(Range("M" & (i + 3)).value)
    Next i
    
    
    Cells(4, "g").value = Round(t_sum / nofwell, 4)
    Cells(4, "g").numberFormat = "0.0000"
    
    Cells(4, "j").value = Round(daesoo_sum / nofwell, 1)
    Cells(4, "j").numberFormat = "0.0"
        
    Cells(4, "l").value = Round(direction_sum / nofwell, 1)
    Cells(4, "l").numberFormat = "0.0"
        
    Cells(4, "n").value = Round(gradient_sum / nofwell, 4)
    Cells(4, "n").numberFormat = "0.0000"
       
    Cells(4, "o").value = "���������"
    Cells(4, "h").value = 0.03
    
    Call Merge_Cells("d", nofwell)
    Call Merge_Cells("g", nofwell)
    Call Merge_Cells("j", nofwell)
    Call Merge_Cells("l", nofwell)
    Call Merge_Cells("n", nofwell)
    Call Merge_Cells("o", nofwell)
    Call Merge_Cells("h", nofwell)

End Sub


Sub Merge_Cells(cel As String, ByVal nofwell As Integer)

    Range(cel & CStr(4) & ":" & cel & CStr(nofwell + 3)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
End Sub

Sub FindMergedCellsRange()
    Dim mergedRange As Range
    Set mergedRange = ActiveSheet.Range("A1").MergeArea
    MsgBox "The range of merged cells is " & mergedRange.Address
End Sub



Sub UnmergeAllCells()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            cell.MergeCells = False
        End If
    Next cell
    
    Call SetBorderLine_Default
End Sub


Sub SetBorderLine_Default()

    Range("C4:O17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub



Sub DrawOutline()

    Application.ScreenUpdating = False
    
    Range("C3:O34").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B31").Select
    
    Application.ScreenUpdating = True
End Sub





Const WELL_BUFFER = 30


'********************************************************
' 2025/3/2
' in case of add unit, then max min value disappear
' so always show the max, min value
'********************************************************

Public ROI_Max As Double
Public ROI_Min As Double
Public LONGAXIS_Max As Double
Public LONGAXIS_Min As Double
Public gbIsFirstTime  As Boolean

Sub Test_NameManager()
    Dim acColumn, acRow As Variant
    
    acColumn = Split(Range("ip_motor_simdo").Address, "$")(1)
    acRow = Split(Range("ip_motor_simdo").Address, "$")(2)
    
    '  Row = ActiveCell.Row
    '  col = ActiveCell.Column
    
    Debug.Print acColumn, acRow
End Sub


Sub Write23_SummaryDevelopmentPotential()
' Groundwater Development Potential, ���ϼ����߰��ɷ�
    
    Range("D4").value = Worksheets(CStr(1)).Range("e17").value
    Range("e4").value = Worksheets(CStr(1)).Range("g14").value
    Range("f4").value = Worksheets(CStr(1)).Range("f19").value
    Range("g4").value = Worksheets(CStr(1)).Range("g13").value
    
    Range("h4").value = Worksheets(CStr(1)).Range("g19").value
    Range("i4").value = Worksheets(CStr(1)).Range("f21").value
    Range("j4").value = Worksheets(CStr(1)).Range("e21").value
    Range("k4").value = Worksheets(CStr(1)).Range("g21").value
    
    ' --------------------------------------------------------------------
    
    Range("D8").value = Worksheets(CStr(1)).Range("e17").value
    Range("e8").value = Worksheets(CStr(1)).Range("g14").value
    Range("f8").value = Worksheets(CStr(1)).Range("f19").value
    Range("g8").value = Worksheets(CStr(1)).Range("g13").value
    
    Range("h8").value = Worksheets(CStr(1)).Range("g19").value
    Range("i8").value = Worksheets(CStr(1)).Range("f21").value
    Range("j8").value = Worksheets(CStr(1)).Range("h19").value
    Range("k8").value = Worksheets(CStr(1)).Range("e21").value

End Sub


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>


'2024-7-31 , change left and right

Sub Write26_AquiferCharacterization(nofwell As Integer)
' ����� �м��� ����ä���� �м�
    Dim i, ip_Row, remainder As Integer
    Dim unit, rngString As String
    Dim values As Variant
    
    values = GetRowColumn("AggSum_26_AC")
    ip_Row = values(2)
    'ip_row = "12" �� String
    
    rngString = "D" & ip_Row & ":K" & (CInt(ip_Row) + WELL_BUFFER - 1)
    
    
    Call EraseCellData(rngString)
        
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(11 + i, "d"), Cells(11 + i, "j")), True)
        Else
                Call BackGroundFill(Range(Cells(11 + i, "d"), Cells(11 + i, "j")), False)
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "D").value = "W-" & CStr(i)
        ' �ɵ�
        Cells(11 + i, "E").value = Worksheets("Well").Cells(i + 3, 8).value
        ' �����
        Cells(11 + i, "F").value = Worksheets("Well").Cells(i + 3, 10).value
        
        ' �ڿ�����
        Cells(11 + i, "G").value = Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "G").numberFormat = "0.00"
        
        ' ��������
        Cells(11 + i, "H").value = Worksheets(CStr(i)).Range("c21").value
        Cells(11 + i, "H").numberFormat = "0.00"
        
         '�������Ϸ�
        Cells(11 + i, "I").value = Worksheets(CStr(i)).Range("c21").value - Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "I").numberFormat = "0.00"
        
        ' ���������
        Cells(11 + i, "J").value = Worksheets(CStr(i)).Range("E7").value
        Cells(11 + i, "J").numberFormat = "0.0000"
        
        ' �������
        Cells(11 + i, "K").value = Worksheets(CStr(i)).Range("G7").value
        Cells(11 + i, "K").numberFormat = "0.0000000"
    Next i
End Sub


'2024-7-31 , change left and right
Sub Write26_Right_AquiferCharacterization(nofwell As Integer)
' ����� �м��� ����ä���� �м�
    Dim i, ip_Row, remainder As Integer
    Dim unit, rngString As String
    Dim values As Variant
    
    values = GetRowColumn("AggSum_26_RightAC")
    ip_Row = values(2)
    
    rngString = "L" & ip_Row & ":S" & (ip_Row + WELL_BUFFER - 1)
    
    Call EraseCellData(rngString)
            
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(11 + i, "L"), Cells(11 + i, "S")), True)
        Else
                Call BackGroundFill(Range(Cells(11 + i, "L"), Cells(11 + i, "S")), False)
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "M").value = "W-" & CStr(i)
        ' �ɵ�
        Cells(11 + i, "N").value = Worksheets("Well").Cells(i + 3, 8).value
        ' �����
        Cells(11 + i, "O").value = Worksheets("Well").Cells(i + 3, 10).value
        
        ' �ڿ�����
        Cells(11 + i, "P").value = Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "P").numberFormat = "0.00"
        
        ' ��������
        Cells(11 + i, "Q").value = Worksheets(CStr(i)).Range("c21").value
        Cells(11 + i, "Q").numberFormat = "0.00"
        
        '�������Ϸ�
'        Cells(11 + i, "Q").value = Worksheets(CStr(i)).Range("c21").value - Worksheets(CStr(i)).Range("c20").value
'        Cells(11 + i, "Q").NumberFormat = "0.00"
        
        ' ���������
        Cells(11 + i, "R").value = Worksheets(CStr(i)).Range("E7").value
        Cells(11 + i, "R").numberFormat = "0.0000"
         
        ' �������
        Cells(11 + i, "S").value = Worksheets(CStr(i)).Range("G7").value
        Cells(11 + i, "S").numberFormat = "0.0000000"
    Next i
End Sub

'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>



Sub Write_RadiusOfInfluence(nofwell As Integer)
' �������ݰ�

    Dim i, ip_Row, remainder As Integer
    Dim unit, rngString01, rngString02 As String
    Dim values As Variant
    
    Dim ws As Worksheet
    Dim rng As Range
        
    Set ws = ThisWorkbook.Sheets("AggSum")
    Set rng_ROI = ws.Range("E45:E74")
    Set rng_JANGCHOOK = ws.Range("F45:F74")
    
        
    values = GetRowColumn("AggSum_ROI")
    ip_Row = values(2)
    
    rngString01 = "D" & ip_Row & ":G" & (ip_Row + WELL_BUFFER - 1)
    rngString02 = "M" & ip_Row & ":O" & (ip_Row + WELL_BUFFER - 1)
    
    
    Call EraseCellData(rngString01)
    Call EraseCellData(rngString02)
        
        
    If gbIsFirstTime = False And Sheets("AggSum").CheckBox1.value = True Then
        ROI_Max = Range("R48").value
        ROI_Min = Range("R49").value
        LONGAXIS_Max = Range("R51").value
        LONGAXIS_Min = Range("R52").value
        
        gbIsFirstTime = True
    End If
        
        
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        ' WellNum
        Cells(ip_Row - 1 + i, "D").value = "W-" & CStr(i)
        ' �������ݰ�, �̰��� ������ ���� �ٸ���,
        ' �ϴ��� �ִ밪, shultz, webber, jcob �� �ִ밪�� �����ϴ°����� �Ѵ�.
        ' �׸��� �ʿ��� �κ���, �Ŀ� �߰������ش�.
        Cells(ip_Row - 1 + i, "E").value = Worksheets(CStr(i)).Range("H9").value & unit
        Cells(ip_Row - 1 + i, "F").value = Worksheets(CStr(i)).Range("K6").value & unit
        Cells(ip_Row - 1 + i, "G").value = Worksheets(CStr(i)).Range("K7").value & unit
        
        
        '����ݰ��� �ִ�, �ּ�, ��հ��� �߰����ش�.
        Cells(ip_Row - 1 + i, "M").value = Worksheets(CStr(i)).Range("H9").value & unit
        Cells(ip_Row - 1 + i, "N").value = Worksheets(CStr(i)).Range("H10").value & unit
        Cells(ip_Row - 1 + i, "O").value = Worksheets(CStr(i)).Range("H11").value & unit
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "d"), Cells(ip_Row - 1 + i, "g")), True)
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "m"), Cells(ip_Row - 1 + i, "o")), True)
        Else
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "d"), Cells(ip_Row - 1 + i, "j")), False)
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "m"), Cells(ip_Row - 1 + i, "o")), False)
        End If
    Next i
    
    If Sheets("AggSum").CheckBox1.value = True Then
        Range("R48").value = ROI_Max
        Range("R49").value = ROI_Min
        Range("R51").value = LONGAXIS_Max
        Range("R52").value = LONGAXIS_Min
    Else
        ROI_Max = Application.WorksheetFunction.max(rng_ROI)
        ROI_Min = Application.WorksheetFunction.min(rng_ROI)
        LONGAXIS_Max = Application.WorksheetFunction.max(rng_JANGCHOOK)
        LONGAXIS_Min = Application.WorksheetFunction.min(rng_JANGCHOOK)
                
        Range("R48").value = ROI_Max
        Range("R49").value = ROI_Min
        Range("R51").value = LONGAXIS_Max
        Range("R52").value = LONGAXIS_Min
    End If
    
    
End Sub


Sub Write_DrasticIndex(nofwell As Integer)
' ���ƽ �ε���
    Dim i, ip_Row, remainder As Integer
    Dim unit, rngString As String
    Dim values As Variant
    
    values = GetRowColumn("AggSum_DI")
    ip_Row = values(2)
    
    rngString = "I" & values(2) & ":K" & (values(2) + WELL_BUFFER - 1)
    Call EraseCellData(rngString)
    
    For i = 1 To nofwell
        ' WellNum
        Cells(ip_Row - 1 + i, "I").value = "W-" & CStr(i)
        Cells(ip_Row - 1 + i, "J").value = Worksheets(CStr(i)).Range("k30").value
        Cells(ip_Row - 1 + i, "K").value = Worksheets(CStr(i)).Range("k31").value
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "i"), Cells(ip_Row - 1 + i, "k")), True)
        Else
                Call BackGroundFill(Range(Cells(ip_Row - 1 + i, "i"), Cells(ip_Row - 1 + i, "k")), False)
        End If
        
    Next i
End Sub


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>


Sub TestColumnLetter()

    ' ColumnNumberToLetter
    ' ColumnLetterToNumber
    
    Debug.Print ColumnLetterToNumber("D")
    Debug.Print ColumnLetterToNumber("AG")
    ' 4
    ' 33
    ' 33 = 4 + 30 - 1

End Sub

Sub Write_Data(nofwell As Integer, category As String, numberFormatString As String, rangeCell As String, unitSuffix As String)
    ' Generalized subroutine to write data based on the category
    Dim i, ip_Row As Integer
    Dim unit, rngString As String
    Dim values As Variant
    Dim startCol, endCol As String

    values = GetRowColumn(category)
    ip_Row = values(2)

    startCol = values(1)
    endCol = ColumnNumberToLetter(ColumnLetterToNumber(startCol) + WELL_BUFFER - 1)

    rngString = startCol & ip_Row & ":" & endCol & (ip_Row + 1)
    Call EraseCellData(rngString)

    If Sheets("AggSum").CheckBox1.value = True Then
        unit = unitSuffix
    Else
        unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip_Row, (i + 3)).value = "W-" & CStr(i)
        With Cells(ip_Row + 1, (i + 3))
            ' 2025-3-30, change numberformat
            .value = format(Worksheets(CStr(i)).Range(rangeCell).value, numberFormatString) & unit
        End With
    Next i
End Sub

Sub Write_WaterIntake(nofwell As Integer)
'  Sheets("drastic").Range("a16").value - unit m^3/day
' �����ȹ��

    Write_Data nofwell, "AggSum_Intake", "#,##0.0", "C15", Sheets("drastic").Range("a16").value
End Sub

Sub Write_DiggingDepth(nofwell As Integer)
' �����ɵ�

    Write_Data nofwell, "AggSum_Simdo", "#,##0", "C7", " m"
End Sub

Sub Write_MotorPower(nofwell As Integer)
' ���͸���
    
    Write_Data nofwell, "AggSum_MotorHP", "#,##0.0", "C17", " Hp"
End Sub


Sub Write_NaturalLevel(nofwell As Integer)
' 2024,3,4 �ڿ�����
' �ڿ�����, �������� �ִ밪 �ּҰ��� FX ���� ������ �´�.

    Dim ip_Row As Integer
    Dim values As Variant
    

    values = GetRowColumn("AggSum_NaturalLevel")
    ip_Row = values(2)


    Write_Data nofwell, "AggSum_NaturalLevel", "#,##0.00", "C20", " m"
    
    
    Cells(ip_Row - 1, "E").value = Application.WorksheetFunction.max(Sheets("YangSoo").Range("B5:B37"))
    Cells(ip_Row - 1, "F").value = Application.WorksheetFunction.min(Sheets("YangSoo").Range("B5:B37"))
    
    ' Debug.Print "Range(""D99"").value :", "", ConvertToDouble(Range("D99").value), ""
End Sub


Sub Write_StableLevel(nofwell As Integer)
' 2024,3,4 ��������
' �ڿ�����, �������� �ִ밪 �ּҰ��� FX ���� ������ �´�.
    
    Dim ip_Row As Integer
    Dim values As Variant
    
    values = GetRowColumn("AggSum_StableLevel")
    ip_Row = values(2)

    Write_Data nofwell, "AggSum_StableLevel", "#,##0.00", "C21", " m"
    
    Cells(ip_Row - 1, "E").value = Application.WorksheetFunction.max(Sheets("YangSoo").Range("C5:C37"))
    Cells(ip_Row - 1, "F").value = Application.WorksheetFunction.min(Sheets("YangSoo").Range("C5:C37"))
End Sub


Sub Write_MotorTochool(nofwell As Integer)
' ���ⱸ��

    Write_Data nofwell, "AggSum_ToChool", "#,##0", "C19", " mm"
End Sub

Sub Write_MotorSimdo(nofwell As Integer)
' ���ͽɵ�

    Write_Data nofwell, "AggSum_MotorSimdo", "#,##0", "C18", " m"
End Sub

Sub Write_WellDiameter(nofwell As Integer)
' ��������

    Write_Data nofwell, "AggSum_WellDiameter", "#,##0", "C8", " mm"
End Sub

Sub Write_CasingDepth(nofwell As Integer)
' ���̽̽ɵ�
    Write_Data nofwell, "AggSum_CasingDepth", "#,##0", "C9", " m"
End Sub


Function CheckDrasticIndex(val As Integer) As String
    
    Dim value As Integer
    Dim result As String
    
    Select Case val
        Case Is <= 100
            result = "�ſ쳷��"
        Case Is <= 120
            result = "����"
        Case Is <= 140
            result = "��������"
        Case Is <= 160
            result = "�߰�����"
        Case Is <= 180
            result = "����"
        Case Else
            result = "�ſ����"
    End Select
    
    CheckDrasticIndex = result
End Function


Sub Check_DI()
' Drastic Index �� ������ �߷��� ...

    Dim i, ip_Row, ip_Column As Integer
    Dim unit, rngString01 As String
    Dim values As Variant
    
    values = GetRowColumn("AggSum_Statistic_DrasticIndex")
    
    ip_Column = ColumnLetterToNumber(values(1))
    ip_Row = values(2)
    
    Range(ColumnNumberToLetter(ip_Column + 1) & ip_Row).value = CheckDrasticIndex(Range(ColumnNumberToLetter(ip_Column) & ip_Row))
    Range(ColumnNumberToLetter(ip_Column + 1) & (ip_Row + 1)).value = CheckDrasticIndex(Range(ColumnNumberToLetter(ip_Column) & (ip_Row + 1)))

End Sub


Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:02"), "Popup_CloseUserForm"
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
   
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


Function GetProvince_Dictionary(city As String) As String
    Dim cityProvinceMap As Object
    Set cityProvinceMap = CreateObject("Scripting.Dictionary")
    
    ' ��û��
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "��õ", "��û��"
    cityProvinceMap.Add "û��", "��û��"
    cityProvinceMap.Add "��ǳ��", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "�ݻ�", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "�ο�", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "õ��", "��û��"
    cityProvinceMap.Add "ȫ��", "��û��"

    ' ������
    cityProvinceMap.Add "���ǻ�", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "��ȭ", "������"
    cityProvinceMap.Add "��ɵ�", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "����õ", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "����", "������"

    ' ������
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "�����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "�ϰ���", "������"
    cityProvinceMap.Add "����õ", "������"
    cityProvinceMap.Add "��ô", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "������", "������"
    cityProvinceMap.Add "ö��", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "�¹�", "������"
    cityProvinceMap.Add "ȫõ", "������"

    ' ����
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "��â", "����"
    cityProvinceMap.Add "��â��", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�ξ�", "����"
    cityProvinceMap.Add "��â��", "����"
    cityProvinceMap.Add "�ӽ�", "����"
    cityProvinceMap.Add "���", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "��õ", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "�ϵ�", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�־�", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "÷ö��", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "�س�", "����"
    cityProvinceMap.Add "��굵", "����"

    ' ���
    cityProvinceMap.Add "�뱸", "���"
    cityProvinceMap.Add "�뱸(��)", "���"
    cityProvinceMap.Add "���", "���"
    cityProvinceMap.Add "�λ�", "���"
    cityProvinceMap.Add "���ֽ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��ȭ", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�ȵ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��õ", "���"
    cityProvinceMap.Add "�︪��", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�Ǽ�", "���"
    cityProvinceMap.Add "û�۱�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��â", "���"
    cityProvinceMap.Add "���ؽ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�о�", "���"
    cityProvinceMap.Add "��â��", "���"
    cityProvinceMap.Add "��û", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�Ƿɱ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "â��", "���"
    cityProvinceMap.Add "�뿵", "���"
    cityProvinceMap.Add "�Ծ籺", "���"
    cityProvinceMap.Add "��õ", "���"

    ' ���ֵ�
    cityProvinceMap.Add "���", "���ֵ�"
    cityProvinceMap.Add "������", "���ֵ�"
    cityProvinceMap.Add "����", "���ֵ�"
    cityProvinceMap.Add "����2", "���ֵ�"
    cityProvinceMap.Add "������", "���ֵ�"
    
    ' Return the province if found
    If cityProvinceMap.Exists(city) Then
        GetProvince_Dictionary = cityProvinceMap(city)
    Else
        GetProvince_Dictionary = "Not in list"
    End If
    
    ' Clean up
    Set cityProvinceMap = Nothing
End Function



Function GetProvince_Vlookup(city As String) As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim result As Variant
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
    Set rng = ws.Range("A1:B100") ' Adjust the range to match your data

    ' Use VLOOKUP to find the province
    result = Application.WorksheetFunction.VLookup(city, rng, 2, False)
    
    ' Return the result or handle the error
    If IsError(result) Then
        GetProvince_Vlookup = "Not in list"
    Else
        GetProvince_Vlookup = result
    End If
End Function


Function GetProvince_Collection(city As String) As String
    Dim cityProvinceArray As Collection
    Dim cityProvince As Variant
    Dim i As Integer
    
    ' Initialize the collection with city-province pairs
    Set cityProvinceArray = New Collection
    
    ' Add city-province pairs to the collection
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("��õ", "��û��")
    cityProvinceArray.Add Array("û��", "��û��")
    cityProvinceArray.Add Array("��ǳ��", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("�ݻ�", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("�ο�", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("õ��", "��û��")
    cityProvinceArray.Add Array("ȫ��", "��û��")
    cityProvinceArray.Add Array("���ǻ�", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("��ȭ", "������")
    cityProvinceArray.Add Array("��ɵ�", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("����õ", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("�����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("�ϰ���", "������")
    cityProvinceArray.Add Array("����õ", "������")
    cityProvinceArray.Add Array("��ô", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("������", "������")
    cityProvinceArray.Add Array("ö��", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("�¹�", "������")
    cityProvinceArray.Add Array("ȫõ", "������")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("��â", "����")
    cityProvinceArray.Add Array("��â��", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�ξ�", "����")
    cityProvinceArray.Add Array("��â��", "����")
    cityProvinceArray.Add Array("�ӽ�", "����")
    cityProvinceArray.Add Array("���", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("��õ", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("�ϵ�", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�־�", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("÷ö��", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("�س�", "����")
    cityProvinceArray.Add Array("��굵", "����")
    cityProvinceArray.Add Array("�뱸", "���")
    cityProvinceArray.Add Array("�뱸(��)", "���")
    cityProvinceArray.Add Array("���", "���")
    cityProvinceArray.Add Array("�λ�", "���")
    cityProvinceArray.Add Array("���ֽ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��ȭ", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�ȵ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��õ", "���")
    cityProvinceArray.Add Array("�︪��", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�Ǽ�", "���")
    cityProvinceArray.Add Array("û�۱�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��â", "���")
    cityProvinceArray.Add Array("���ؽ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�о�", "���")
    cityProvinceArray.Add Array("��â��", "���")
    cityProvinceArray.Add Array("��û", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�Ƿɱ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("â��", "���")
    cityProvinceArray.Add Array("�뿵", "���")
    cityProvinceArray.Add Array("�Ծ籺", "���")
    cityProvinceArray.Add Array("��õ", "���")
    cityProvinceArray.Add Array("���", "���ֵ�")
    cityProvinceArray.Add Array("������", "���ֵ�")
    cityProvinceArray.Add Array("����", "���ֵ�")
    cityProvinceArray.Add Array("����2", "���ֵ�")
    cityProvinceArray.Add Array("������", "���ֵ�")
    
    ' Loop through the collection to find the city and return the corresponding province
    For Each cityProvince In cityProvinceArray
        If cityProvince(0) = city Then
            GetProvince_Collection = cityProvince(1)
            Exit Function
        End If
    Next cityProvince
    
    ' If city is not found in the collection, return "Not in list"
    GetProvince_Collection = "Not in list"
End Function



Sub importRainfall()
    Dim myArray As Variant
    Dim rng As Range

    Select Case UCase(Range("T6").value)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
    End Select

    Dim indexString As String
    indexString = "data_" & UCase(Range("T6").value)

    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "���û"
End Sub

'
' 2025/03/02 �泲���� ��ư �߰�
'
'

Sub importRainfall_button(ByVal AREA As String)
    Dim myArray As Variant
    Dim rng As Range

    Dim indexString As String
    indexString = "data_" & UCase(AREA)


    Select Case UCase(AREA)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
            
        Case "BORYUNG"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "DAEJEON"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "SEOSAN"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "BUYEO"
            Range("S5").value = "��û��"
            Range("T5").value = "�ο�"
        
        Case "CHEONAN"
            Range("S5").value = "��û��"
            Range("T5").value = "õ��"
        
        Case "CHEONGJU"
            Range("S5").value = "��û��"
            Range("T5").value = "û��"
        
        Case "GEUMSAN"
            Range("S5").value = "��û��"
            Range("T5").value = "�ݻ�"
             
        
        Case "SEOUL"
            Range("S5").value = "������"
            Range("T5").value = "����"
            
        Case "SUWON"
            Range("S5").value = "������"
            Range("T5").value = "����"
            
        Case "INCHEON"
            Range("S5").value = "������"
            Range("T5").value = "��õ"
            
    End Select



    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "���û"
End Sub




Function GetProvince_Case(city As String) As String
    Select Case city
        ' ��û��
        Case "����", "��õ", "û��", "��ǳ��", "����", "����", "�ݻ�", "����", "�ο�", "����", "õ��", "ȫ��"
            GetProvince_Case = "��û��"
        ' ������
        Case "���ǻ�", "����", "��ȭ", "��ɵ�", "��õ", "����õ", "����", "����", "��õ", "����"
            GetProvince_Case = "������"
        ' ������
        Case "����", "�����", "����", "�ϰ���", "����õ", "��ô", "����", "����", "����", "����", "������", "ö��", "��õ", "�¹�", "ȫõ"
            GetProvince_Case = "������"
        ' ����
        Case "����", "��â", "��â��", "����", "����", "�ξ�", "��â��", "�ӽ�", "���", "����", "����", "������", "����", "�����", "����", "����", "������", "��õ", "����", "������", "�ϵ�", "����", "�־�", "����", "÷ö��", "������", "�س�", "��굵"
            GetProvince_Case = "����"
        ' ���
        Case "�뱸", "�뱸(��)", "���", "�λ�", "���ֽ�", "����", "����", "��ȭ", "����", "�ȵ�", "����", "����", "��õ", "�︪��", "����", "�Ǽ�", "û�۱�", "����", "����", "��â", "���ؽ�", "����", "�о�", "��â��", "��û", "����", "�Ƿɱ�", "����", "â��", "�뿵", "�Ծ籺", "��õ"
            GetProvince_Case = "���"
        ' ���ֵ�
        Case "���", "������", "����", "����2", "������"
            GetProvince_Case = "���ֵ�"
        ' Default case
        Case Else
            GetProvince_Case = "Not in list"
    End Select
End Function



Sub ResetWeatherData(ByVal AREA As String)

    Dim Province As String
    
    Sheets("All").Activate
'    Range("S5") = "��û��"
'    Range("T5") = "û��"
    
    Province = GetProvince_Case(AREA)

    If CheckSubstring(Province, "Not in list") Then
        Popup_MessageBox (" Province is Not in list .... ")
        Exit Sub
    End If

    Range("S5") = Province
    Range("T5") = AREA
    
    
    Popup_MessageBox ("Clear Contents")
    Range("b5:n34").ClearContents
    
    Popup_MessageBox (" Load 30 year Weather Data ")
    Call modProvince.importRainfall
    

End Sub


Sub test()

    Call ResetWeatherData("����")

End Sub

'This Module is Empty 
Sub ImportEachWell(ByVal well_no As Integer)
    ' well_no -- well number
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    Dim nl, sl, DeltaS As Double
    Dim Casing As Integer
    Dim wsYangSoo As Worksheet

    ' Set the well number
    i = well_no
    
    ' Reference the YangSoo worksheet
    Set wsYangSoo = Worksheets("YangSoo")
    
    ' Turn off additional processes or features
    BaseData_ETC_02.TurnOffStuff
    
    ' Read data from the worksheet
    DeltaS = wsYangSoo.Cells(4 + i, "L").value
    nl = wsYangSoo.Cells(4 + i, "B").value
    sl = wsYangSoo.Cells(4 + i, "C").value
    Casing = wsYangSoo.Cells(4 + i, "J").value
    T1 = wsYangSoo.Cells(4 + i, "O").value
    S1 = wsYangSoo.Cells(4 + i, "R").value
    T2 = wsYangSoo.Cells(4 + i, "P").value
    S2 = wsYangSoo.Cells(4 + i, "S").value
    S3 = wsYangSoo.Cells(4 + i, "AQ").value
    Skin = wsYangSoo.Cells(4 + i, "Y").value
    RI1 = wsYangSoo.Cells(4 + i, "V").value
    RI2 = wsYangSoo.Cells(4 + i, "W").value
    RI3 = wsYangSoo.Cells(4 + i, "X").value
    
    ' Calculate the effective radius
    ir = GetEffectiveRadiusFromFX(i)
    
    ' Set the values in the target worksheet
    SetCellValueAndFormat Range("C20"), nl, "0.00"
    SetCellValueAndFormat Range("C21"), sl, "0.00"
    SetCellValueAndFormat Range("C10"), 5, "0"
    SetCellValueAndFormat Range("C11"), Casing - 5, "0"
    SetCellValueAndFormat Range("G6"), S3, "0.00"
    SetCellValueAndFormat Range("E5"), T1, "0.0000"
    SetCellValueAndFormat Range("E6"), T2, "0.0000"
    SetCellValueAndFormat Range("G5"), S2, "0.0000000"
    SetCellValueAndFormat Range("G4"), S1, "0.00000"
    SetCellValueAndFormat Range("H5"), Skin, "0.0000"
    SetCellValueAndFormat Range("H6"), ir, "0.0000"
    SetCellValueAndFormat Range("E10"), RI1, "0.0"
    SetCellValueAndFormat Range("F10"), RI2, "0.0"
    SetCellValueAndFormat Range("G10"), RI3, "0.0"
    SetCellValueAndFormat Range("C23"), Round(DeltaS, 2), "0.00"
    
    ' Turn on additional processes or features
    BaseData_ETC_02.TurnOnStuff
End Sub

' Helper function to set cell value and format
Sub SetCellValueAndFormat(cell As Range, value As Variant, format As String)
    cell.value = value
    cell.numberFormat = format
End Sub


Sub ImportWellSpecFX(ByVal well_no As Integer)
'
' well_no -- well number
'
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    
    ' nl : natural level, sl : stable level
    ' s3 - Recover Test �� S��
    
    Dim nl, sl, DeltaS As Double
    Dim Casing As Integer
    Dim wsYangSoo As Worksheet
    
    i = well_no
    Set wsYangSoo = Worksheets("YangSoo")
    BaseData_ETC_02.TurnOffStuff
    
    ' delta s : ����1���� ��������
    DeltaS = wsYangSoo.Cells(4 + i, "L").value
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
    nl = wsYangSoo.Cells(4 + i, "B").value
    sl = wsYangSoo.Cells(4 + i, "C").value
    Casing = wsYangSoo.Cells(4 + i, "J").value
    
    
    T1 = wsYangSoo.Cells(4 + i, "O").value
    S1 = wsYangSoo.Cells(4 + i, "R").value
    T2 = wsYangSoo.Cells(4 + i, "P").value
    S2 = wsYangSoo.Cells(4 + i, "S").value
    S3 = wsYangSoo.Cells(4 + i, "AQ").value
    
    ' ��Ų���
    Skin = wsYangSoo.Cells(4 + i, "Y").value
    
    ' yangsoo radius of influence
    RI1 = wsYangSoo.Cells(4 + i, "V").value  ' schultze
    RI2 = wsYangSoo.Cells(4 + i, "W").value  ' webber
    RI3 = wsYangSoo.Cells(4 + i, "X").value  ' jcob
    
    ' ��ȿ�칰�ݰ� , �������� ����
    ' ir = GetEffectiveRadius(WBNAME)
    ir = GetEffectiveRadiusFromFX(i)
    
      ' Set the values in the target worksheet
    SetCellValueAndFormat Range("C20"), nl, "0.00"
    SetCellValueAndFormat Range("C21"), sl, "0.00"
    SetCellValueAndFormat Range("C10"), 5, "0"
    SetCellValueAndFormat Range("C11"), Casing - 5, "0"
    SetCellValueAndFormat Range("G6"), S3, "0.00"
    SetCellValueAndFormat Range("E5"), T1, "0.0000"
    SetCellValueAndFormat Range("E6"), T2, "0.0000"
    SetCellValueAndFormat Range("G5"), S2, "0.0000000"
    SetCellValueAndFormat Range("G4"), S1, "0.00000"
    SetCellValueAndFormat Range("H5"), Skin, "0.0000"
    SetCellValueAndFormat Range("H6"), ir, "0.0000"
    SetCellValueAndFormat Range("E10"), RI1, "0.0"
    SetCellValueAndFormat Range("F10"), RI2, "0.0"
    SetCellValueAndFormat Range("G10"), RI3, "0.0"
    SetCellValueAndFormat Range("C23"), Round(DeltaS, 2), "0.00"

    BaseData_ETC_02.TurnOnStuff

End Sub





Private Sub ImportEachWell_OLD()
    Dim WkbkName As Object
    Dim WBNAME, cell1 As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, DeltaS As Double
    Dim Casing As Integer
    
    BaseData_ETC_02.TurnOffStuff
    
    i = 2
    ' Range("i1") = Workbooks.count
    ' WBName = Range("i2").value
    
    cell1 = Range("b2").value
    WBNAME = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBNAME) Then
        MsgBox "Please open the yangsoo data ! " & WBNAME
        Exit Sub
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
    
    Skin = Workbooks(WBNAME).Worksheets("SkinFactor").Range("G6").value
    
    ' yangsoo radius of influence
    ' ����, ����ݰ�
    RI1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C13").value
    ' ����, ����ݰ�
    RI2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C18").value
    ' ������, ����ݰ�
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
    
    '2024/6/10 move to s1 this G4 cell
    Range("G4") = S1
    
    
    Range("h5") = Skin 'skin coefficient
    Range("h6") = ir    'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(DeltaS, 2) 'deltas
    
    BaseData_ETC_02.TurnOnStuff
        
End Sub


Public Sub MakeFrameFromSelection()
' e30:i43

    Dim result As Variant
    Dim start_alpha, end_alpha As String
    Dim start_num, end_num, i  As Integer
    
    Dim rng_str As String
    
    rng_str = GetRangeStringFromSelection()
    rng_str = Replace(rng_str, "$", "")
    
    If Not CheckSubstring(rng_str, ":") Then
        Exit Sub
    End If

    Debug.Print rng_str
    Call MakeColorFrame(rng_str)
   
End Sub


Public Sub MakeColorFrame(ByVal str_rng As String)
' e30:i43

    Dim result As Variant
    Dim start_alpha, end_alpha As String
    Dim start_num, end_num, i  As Integer
    
    Range(str_rng).Select
    
    result = ExtractStringPartsUsingRegex(str_rng)
    start_alpha = result(0)
    end_alpha = result(1)
    start_num = result(2)
    end_num = result(3)
    
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    
    
    Range(start_alpha & start_num & ":" & end_alpha & start_num).Select
    
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    
    For i = start_num + 2 To end_num Step 2
    
        Range(start_alpha & i & ":" & end_alpha & i).Select
        
        With Selection.Interior
            .pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .themeColor = xlThemeColorDark1
            .TintAndShade = -4.99893185216834E-02
            .PatternTintAndShade = 0
        End With
    
    Next i
    
End Sub



Function ExtractStringPartsUsingRegex(inputString As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim result(3) As Variant
    
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the pattern
    ' The pattern ([a-zA-Z]+)(\d+) matches any letter(s) followed by any digit(s)
    regex.pattern = "([a-zA-Z]+)(\d+)"
    regex.Global = True
    
    ' Execute the regex pattern on the input string
    Set matches = regex.Execute(inputString)
    
    ' Check if the pattern matched the expected number of parts
    If matches.count >= 2 Then
        ' Store the results in the result array
        result(0) = matches(0).SubMatches(0) ' Start Letter
        result(1) = matches(1).SubMatches(0) ' End Letter
        result(2) = CLng(matches(0).SubMatches(1)) ' Start Number
        result(3) = CLng(matches(1).SubMatches(1)) ' End Number
    Else
        ' Handle the case where the pattern did not match
        result(0) = ""
        result(1) = ""
        result(2) = 0
        result(3) = 0
    End If
    
    ' Return the result array
    ExtractStringPartsUsingRegex = result
End Function

Sub TestExtractStringParts()
    Dim inputString As String
    Dim result As Variant
    
    ' Given string
    inputString = "e30:i43"
    
    ' Call the function and get the result
    result = ExtractStringPartsUsingRegex(inputString)
    
    ' Output the results
    Debug.Print "Start Letter: " & result(0)
    Debug.Print "End Letter: " & result(1)
    Debug.Print "Start Number: " & result(2)
    Debug.Print "End Number: " & result(3)
End Sub


Sub test()
    Call MakeColorFrame("o22:s33")
End Sub

'
' 2024-07-10 05:35:16
'
' Autohotkey HotKey String is based run ,,,,
'
' 2024/7/9 ��
' Refactoring, ImportWellSpec, ImportWellSpecFX, ImportEachWell
' change module name :  Module_ImportWellSpec --> mod_DuplicateWellSpec
'
' modFrame module --> MakeColorFrame, in given selection make boader Frame
'
' next 2 function generation ...
'
' Function GetER_ModeFX(ByVal well_no As Integer) As Integer
' Function GetEffectiveRadiusFromFX(ByVal well_no As Integer) As Double

' FX Sheet : add field, Address, Company, S3 field
' FX Sheet : detail tuning



Function CellContains(searchRange As Range, searchValue As String) As Boolean
    CellContains = InStr(1, LCase(searchRange.value), LCase(searchValue)) > 0
End Function


Function FindCellByLoopingPartialMatch() As String

    Dim ws As Worksheet
    Dim cell As Range
    Dim Address As String
     
     For Each cell In Range("A1:AZ1").Cells
        Debug.Print cell.Address, cell.value
    
        If CellContains(cell, "") Then
            Address = cell.Address
            Exit For
        End If
    Next
    FindCellByLoopingPartialMatch = Address
    
End Function

Sub test()
    Dim rg As String
    rg = FindCellByLoopingPartialMatch
    Debug.Print "the result: ", rg, Range(rg).value
End Sub

Option Explicit

' Constants for better maintainability
Private Const SUMMARY_START_ROW As Long = 80
Private Const TS_ANALYSIS_START_ROW As Long = 48
Private Const WELL_DATA_START_ROW As Long = 2

' Type definition for well parameters
Private Type WellParameters
    Q As Double
    Natural As Double
    Stable As Double
    Recover As Double
    
    Radius As Double
    DeltaS As Double
    DeltaH As Double
    
    DaeSoo As Double
    
    T1 As Double
    T2 As Double
    TA As Double
    
    K As Double
    Time As Double
    
    S1 As Double
    S2 As Double
    
    Schultz As Double
    Webber As Double
    Jcob As Double
    
    Skin As Double
    Er As Double
End Type

' Main import procedure
Sub GROK_ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim wsAggregate As Worksheet
    Dim wsYangSoo As Worksheet
    Dim wellCount As Integer
    Dim i As Integer
    Dim params As WellParameters

    Set wsAggregate = Worksheets("Aggregate2")
    Set wsYangSoo = Worksheets("YangSoo")
    wsAggregate.Activate

    wellCount = GetNumberOfWell()
    
    Call TurnOffStuff
    If Not isSingleWellImport Then
        Call ClearAllDataRanges
    End If

    For i = 1 To wellCount
        If ShouldProcessWell(i, singleWell, isSingleWellImport) Then
            params = GetWellParameters(wsYangSoo, i)

            Application.ScreenUpdating = False
            GROK_WriteAllWellData i, params, isSingleWellImport
            GROK_WriteSummaryTS i, params
            Application.ScreenUpdating = True
        End If
    Next i

    wsAggregate.Range("A1").Select
    Application.CutCopyMode = False
    Call TurnOnStuff
End Sub

' Helper function to determine if well should be processed
Private Function ShouldProcessWell(currentWell As Integer, singleWell As Integer, _
                                 isSingle As Boolean) As Boolean
    ShouldProcessWell = Not isSingle Or (isSingle And currentWell = singleWell)
End Function

' Get well parameters from YangSoo sheet
Private Function GetWellParameters(ws As Worksheet, wellIndex As Integer) As WellParameters
    Dim params As WellParameters
    Dim row As Long: row = 4 + wellIndex

    With params
        .Q = ws.Cells(row, "K").value
        .Natural = ws.Cells(row, "B").value
        .Stable = ws.Cells(row, "C").value
        .Recover = ws.Cells(row, "D").value
        
        .Radius = ws.Cells(row, "H").value
        .DeltaS = ws.Cells(row, "L").value
        .DeltaH = ws.Cells(row, "F").value
        
        .DaeSoo = ws.Cells(row, "N").value
        
        .T1 = ws.Cells(row, "O").value
        .T2 = ws.Cells(row, "P").value
        .TA = ws.Cells(row, "Q").value
        
        .Time = ws.Cells(row, "U").value
        .S1 = ws.Cells(row, "R").value
        .S2 = ws.Cells(row, "S").value
        .K = ws.Cells(row, "T").value
        
        .Schultz = ws.Cells(row, "V").value
        .Webber = ws.Cells(row, "W").value
        .Jcob = ws.Cells(row, "X").value
        
        .Skin = ws.Cells(row, "Y").value
        .Er = ws.Cells(row, "Z").value
    End With

    GetWellParameters = params
End Function

' Write all well data sections
Private Sub GROK_WriteAllWellData(wellIndex As Integer, params As WellParameters, _
                            isSingleImport As Boolean)
    Call GROK_WriteWellData(wellIndex, params, isSingleImport)
    Call GROK_WriteRadiusOfInfluence(wellIndex, params, isSingleImport)
    Call GROK_WriteTSAnalysis(wellIndex, params, isSingleImport)
    Call GROK_WriteRadiusOfInfluence(wellIndex, params, isSingleImport)
    Call GROK_WriteSkinFactor(wellIndex, params, isSingleImport)
    Call GROK_WriteRoiResult(wellIndex, params, isSingleImport)
End Sub

' Write summary T and S values
Private Sub GROK_WriteSummaryTS(wellIndex As Integer, params As WellParameters)
    Dim row As Long: row = SUMMARY_START_ROW + wellIndex - 1

    With Range("H" & row)
        .value = "W-" & wellIndex
        .Offset(0, 1).value = params.T2
        .Offset(0, 2).value = params.S2
    End With
End Sub

' Write well test data (Section 3-3, 3-4, 3-5)
Private Sub GROK_WriteWellData(wellIndex As Integer, params As WellParameters, _
                         isSingleImport As Boolean)
    Dim row As Long: row = WELL_DATA_START_ROW + wellIndex
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "C" & row & ":U" & row
    End If

    With Range("C" & row)
        ' Section 3-3
        .value = "W-" & wellIndex
        .Offset(0, 1).value = 2880
        .Offset(0, 2).value = params.Q
        .Offset(0, 9).value = params.Q
        .Offset(0, 3).value = params.Natural
        .Offset(0, 4).value = params.Stable
        .Offset(0, 5).value = params.Stable - params.Natural
        .Offset(0, 6).value = params.Radius
        .Offset(0, 7).value = params.DeltaS

        ' Section 3-4
        .Offset(0, 10).value = params.Radius
        .Offset(0, 11).value = params.Radius
        .Offset(0, 12).value = params.DaeSoo
        .Offset(0, 13).value = params.T1
        .Offset(0, 14).value = params.S1

        ' Section 3-5
        .Offset(0, 16).value = params.Stable
        .Offset(0, 17).value = params.Recover
        .Offset(0, 18).value = params.Stable - params.Recover
    End With

    Call ApplyBackgroundFill(Range(Cells(row, "C"), Cells(row, "J")), isEven)
    Call ApplyBackgroundFill(Range(Cells(row, "L"), Cells(row, "Q")), isEven)
    Call ApplyBackgroundFill(Range(Cells(row, "S"), Cells(row, "U")), isEven)
End Sub


' WriteData34_skinfactor
Private Sub GROK_WriteSkinFactor(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim values As Variant: values = GetRowColumn("agg2_34_skinfactor")
    Dim baseRow As Long: baseRow = values(2)
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "P" & baseRow & ":R" & baseRow
    End If

    With Range("P" & (baseRow + wellIndex - 1))
        .value = "W-" & wellIndex
        
        With .Offset(0, 1)
            .value = params.Skin: .numberFormat = "0.0000"
        End With
        
        With .Offset(0, 2)
            .value = params.Er: .numberFormat = "0.0000"
        End With
    End With

    Call ApplyBackgroundFill(Range(Cells(baseRow + wellIndex - 1, "P"), Cells(baseRow + wellIndex - 1, "R")), isEven)
End Sub


' WriteData38_ROI result
Private Sub GROK_WriteRoiResult(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim values As Variant: values = GetRowColumn("agg2_38_roi_result")
    Dim baseRow As Long: baseRow = values(2)
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "H" & baseRow & ":N" & baseRow
    End If

    With Range("H" & (baseRow + wellIndex - 1))
        .value = "W-" & wellIndex
        
        With .Offset(0, 1)
            .value = params.Schultz: .numberFormat = "0.0"
        End With
        
        With .Offset(0, 2)
            .value = params.Webber: .numberFormat = "0.0"
        End With
        
        
        With .Offset(0, 3)
            .value = params.Jcob: .numberFormat = "0.0"
        End With
        
        With .Offset(0, 4)
            .value = (params.Schultz + params.Webber + params.Jcob) / 3: .numberFormat = "0.0"
        End With
        
        With .Offset(0, 5)
            .value = Application.WorksheetFunction.max(params.Schultz, params.Webber, params.Jcob): .numberFormat = "0.0"
        End With
        
        With .Offset(0, 6)
            .value = Application.WorksheetFunction.min(params.Schultz, params.Webber, params.Jcob): .numberFormat = "0.0"
        End With
        
    End With

    Call ApplyBackgroundFill(Range(Cells(baseRow + wellIndex - 1, "H"), Cells(baseRow + wellIndex - 1, "N")), isEven)
End Sub




' Write radius of influence (Section 3-7)
Private Sub GROK_WriteRadiusOfInfluence(wellIndex As Integer, params As WellParameters, _
                                  isSingleImport As Boolean)
    Dim values As Variant: values = GetRowColumn("agg2_37_roi")
    Dim startRow As Long: startRow = values(2)
    Dim col As Long: col = 3 + wellIndex
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData ColumnNumberToLetter(col) & startRow & ":" & _
                     ColumnNumberToLetter(col) & (startRow + 6)
    End If

    With Cells(startRow, col)
        .Offset(0, 0).value = "W-" & wellIndex
        With .Offset(1, 0)
            .value = params.TA: .numberFormat = "0.0000"
        End With
        With .Offset(2, 0)
            .value = params.K: .numberFormat = "0.0000"
        End With
        With .Offset(3, 0)
            .value = params.S2: .numberFormat = "0.0000000"
        End With
        With .Offset(4, 0)
            .value = params.Time: .numberFormat = "0.0000"
        End With
        With .Offset(5, 0)
            .value = params.DeltaH: .numberFormat = "0.00"
        End With
        .Offset(6, 0).value = params.DaeSoo
    End With

    Call ApplyBackgroundFill(Range(Cells(startRow + 1, col), Cells(startRow + 6, col)), isEven)
End Sub

' Write TS analysis (Section 3-6)
Private Sub GROK_WriteTSAnalysis(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim values As Variant: values = GetRowColumn("agg2_36_surisangsoo")
    Dim baseRow As Long: baseRow = values(2) + (wellIndex - 1) * 3
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "C" & baseRow & ":F" & (baseRow + 2)
    End If


    Range("C" & baseRow).value = "W-" & wellIndex
        
    With Range("D" & baseRow)
        .Offset(0, 0).value = "���������"
        .Offset(1, 0).value = "����ȸ������"
        .Offset(2, 0).value = "����ġ"
    End With


    With Range("E" & baseRow)
        With .Offset(0, 0)
            .value = params.T1: .numberFormat = "0.0000"
        End With
        With .Offset(1, 0)
            .value = params.T2: .numberFormat = "0.0000"
        End With
        With .Offset(2, 0)
            .value = params.TA: .numberFormat = "0.0000": .Font.Bold = True
        End With
        With .Offset(0, 1)
            .value = params.S2: .numberFormat = "0.0000000"
        End With
        With .Offset(2, 1)
            .value = params.S2: .numberFormat = "0.0000000": .Font.Bold = True
        End With
    End With
    
    Call ApplyBackgroundFill(Range(Cells(baseRow, "C"), Cells(baseRow, "F")), isEven)
End Sub


' Helper sub to clear all data ranges
Private Sub ClearAllDataRanges()
    EraseCellData "C3:J33"
    EraseCellData "L3:Q33"
    EraseCellData "S3:U33"
    EraseCellData "D37:AH43"
    EraseCellData "D48:F137"
    EraseCellData "H48:N77"
    EraseCellData "P48:S77"
    EraseCellData "H80:J109"
End Sub

' Helper sub to apply background fill
Private Sub ApplyBackgroundFill(rng As Range, isEven As Boolean)
    Call BackGroundFill(rng, isEven)
End Sub

'
' 2025/3/4, Aggregate1 Refactoring
'
' Type definition for WellDataForAggregate1
'
Private Type WellDataOne
    Q As Double '�����
    Q1 As Double '1�ܰ� �����
    Qg As Double '��ä����
    Qh As Double '�Ѱ�����
    
    Ratio As Double
    
    Sd1 As Double ' 1�ܰ� �������Ϸ�
    Sd2 As Double ' 4�ܰ� �������Ϸ�
    
    C As Double
    B As Double
End Type

' Get well parameters from YangSoo sheet
Private Function GetWellData(wellIndex As Integer) As WellDataOne
    Dim params As WellDataOne
    Dim ws As Worksheet
    Dim row As Long: row = 4 + wellIndex
    
    
    Set ws = Worksheets("YangSoo")

    With params
        .Q = ws.Cells(row, "k").value
        .Qg = ws.Cells(row, "ab").value
        
        .Q1 = ws.Cells(row, "ac").value
        .Qh = ws.Cells(row, "aa").value
        
        .Ratio = ws.Cells(row, "ah").value
        
        .Sd1 = ws.Cells(row, "ad").value
        .Sd2 = ws.Cells(row, "ae").value
        
        .C = ws.Cells(row, "af").value
        .B = ws.Cells(row, "ag").value
    End With

    GetWellData = params
End Function


Sub ImportAggregateData(ByVal targetWell As Integer, ByVal isSingleWellMode As Boolean)
    ' Handles both single well and all wells import operations
    ' isSingleWellMode = True: Imports data for specified well only
    ' isSingleWellMode = False: Imports data for all wells

    Dim wellCount As Integer
    Dim wellIndex As Integer
    Dim wd As WellDataOne
    

    ' Initialize core variables
    wellCount = GetNumberOfWell()
    
    
    Sheets("Aggregate1").Activate

    Call TurnOffStuff
    ' Clear data ranges if importing all wells
    If Not isSingleWellMode Then
        ClearRange "G3:K35"
        ClearRange "Q3:S35"
        ClearRange "F43:I102"
    End If

    ' Process each well
    For wellIndex = 1 To wellCount
        If ShouldProcessWell(wellIndex, targetWell, isSingleWellMode) Then
            ' Fetch well data from YangSoo worksheet
           
            wd = GetWellData(wellIndex)
            
            ' Process data with optimizations disabled
            
            Call WriteWellSummary(wd, wellIndex, isSingleWellMode)
            Call WriteWaterIntake(wd, wellIndex, isSingleWellMode)
        End If
    Next wellIndex

    ' Clean up
    Application.CutCopyMode = False
    Range("L1").Select
    Call TurnOnStuff
End Sub

Private Sub WriteWellSummary(WellData As WellDataOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
    ' Writes well summary data to columns G:K and Q:S for a specific well
    ' Parameters:
    '   wellData: Structure containing well measurement data
    '   wellIndex: Index of the well being processed
    '   isSingleWellMode: Flag indicating single well (True) or all wells (False) operation
    
    Dim rowNumber As Integer
    Dim wellLabel As String
    
    ' Calculate target row and well identifier
    rowNumber = wellIndex + 2
    wellLabel = "W-" & wellIndex
    
    ' Clear existing data if in single well mode
    If isSingleWellMode Then
        ClearRange "G" & rowNumber & ":K" & rowNumber
        ClearRange "Q" & rowNumber & ":S" & rowNumber
    End If
    
    ' Write summary data using With blocks for efficiency
    With Range("G" & rowNumber)
        .value = wellLabel
        .Offset(0, 1).value = WellData.Qh
        .Offset(0, 2).value = WellData.Qg
        .Offset(0, 3).value = WellData.Q
        .Offset(0, 4).value = WellData.Ratio
    End With
    
    With Range("Q" & rowNumber)
        .value = wellLabel
        .Offset(0, 1).value = WellData.C
        .Offset(0, 2).value = WellData.B
    End With
    
    ' Apply alternating background formatting
    ApplyBackgroundFormatting rowNumber, "G", "K", wellIndex
    ApplyBackgroundFormatting rowNumber, "Q", "S", wellIndex
End Sub

Private Sub WriteWaterIntake(wd As WellDataOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
    ' Calculates and writes tentative water intake data starting at row 43

    Dim startRow As Integer
    Dim baseRow As Integer
    Dim values As Variant

    ' Get starting row from configuration
    values = GetRowColumn("Agg1_Tentative_Water_Intake")
    startRow = values(2)
    baseRow = startRow + (wellIndex - 1) * 2

    ' Clear specific rows if in single well mode
    If isSingleWellMode Then
        ClearRange "F" & baseRow & ":I" & (baseRow + 1)
    End If

    ' Write water intake data
    Cells(baseRow, "F").value = "W-" & CStr(wellIndex)
    Cells(baseRow, "G").value = wd.Q1
    Cells(baseRow, "H").value = wd.Sd2
    Cells(baseRow + 1, "H").value = wd.Sd1
    Cells(baseRow, "I").value = wd.Qg

    ' Apply background formatting
    ApplyBackgroundFormatting baseRow, "F", "I", wellIndex, 2
End Sub

Private Function ShouldProcessWell(ByVal currentIndex As Integer, ByVal targetWell As Integer, _
                                 ByVal isSingleWellMode As Boolean) As Boolean
    ' Determines if a well should be processed based on import mode
    ShouldProcessWell = Not isSingleWellMode Or (isSingleWellMode And currentIndex = targetWell)
End Function

Private Sub ApplyBackgroundFormatting(ByVal startRow As Integer, ByVal startCol As String, _
                                    ByVal endCol As String, ByVal wellIndex As Integer, _
                                    Optional ByVal rowSpan As Integer = 1)
    ' Applies alternating background colors to specified range
    Dim targetRange As Range
    Set targetRange = Range(Cells(startRow, startCol), Cells(startRow + rowSpan - 1, endCol))
    BackGroundFill targetRange, (wellIndex Mod 2 = 0)
End Sub

Private Sub ClearRange(ByVal rangeAddress As String)
    ' Clears content in specified range
    Range(rangeAddress).ClearContents
End Sub

'
' Refactor By User Defined Type
'
' Define User-Defined Type for Well Data


Private Type WellData
    Natural As Double
    Stable As Double
    Recover As Double
    
    DeltaH As Double
    Sw As Double
    Radius As Double
    
    Rw As Double
    WellDepth As Double
    Casing As Double
    
    Q As Double
    DeltaS As Double
    Hp As Double
    DaeSoo As Double
    
    T1 As Double
    T2 As Double
    TA As Double
    S1 As Double
    S2 As Double
    
    K As Double
    Time As Double
    
    Shultze As Double
    Webber As Double
    Jacob As Double
    Skin As Double
    
    Er As Double
    ER1 As Double
    ER2 As Double
    ER3 As Double
    
    Qh As Double
    Qg As Double
    Sd1 As Double
    Sd2 As Double
    Q1 As Double
    
    C As Double
    B As Double
    Ratio As Double
    
    T0 As Double
    S0 As Double
    ERMode As String
    Address As String
    Company As String
    
    S3 As Double
    Title As String
End Type

' Main procedure using UDT
Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String
    
    ' Get total number of wells
    nofwell = GetNumberOfWell()
    
    Call TurnOffStuff
    ' Determine range to clear based on import type
    If Not isSingleWellImport And singleWell = 999 Then
        rngString = "A5:AR37"
    Else
        rngString = "A" & (singleWell + 4) & ":AR" & (singleWell + 4)
    End If
    
    Call EraseCellData(rngString)
    
    ' Process each well
    For i = 1 To nofwell
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            Dim well As WellData
            well = ImportDataForWell(i)
            Call SetWellDataToSheet(i, well)
        End If
    Next i
    
    Call TurnOnStuff
End Sub

' Function to import well data into UDT
Private Function ImportDataForWell(ByVal wellIndex As Integer) As WellData
    Dim fName As String
    Dim wb As Workbook
    Dim well As WellData

    ' Construct filename and check if workbook is open
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Function
    End If

    Set wb = Workbooks(fName)

    
    ' Get worksheet references
    With wb
        Dim wsInput As Worksheet: Set wsInput = .Worksheets("Input")
        Dim wsSkinFactor As Worksheet: Set wsSkinFactor = .Worksheets("SkinFactor")
        Dim wsSafeYield As Worksheet: Set wsSafeYield = .Worksheets("SafeYield")

        ' Populate UDT with data
        With well
            .Natural = wsInput.Range("m48").value ' �ڿ�����
            .Stable = wsInput.Range("m49").value ' ��������
            .Recover = wsSkinFactor.Range("c10").value ' ȸ������
            .Sw = wsSkinFactor.Range("c11").value ' ȸ������

            .DeltaH = wsSkinFactor.Range("b16").value  ' �������Ϸ�
            .Radius = wsInput.Range("m44").value '200 , �칰����
            .Rw = wsSkinFactor.Range("e4").value '0.1

            .WellDepth = wsInput.Range("m45").value ' �����ɵ�
            .Casing = wsInput.Range("i52").value ' ���̽̽ɵ�
            .Q = wsInput.Range("m51").value '�����, ä����ȹ��
            .C = wsInput.Range("a31").value ' �칰�սǰ��
            .B = wsInput.Range("b31").value ' ������սǰ��

            .DeltaS = wsSkinFactor.Range("b4").value '���� 1�а� �������Ϸ�
            .Hp = wsInput.Range("i48").value  ' ���͸���
            .DaeSoo = wsSkinFactor.Range("c16").value ' ������β�

            .T1 = wsSkinFactor.Range("d5").value
            .T2 = wsSkinFactor.Range("h13").value
            .TA = wsSkinFactor.Range("d16").value
            .S1 = wsSkinFactor.Range("e10").value
            .S2 = wsSkinFactor.Range("i16").value

            .K = wsSkinFactor.Range("e16").value
            .Time = wsSkinFactor.Range("h16").value

            .Shultze = wsSkinFactor.Range("c13").value
            .Webber = wsSkinFactor.Range("c18").value
            .Jacob = wsSkinFactor.Range("c23").value

            .Skin = wsSkinFactor.Range("g6").value
            .Er = wsSkinFactor.Range("c8").value

            .ER1 = wsSkinFactor.Range("k8").value
            .ER2 = wsSkinFactor.Range("k9").value
            .ER3 = wsSkinFactor.Range("k10").value

            .Qh = wsInput.Range("d6").value ' �Ѱ�����
            .Qg = wsSafeYield.Range("b7").value  ' ��ä����

            .Sd1 = wsSafeYield.Range("b3").value '1�ܰ� ���Ϸ�
            .Sd2 = wsSafeYield.Range("b4").value '4�ܰ� ���Ϸ�
            .Q1 = wsSafeYield.Range("b2").value ' 1�ܰ� �����
            .Ratio = wsSafeYield.Range("b11").value

            .T0 = wsSkinFactor.Range("d4").value
            .S0 = wsSkinFactor.Range("f4").value

            .ERMode = wsSkinFactor.Range("h10").value
            .Address = wsInput.Range("i46").value
            .Company = wsInput.Range("i47").value

            .S3 = wsSkinFactor.Range("i13").value
            .Title = wsInput.Range("i44").value
        End With
    End With

    
    ImportDataForWell = well
End Function

' Procedure to set well data to worksheet
Private Sub SetWellDataToSheet(ByVal wellIndex As Integer, well As WellData)
    Dim row As Long: row = 4 + wellIndex
    Cells(row, 1).value = "W-" & wellIndex
    
    
    With well
        SetCellValue row, 2, .Natural, "0.00"
        SetCellValue row, 3, .Stable, "0.00"
        SetCellValue row, 4, .Recover, "0.00"
        SetCellValue row, 5, .Sw, "0.00"
        
        SetCellValue row, 6, .DeltaH, "0.00"
        SetCellValue row, 7, .Radius, "0."
        
        SetCellValue row, 8, .Rw, "0.000"
        SetCellValue row, 9, .WellDepth, "0"
        
        SetCellValue row, 10, .Casing, "0"
        SetCellValue row, 11, .Q, "0"
        SetCellValue row, 12, .DeltaS, "0.00"
        SetCellValue row, 13, .Hp, "0.0"
        SetCellValue row, 14, .DaeSoo, "0.00"
        
        SetCellValue row, 15, .T1, "0.0000"
        SetCellValue row, 16, .T2, "0.0000"
        SetCellValue row, 17, .TA, "0.0000"
        
        SetCellValue row, 18, .S1, "0.0000000"
        SetCellValue row, 19, .S2, "0.0000000"
        
        SetCellValue row, 20, .K, "0.0000"
        SetCellValue row, 21, .Time, "0.0000"
        SetCellValue row, 22, .Shultze, "0.00"
        SetCellValue row, 23, .Webber, "0.00"
        SetCellValue row, 24, .Jacob, "0.00"
        
        SetCellValue row, 25, .Skin, "0.0000"
        SetCellValue row, 26, .Er, "0.0000"
        SetCellValue row, 27, .Qh, "0"
        SetCellValue row, 28, .Qg, "0.00"
        
        SetCellValue row, 29, .Q1, "0.00"
        SetCellValue row, 30, .Sd1, "0.00"
        SetCellValue row, 31, .Sd2, "0.00"
        
        SetCellValue row, 32, .C, "0.00"
        SetCellValue row, 33, .B, "0.00"
        SetCellValue row, 34, .Ratio, "0.0%"
        
        SetCellValue row, 35, .T0, "0.0000"
        SetCellValue row, 36, .S0, "0.0000"
        
        SetCellValue row, 37, .ERMode, ""
        SetCellValue row, 38, .ER1, "0.0000"
        SetCellValue row, 39, .ER2, "0.0000"
        SetCellValue row, 40, .ER3, "0.0000"
        
        SetCellValue row, 41, .Address, ""
        SetCellValue row, 42, .Company, ""
        SetCellValue row, 43, .S3, "0.00"
        SetCellValue row, 44, .Title, ""
    End With
    
End Sub

' Helper procedure to set cell value and format
Private Sub SetCellValue(ByVal row As Long, ByVal col As Integer, ByVal value As Variant, ByVal numberFormat As String)
    With Cells(row, col)
        .value = value
        If Len(numberFormat) > 0 Then .numberFormat = numberFormat
    End With
End Sub

' Keep existing helper functions
Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True
      DoEvents
      Debug.Print String(30, vbCrLf)
   End If
   Debug.Print sPrintStr
End Sub

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim Er As String, R As String
    Er = ERMode
    R = Mid(Er, 5, 1)
    
    If R = "F" Then
        DetermineEffectiveRadius = 0 ' Assuming erRE0 = 0, adjust if different
    Else
        DetermineEffectiveRadius = val(R)
    End If
End Function

Function CheckFileExistence(filePath As String) As Boolean
    CheckFileExistence = (Dir(filePath) <> "")
End Function



Sub WriteFormula()
 
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double
    
    ' Save array to a file
    Dim filePath As String
    Dim FileNum As Integer
    
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    DeleteFileIfExists filePath
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
        
    Call MyDebug("Formula SkinFactor ... ", True)
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    ' ��Ų���
    Call FormulaSkinFactorAndER("SKIN", FileNum)
    
    ' ��ȿ�칰�ݰ�
    Call FormulaSkinFactorAndER("ER", FileNum)
    
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    Call FormulaChwiSoo(FileNum)
    ' 3-7, ���������
    
    Call FormulaRadiusOfInfluence(FileNum)
    ' �������ݰ�
        
    Close FileNum

End Sub



Sub FormulaSkinFactorAndER(ByVal Mode As String, ByVal FileNum As Integer)
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double
    
        
    Call MyDebug("Formula SkinFactor ... ", True)
    
    nofwell = GetNumberOfWell()
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        T = format(Cells(4 + i, "o").value, "0.0000")
        Q = Cells(4 + i, "k").value
        
        T0 = format(Cells(4 + i, "AI").value, "0.0000")
        S0 = format(Cells(4 + i, "AJ").value, "0.0000")
        S1 = Cells(4 + i, "R").value
                
        delta_s = format(Cells(4 + i, "l").value, "0.00")
        Radius = format(Cells(4 + i, "h").value, "0.000")
        Skin = format(Cells(4 + i, "y").value, "0.0000")
        Er = format(Cells(4 + i, "z").value, "0.0000")
        
        
        B = format(Cells(4 + i, "AG").value, "0.0000")
        ER1 = format(Cells(4 + i, "AL").value, "0.0000")
        ER2 = format(Cells(4 + i, "AM").value, "0.0000")
        ER3 = format(Cells(4 + i, "AN").value, "0.0000")
        
        
        Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
        ' ����� 1��
        Case erRE1
            
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
            formula2 = "erRE1, ����� 1��"
            
        ' ����� 2��
        Case erRE2
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
            formula2 = "erRE2, ����� 2��"
        ' ����� 3��
        Case erRE3
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~" & Radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
            formula2 = "erRE3, ����� 3��"
            
        Case Else
            ' ��Ų���
            formula1 = "W-" & i & "ȣ��~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {" & S0 & " TIMES  (" & Radius & " TIMES  " & Radius & ")} =`" & Skin
            ' ��ȿ�칰�ݰ�
            formula2 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & Radius & " TIMES e ^{-(" & Skin & ")} =" & Er & "m"
        End Select
        
        
        If Mode = "SKIN" Then
            Debug.Print formula1
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            
            Print #FileNum, formula1
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Else
            Debug.Print formula2
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                
            Print #FileNum, formula2
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    Next i

End Sub


Sub DeleteFileIfExists(filePath As String)
    If Len(Dir(filePath)) > 0 Then ' Check if file exists
        On Error Resume Next
        Kill filePath ' Attempt to delete the file
        
        On Error GoTo 0
        If Len(Dir(filePath)) > 0 Then
            MsgBox "Unable to delete file '" & filePath & "'. The file may be in use.", vbExclamation
        Else
            ' MsgBox "File '" & filePath & "' has been deleted.", vbInformation
            
        End If
    Else
        ' MsgBox "File '" & filePath & "' does not exist.", vbExclamation
    End If
End Sub



Sub FormulaChwiSoo(FileNum As Integer)
' 3-7, ���������

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim Q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        Q1 = Cells(4 + i, "ac").value
 
        S1 = format(Cells(4 + i, "ad").value, "0.00")
        S2 = format(Cells(4 + i, "ae").value, "0.00")
        res = format(Cells(4 + i, "ab").value, "0.00")
        
        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =165.52��/��
        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =`165.52��/��
        
        formula = "W-" & i & "ȣ��~~Q_{& 2} = " & Q1 & " TIMES ({" & S2 & "} over {" & S1 & "}) ^{2/3} = `" & res & " ��/��"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
      
End Sub


Sub FormulaRadiusOfInfluence(FileNum As Integer)
' �������ݰ�

    Call FormulaSUB_ROI("SCHULTZE", FileNum)
    Call FormulaSUB_ROI("WEBBER", FileNum)
    Call FormulaSUB_ROI("JCOB", FileNum)
    
End Sub




Sub FormulaSUB_ROI(Mode As String, FileNum As Integer)
  Dim formula1, formula2, formula3 As String
    ' ����, ����, ������
    
    Dim nofwell As String
    Dim i As Integer
    Dim Shultze, Webber, Jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        schultze = CStr(format(Cells(4 + i, "v").value, "0.0"))
        Webber = CStr(format(Cells(4 + i, "w").value, "0.0"))
        Jacob = CStr(format(Cells(4 + i, "x").value, "0.0"))
        
        T = CStr(format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(format(Cells(4 + i, "f").value, "0.00"))
        time_ = CStr(format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & schultze & "m"
        formula2 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & Webber & "`m"
        formula3 = "W-" & i & "ȣ��~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & Jacob & "m"
        
        
        Select Case Mode
            Case "SCHULTZE"
                Debug.Print formula1
                Print #FileNum, formula1
            
            Case "WEBBER"
                Debug.Print formula2
                Print #FileNum, formula2
                
            Case "JCOB"
                Debug.Print formula3
                Print #FileNum, formula3
        End Select
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
End Sub





