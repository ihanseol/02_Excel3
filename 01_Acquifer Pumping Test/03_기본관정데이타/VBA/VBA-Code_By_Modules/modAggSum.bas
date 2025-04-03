Attribute VB_Name = "modAggSum"

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

