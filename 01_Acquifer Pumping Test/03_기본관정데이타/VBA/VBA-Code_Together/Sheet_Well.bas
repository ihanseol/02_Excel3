Option Explicit

Private Sub CommandButton_AddWell_Click()
    BaseData_ETC_02.TurnOffStuff
    Call AddWell_CopyOneSheet
    BaseData_ETC_02.TurnOnStuff
End Sub

Private Sub CommandButton_Agg1_Click()
    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
End Sub

Private Sub CommandButton_Agg2_Click()
' Aggregate2 Button
' �����Լ� 2��

    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
End Sub

Private Sub CommandButton_Chart_Click()
    Sheets("AggChart").Visible = True
    Sheets("AggChart").Select
End Sub

Private Sub CommandButton_DeleteLast_Click()
    Call DeleteLast
End Sub

Private Sub CommandButton_Duplicate_Click()
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

Private Sub CommandButton_FX_Click()
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
End Sub

Private Sub CommandButton_Step_Click()
  Sheets("AggStep").Visible = True
  Sheets("AggStep").Select
End Sub

Private Sub CommandButton_Summary_Click()
    Sheets("AggSum").Visible = True
    Sheets("AggSum").Select
End Sub

Private Sub CommandButton_Water_Click()
    Sheets("water").Visible = True
    Sheets("water").Select
End Sub

Private Sub CommandButton_Whpa_Click()
    Sheets("aggWhpa").Visible = True
    Sheets("aggWhpa").Select
End Sub



Private Sub CommandButton_Jojung_Click()
    Call JojungButton
End Sub

Private Sub CommandButton_One_Click()
   Call Make_OneButton
End Sub

Private Sub CommandButton_PressAll_Click()
    Call PressAll_Button
End Sub

Private Sub CommandButton_SingleMain_Click()
' SingleWell Import
' Open FX Sheet, SingleWell Import, ImportMainWellPage

   
    Dim WellNumber  As Integer
    Dim WB_NAME As String
    
    WB_NAME = BaseData_ETC.GetOtherFileName
    
    If WB_NAME = "Empty" Then
        MsgBox " SingleWell Import, YangSoo WorkBook must be One ... "
        Exit Sub
    Else
        WellNumber = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (WellNumber)
    End If
    
    Call modWell.ImportSingleWell_Main(WellNumber)
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




Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub




