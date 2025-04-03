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
