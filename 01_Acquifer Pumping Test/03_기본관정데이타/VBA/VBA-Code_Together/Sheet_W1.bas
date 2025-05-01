
Private Sub CommandButton4_Click()
    Call delete_allWhpaData
End Sub


Private Sub CommandButton_GetDirectionChar_Click()
    Dim angle As Integer
    Dim grade As String
    
    angle = getDirectionFromWellActiveSheet()

    Select Case angle
       Case 0 To 10
         grade = "����"
       
       Case 11 To 79
         grade = "�ϵ���"
       Case 80 To 100
         grade = "��"
       Case 101 To 169
         grade = "�ϼ���"
         
       Case 170 To 190
         grade = "����"
         
       Case 191 To 259
         grade = "������"
       
       Case 260 To 280
         grade = "����"
       
       Case 281 To 349
         grade = "������"
         
       Case 350 To 360
         grade = "����"
         
       Case Else
         grade = "Invalid Score"
     End Select

    Range("L13").value = grade
    
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


