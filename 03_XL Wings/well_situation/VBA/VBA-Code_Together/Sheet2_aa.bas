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
    
    Call MakeFieldList
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
