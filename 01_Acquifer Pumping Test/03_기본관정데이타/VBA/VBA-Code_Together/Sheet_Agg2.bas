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







