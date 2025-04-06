
Private Sub CommandButton_CollectData_Click()
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

Private Sub CommandButton_Formula_Click()
    ' Write Formula Button
       
       Call WriteFormula
    ' End of Write Formula Button
End Sub

Private Sub CommandButton_HideSheet_Click()
'Hide YangSoo

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton_SingleWell_Import_Click()
    'single well import
    
    Dim WellNumber  As Integer
    Dim WB_NAME As String
    
    ' 영수시험 데이터 파일이름, 불러오기
    WB_NAME = BaseData_ETC.GetOtherFileName
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        WellNumber = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (WellNumber)
    End If
    
    ' Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Call modAggFX_A.GetBaseDataFromYangSoo(WellNumber, True)

End Sub




'
'<><>><><><><><><>><><><><><><>><><><><><><>><><><><><><>><><><><><><>><><><><><><>><><><><><><>><><><><><><>><><><><>
'




