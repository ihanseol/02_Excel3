Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").CheckBox_SingleColor.value = True
    Sheets("Well").CheckBox_GetChart.value = True
    Sheets("Recharge").cbCheSoo.value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub


