
Private Sub CommandButton2_Click()
    UserFormSTime.Show
    ' Call frame_time_setting
    Call TimeSetting
    
    ActiveWindow.SmallScroll Down:=-66
    Range("O10").Select
End Sub

Private Sub frame_time_setting()
    Dim i As Integer
    Dim dStableTime As Integer
    
    Call initDictionary
    
    dStableTime = CInt(shW_LongTEST.ComboBox1.Value)
    MY_TIME = gDicStableTime(dStableTime)

End Sub

Private Sub CommandButton3_Click()
    Call set_daydifference
End Sub

Private Sub CommandButton4_Click()
    Call findAnswer_LongTest
End Sub

Private Sub CommandButton5_Click()
    Call resetValue
End Sub

Private Sub CommandButton6_Click()
    UserFormTS.Show
End Sub

Private Sub CommandButton7_Click()
    Call check_LongTest
End Sub



Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong, occur As Long
    
    
    Debug.Print ActiveSheet.name
    Call initDictionary
    
    If ActiveSheet.name <> "LongTest" Then
        Exit Sub
    End If
       
    If MY_TIME = 0 Then
        MY_TIME = initialize_myTime
        shW_LongTEST.ComboBox1.Value = gDicMyTime(MY_TIME)
    End If

'   gong = Val(CleanString(shInput.Range("J48").Value))
'   gong1 = "W-" & CStr(gong)
'   gong2 = shInput.Range("i54").Value
'
'   If gong1 <> gong2 Then
'        shInput.Range("i54").Value = gong1
'   End If
    
End Sub


