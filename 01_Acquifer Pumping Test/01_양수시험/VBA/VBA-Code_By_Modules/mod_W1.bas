Attribute VB_Name = "mod_W1"

Sub Restore2880()
    Dim SaveFormula(5, 3) As String
    Dim i, j As Integer
    
    SaveFormula(0, 0) = "=$C$6+C21/1440"
    SaveFormula(1, 0) = "1680"
    SaveFormula(2, 0) = "=$D$20"
    SaveFormula(3, 0) = "=$E$20"
    SaveFormula(4, 0) = "=$F$20"
    
    SaveFormula(0, 1) = "=$C$6+C22/1440"
    SaveFormula(1, 1) = "1920"
    SaveFormula(2, 1) = "=$D$20"
    SaveFormula(3, 1) = "=$E$20"
    SaveFormula(4, 1) = "=$F$20"
    
    SaveFormula(0, 2) = "=$C$6+C23/1440"
    SaveFormula(1, 2) = "2880"
    SaveFormula(2, 2) = "=$D$20"
    SaveFormula(3, 2) = "=$E$20"
    SaveFormula(4, 2) = "=$F$20"
    
    For i = 0 To 2
        For j = 0 To 4
            Sheets("w1").Cells(21 + i, 2 + j).Formula = SaveFormula(j, i)
        Next j
    Next i
    
End Sub

Sub Delete2880()

    Dim i, j As Integer

    For i = 0 To 2
        For j = 0 To 4
            Sheets("w1").Cells(21 + i, 2 + j).Formula = ""
        Next j
    Next i

End Sub
