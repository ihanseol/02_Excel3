Attribute VB_Name = "Module1"
Sub 매크로1()
Attribute 매크로1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로1 매크로
'

'
    ActiveWindow.SmallScroll Down:=-22
    Range("E2").Select
    Selection.End(xlDown).Select
End Sub
