
Private Sub CommandButton1_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open �������, ���� ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   
   ' Call mDeleteAllActiveXButtons(WB_NAME)
   
   
   Call DeleteAllActiveXControls(WB_NAME)
   Call SaveJustXLSX(WB_NAME)

End Sub


Private Sub CommandButton2_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open �������, ���� ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   Call DeleteHiddenSheets(WB_NAME)

End Sub


' Erase All ActiveX Control In Send
Private Sub CommandButton3_Click()


    Call OpenXLSMFilesInSend
  

End Sub


Sub OpenXLSMFilesInSend()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    
    folderPath = "d:\05_Send\"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Downloads folder not found!"
        Exit Sub
    End If
    
    fileName = Dir(folderPath & "*.xlsm")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        Call DeleteAllActiveXControls(fileName)
        Call SaveJustXLSX(fileName)
        
        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
    
    MsgBox "All .xlsm files in Downloads folder have been opened."
End Sub



Sub OpenXLSMFilesInDownloads()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    
    folderPath = Environ("USERPROFILE") & "\Downloads\"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Downloads folder not found!"
        Exit Sub
    End If
    fileName = Dir(folderPath & "*.xlsm")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)

        fileName = Dir
    Loop
    
    MsgBox "All .xlsm files in Downloads folder have been opened."
End Sub


Sub OpenXlsmFiles()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sFolder As String: sFolder = "C:\Users\YourUsername\Downloads"
    
    If Dir(sFolder) <> "" Then
        
        For Each fFile In fso.GetFolder(sFolder).Files
            If Right(fFile.Name, 4) = ".xlsm" Then
                
                Workbooks.Open (fFile.Path)
                Application.ScreenUpdating = True
            
            End If
        
        Next fFile
    
    Else
        MsgBox "The Downloads folder was not found. Please check the path.", vbCritical, "Error"
    
    End If

End Sub

Option Explicit



Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim fName As String
    
    Dim pathToExport As String
    pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        
        fName = item.CodeModule.Name
        Debug.Print lineToPrint
        SaveTextToFile lineToPrint, pathToExport & fName & ".bas"
        
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub


Function GetNumeric2(ByVal CellRef As String) As String
    Dim StringLength, i  As Integer
    Dim result      As String
    
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i
    GetNumeric2 = result
End Function


Function GetNumberOfWell(ByVal fName As String) As Integer
    Dim save_name As String
    Dim n As Integer
    
    Workbooks(fName).Worksheets("Well").Activate
    
    With Workbooks(fName).Sheets("Well")
        n = .Cells(.Rows.Count, "A").End(xlUp).Row
        n = CInt(GetNumeric2(.Cells(n, "A").Value))
    End With
    
    GetNumberOfWell = n
End Function



Function CheckWorksheetExistence(ByVal fName As String) As Boolean
    Dim ws As Worksheet
    Dim wsName As String
    
    wsName = "Aggregate1"
    
    For Each ws In Workbooks(fName).Worksheets
        If ws.Name = wsName Then
            ' MsgBox "Worksheet '" & wsName & "' exists!"
            CheckWorksheetExistence = True
            Exit Function
        End If
    Next ws
    CheckWorksheetExistence = False
End Function




Sub DeleteHiddenSheets(ByVal fName As String)

    Dim ws As Worksheet

    If MsgBox("Are you sure you want to delete all hidden sheets in this workbook?", vbYesNo, "Delete Hidden Sheets?") <> vbYes Then Exit Sub

    Workbooks(fName).Activate
    
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ThisWorkbook.Activate

End Sub


Public Function MyDownPath() As String
    MyDownPath = Environ$("USERPROFILE") & "\" & "Downloads"
    Debug.Print MyDownPath
End Function



Sub SaveJustXLSX(ByVal fName As String)
    Dim mypath, fname0 As String
    Dim fso As New Scripting.FileSystemObject
        
    mypath = MyDownPath
    Debug.Print "path" + mypath
    
    
    Workbooks(fName).Activate
    fname0 = fso.GetBaseName(fName)
    
    ' Disable alerts to overwrite without prompt
    Application.DisplayAlerts = False
    
    
    ActiveWorkbook.SaveAs fileName:= _
        mypath & "\" & fname0 & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        
        
    On Error Resume Next
    'Workbooks(fname0).Close SaveChanges:=False
    Workbooks("'" & fname0 & "'").Close SaveChanges:=False
    On Error GoTo 0
    
    ' Enable alerts back
    Application.DisplayAlerts = True
    
    
    ThisWorkbook.Activate
End Sub



Sub DeleteAllActiveXControls(ByVal fName As String)
    Dim myControl As Object
    Dim nofwell, i As Integer
    
    Workbooks(fName).Activate
    
    If CheckWorksheetExistence(fName) Then ' ���� �⺻��������Ÿ �̶�� ...
    
        nofwell = Service.GetNumberOfWell(fName)
        
        Workbooks(fName).Worksheets("Well").Activate
        For i = 1 To nofwell
            With Workbooks(fName).Worksheets("Well")
                .Range("O" & (i + 3)).Select
                Selection.Copy
                .Range("O" & (i + 3)).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End With
        Next i
        
        For i = 1 To nofwell
            Workbooks(fName).Worksheets(CStr(i)).Activate
            
            With Workbooks(fName).Worksheets(CStr(i))
                .Range("F21").Select
                Selection.Copy
                .Range("F21").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End With
        Next i
        
        For Each ws In Workbooks(fName).Worksheets
            Select Case ws.Name
                Case "AggSum", "YangSoo", "water", "AggStep", "AggChart", "Aggregate2", "Aggregate1", "aggWhpa"
                    ' MsgBox "Worksheet '" & ws.Name & "' is in the list."
                    Application.DisplayAlerts = False
                    ws.Delete
                    Application.DisplayAlerts = True
                Case Else
                    Debug.Print ws.Name
            End Select
        Next ws
    End If
    
    For Each ws In Workbooks(fName).Worksheets
        For Each myControl In ws.OLEObjects
            myControl.Delete
        Next myControl
     Next ws
    
    ThisWorkbook.Activate
End Sub



Sub mDeleteAllActiveXButtons(ByVal fName As String)
    Dim ws As Worksheet
    Dim obj As OLEObject
    
    Workbooks(fName).Activate
          
    For Each ws In Workbooks(fName).Worksheets
        For Each obj In ws.OLEObjects
            If TypeName(obj.Object) = "CommandButton" Then
                obj.Delete
            End If
        Next obj
    Next ws
    
    ThisWorkbook.Activate
    
End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.Name & vbCrLf
    Next
    
    Cells(1, 1).Value = workbookNames
End Sub

'Function GetOtherFileName() As String
'    Dim Workbook As Workbook
'    Dim workbookNames As String
'    Dim bool As Boolean
'    Dim i As Long
'
'    workbookNames = ""
'    bool = False
'
'    For Each Workbook In Application.Workbooks
'        If StrComp(ThisWorkbook.Name, Workbook.Name, vbTextCompare) = 0 Then
'            bool = True
'            GoTo NEXT_ITERATION
'        End If
'
'        If bool Then
'            Exit For
'        End If
'
'NEXT_ITERATION:
'    Next
'
'    GetOtherFileName = Workbook.Name
'End Function


Function GetOtherFileName() As String
' refactor by instr function

    Dim OtherWorkbook As Workbook
    Dim ThisWorkbookName As String
    Dim OtherWorkbookName As String

    ThisWorkbookName = ThisWorkbook.Name
    
    For Each OtherWorkbook In Application.Workbooks
        If InStr(1, OtherWorkbook.Name, ThisWorkbookName, vbTextCompare) = 0 Then
            OtherWorkbookName = OtherWorkbook.Name
            Exit For
        End If
    Next OtherWorkbook
    
    GetOtherFileName = OtherWorkbookName
End Function


'
' refactor function
'
'Function GetOtherFileName() As String
'    Dim OtherWorkbook As Workbook
'    Dim ThisWorkbookName As String
'    Dim OtherWorkbookName As String
'
'    ThisWorkbookName = ThisWorkbook.Name
'
'    For Each OtherWorkbook In Application.Workbooks
'        If StrComp(ThisWorkbookName, OtherWorkbook.Name, vbTextCompare) <> 0 Then
'            OtherWorkbookName = OtherWorkbook.Name
'            Exit For
'        End If
'    Next OtherWorkbook
'
'    GetOtherFileName = OtherWorkbookName
'End Function
'




'Function CheckSubstring(str As String, chk As String) As Boolean
'
'    If InStr(str, chk) > 0 Then
'        ' The string contains "chk"
'        CheckSubstring = True
'    Else
'        ' The string does not contain "chk"
'        CheckSubstring = False
'    End If
'End Function



Function CheckSubstring(str As String, chk As String) As Boolean
    CheckSubstring = (InStr(str, chk) > 0)
End Function


