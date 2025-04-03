Sub TransPose30Year()
    
    Dim i, j        As Integer
    Dim i1, i2      As Integer
    Dim sYear, eYear As Integer
    
    Range("C1").Select
    Selection.End(xlDown).Select
    
    eYear = Year(Now()) - 1
    sYear = eYear - 29
    
    For i = 1 To 30
        
        i1 = 12 * (i - 1) + 9
        i2 = i1 + 11
        
        Range("C" & CStr(i1) & ":C" & CStr(i2)).Select
        Selection.Copy
        
        j = i + 8
        Range("G" & CStr(j)).Select
        
        Range("F" & CStr(j)).Value = sYear + i - 1
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                               False, Transpose:=True
        
    Next i
    
End Sub


Public Function MyDocPath() As String
    MyDocPath = Environ$("USERPROFILE") & "\" & "Downloads" & "\"
    Debug.Print MyDocsPath
End Function

Public Function SelectDocument() As String
    On Error GoTo Trap
    
    Dim fd          As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .InitialFileName = MyDocPath
        
        .Title = "Please Select the file."
        .Filters.Clear
        .Filters.Add "Excel 2003", "*.xls?"
    End With
    
    'if a selection was made, return the file path
    If fd.Show = -1 Then SelectDocument = fd.SelectedItems(1)
    
Leave:
    Set fd = Nothing
    On Error GoTo 0
    Exit Function
    
Trap:
    MsgBox Err.Description, vbCritical
    Resume Leave
End Function

Function GetFileSize(strFile As Variant) As Long
    Dim lngFSize    As Long, lngDSize As Long
    Dim oFO         As Object
    Dim OFS         As Object
    
    lngFSize = 0
    Set OFS = CreateObject("Scripting.FileSystemObject")
    
    If OFS.FileExists(strFile) Then
        Set oFO = OFS.Getfile(strFile)
        GetFileSize = oFO.Size
    End If
    
End Function

Sub import30YearData()
    
    Dim directory   As String, fileName As String
    Dim fd          As Office.FileDialog
    
    Dim file_name   As String
    Dim thisFileName As String
    file_name = SelectDocument()
    
    Debug.Print getFileName(file_name)
    Debug.Print getPath(file_name)
    
    thisFileName = ActiveWorkbook.name
    
    If file_name <> vbNullString Then
        Workbooks.Open fileName:=file_name
    End If
    
    If file_name = "" Then Exit Sub
    If GetFileSize(file_name) < 5000 Then
        file_name = getFileName(file_name)
        Workbooks(file_name).Close SaveChanges:=False
        Exit Sub
    End If
    
    file_name = getFileName(file_name)
    Call TransPose30Year
    
    Range("F9:R38").Select
    Selection.Copy
    
    Windows(thisFileName).Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Workbooks(file_name).Close SaveChanges:=False
    Application.CutCopyMode = False
    
End Sub


Function OpenRecentFile() As String
    
    Dim strFilePath As String
    Dim fileName As String
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim aStack As Object
    
    Dim lngFileCount As Long
    Dim lngCounter As Long

    strFilePath = Environ("USERPROFILE") & "\Downloads\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFilePath)
    Set aStack = CreateObject("System.Collections.Stack")
    
    Debug.Print strFilePath
     
    Set objFolder = objFSO.GetFolder(strFilePath)
    lngFileCount = objFolder.Files.Count
               
    For Each objFile In objFolder.Files
        aStack.Push (objFile.name)
        Debug.Print objFile.name
    Next objFile
        
        
    For lngCounter = 1 To lngFileCount
    
        fileName = aStack.pop
               
        If Right(fileName, 4) = ".xls" Or Right(fileName, 4) = ".csv" Then
            Workbooks.Open fileName:=strFilePath & fileName
            OpenRecentFile = fileName
            Exit For
        End If

    Next lngCounter

    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
    Set aStack = Nothing
    
End Function


Sub import30RecentData()

    Dim thisFileName, file_name As String
    
    thisFileName = ActiveWorkbook.name
    
    file_name = OpenRecentFile
    Call TransPose30Year
    
    Range("F9:R38").Select
    Selection.Copy
    
    Windows(thisFileName).Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    Workbooks(file_name).Close SaveChanges:=False
    
    Application.CutCopyMode = False
   
    
End Sub



Function getPath(FullPath As String, Optional Delim As String = "\") As String
    Dim a
    a = Split(FullPath & "$", Delim)
    getPath = Join(Filter(a, a(UBound(a)), False), Delim)
End Function

Function getFileName(FullPath As String) As String
    getFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
End Function

Function GetDirOrFileSize(strFolder As String, Optional strFile As Variant) As Long
    
    Dim lngFSize    As Long, lngDSize As Long
    Dim oFO         As Object
    Dim oFD         As Object
    Dim OFS         As Object
    
    lngFSize = 0
    Set OFS = CreateObject("Scripting.FileSystemObject")
    
    If strFolder = "" Then strFolder = ActiveWorkbook.Path
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    'Thanks to Jean-Francois Corbett, you can use also OFS.BuildPath(strFolder, strFile)
    
    If OFS.FolderExists(strFolder) Then
        If Not IsMissing(strFile) Then
            
            If OFS.FileExists(strFolder & strFile) Then
                Set oFO = OFS.Getfile(strFolder & strFile)
                GetDirOrFileSize = oFO.Size
            End If
            
        Else
            Set oFD = OFS.GetFolder(strFolder)
            GetDirOrFileSize = oFD.Size
        End If
        
    End If
    
End Function

Sub CallOpenAPI()
    Dim strURL      As String
    Dim strResult   As String
    
    Dim objHttp     As New WinHttpRequest
    
    strURL = "Open API �ּҸ� �Է��ϼ���"
    objHttp.Open "GET", strURL, False
    objHttp.send
    
    If objHttp.Status = 200 Then        '�������� ���
    strResult = objHttp.responseText
    
    'XML�� ����
    Dim objXml      As MSXML2.DOMDocument60
    Set objXml = New DOMDocument60
    objXml.LoadXML (strResult)
    
    '��� ����
    Dim nodeList    As IXMLDOMNodeList
    Dim nodeRow     As IXMLDOMNode
    Dim nodeCell    As IXMLDOMNode
    Dim nRowCount   As Integer
    Dim nCellCount  As Integer
    
    Set nodeList = objXml.SelectNodes("/response/fields/field")
    
    nRowCount = Range("A60000").End(xlUp).Row
    For Each nodeRow In nodeList
        nRowCount = nRowCount + 1
        
        nCellCount = 0
        For Each nodeCell In nodeRow.ChildNodes
            nCellCount = nCellCount + 1
            '������ �� �ݿ�
            Cells(nRowCount, nCellCount).Value = nodeCell.Text
        Next nodeCell
        
    Next nodeRow
    
Else
    MsgBox "���ӿ� ������ �߻��߽��ϴ�"
    
End If
End Sub

