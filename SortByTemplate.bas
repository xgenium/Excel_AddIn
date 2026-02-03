Attribute VB_Name = "Module2"
Public Sub SortByTemplate(control As IRibbonControl)
    Dim templateRange As Range
    Dim nameRange As Range, targetRange As Range, helperColRange As Range
    Dim ws As Worksheet
    Dim rowStart As Long, rowEnd As Long
    Dim colName As Long, colDataEnd As Long, colHelper As Long
    Dim i As Long, nm As String
    Dim dict As Object, templateDict As Object, helperColDict As Object
    Dim targetNameCol As Long
    Dim val As Variant, correctName As String

    On Error Resume Next
    Set templateRange = Application.InputBox("Select template column with names", "Template Selection", Type:=8)
    On Error GoTo 0
    If templateRange Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set nameRange = Application.InputBox("Select the range of NAMES to sort by (one column only)", "Target Names Selection", Type:=8)
    On Error GoTo 0
    If nameRange Is Nothing Then Exit Sub
    If nameRange.Columns.Count <> 1 Then
        MsgBox "Please select exactly ONE column for names.", vbExclamation
        Exit Sub
    End If
    Set ws = nameRange.Worksheet
    rowStart = nameRange.Row
    rowEnd = nameRange.Rows(nameRange.Rows.Count).Row
    colName = nameRange.Column
    
    Application.ScreenUpdating = False
    ' shenanigans
    Dim targetWindow As Window, originalWindow As Window
    Dim targetSheet As Worksheet, templateSheet As Worksheet
    Dim targetWindowSheet As Worksheet, originalWindowSheet As Worksheet
    Set originalWindow = ActiveWindow
    Set templateSheet = templateRange.Worksheet
    Set targetSheet = nameRange.Worksheet
    If originalWindow.ActiveSheet Is targetSheet Then
        Set targetWindow = GetWindowOfSheet(templateSheet)
    Else
        Set targetWindow = GetWindowOfSheet(targetSheet)
    End If
    Set targetWindowSheet = targetWindow.ActiveSheet
    Set originalWindowSheet = originalWindow.ActiveSheet
    targetSheet.Parent.Save
    templateSheet.Parent.Save

    ' find boundaries (last col)
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        MsgBox "No data found on worksheet!", vbCritical
        Exit Sub
    End If
    colDataEnd = lastCell.Column

    ' helper column for sorting order
    colHelper = colDataEnd + 1
    Set helperColRange = ws.Range(ws.Cells(rowStart, colHelper), ws.Cells(rowEnd, colHelper))
    If Application.WorksheetFunction.CountA(helperColRange) > 0 Then
        MsgBox "Cannot use column " & colHelper & " as helper: Data exists there for selected rows. Please clear it and retry.", vbCritical
        Exit Sub
    End If

    Set targetRange = ws.Range(ws.Cells(rowStart, colName), ws.Cells(rowEnd, colDataEnd))
    
    Call Save2RangeBackup(templateRange, targetRange)

    ' sort template dict
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For i = 1 To templateRange.Rows.Count
        nm = NormalizeName(templateRange.Cells(i, 1).Value)
        If nm <> "" And Not dict.Exists(nm) Then dict.Add nm, dict.Count + 1
    Next i

    ws.Columns(colHelper).Insert Shift:=xlToRight

    targetNameCol = 1

    Set helperColDict = CreateObject("Scripting.Dictionary")
    For i = 1 To targetRange.Rows.Count
        val = ws.Cells(rowStart + i - 1, colHelper).Value
        If val <> "" Then helperColDict.Add i, val
    Next i

    ' fill helperCol with needed values
    For i = 1 To targetRange.Rows.Count
        nm = NormalizeName(targetRange.Cells(i, targetNameCol).Value)
        If dict.Exists(nm) Then
            ws.Cells(rowStart + i - 1, colHelper).Value = dict(nm)
            targetRange.Cells(i, targetNameCol).Font.ColorIndex = xlAutomatic
        Else
            correctName = ExistsInDictionary(nm, dict)
            If correctName <> "" Then
                ws.Cells(rowStart + i - 1, colHelper).Value = dict(correctName)
                targetRange.Cells(i, targetNameCol).Font.Color = RGB(0, 255, 0)
            Else
                ws.Cells(rowStart + i - 1, colHelper).Value = dict.Count + i
                targetRange.Cells(i, targetNameCol).Font.Color = RGB(255, 0, 0)
            End If
        End If
    Next i

    ws.Sort.SortFields.Clear
    targetRange.Resize(, targetRange.Columns.Count + 1).Sort _
        Key1:=ws.Range(ws.Cells(rowStart, colHelper), ws.Cells(rowEnd, colHelper)), _
        Order1:=xlAscending, Header:=xlNo

    ws.Columns(colHelper).Delete

    For i = 1 To targetRange.Rows.Count
        If helperColDict.Exists(i) Then
            ws.Cells(rowStart + i - 1, colHelper).Value = helperColDict(i)
        End If
    Next i

    ' highlight names not found in target
    Set templateDict = CreateObject("Scripting.Dictionary")
    For i = 1 To targetRange.Rows.Count
        nm = NormalizeName(targetRange.Cells(i, targetNameCol).Value)
        If nm <> "" And Not templateDict.Exists(nm) Then
            templateDict.Add nm, dict.Count + 1
        End If
    Next i
    For i = 1 To templateRange.Rows.Count
        nm = NormalizeName(templateRange.Cells(i, 1).Value)
        If Not templateDict.Exists(nm) Then
            correctName = ExistsInDictionary(nm, templateDict)
            If correctName = "" Then
                templateRange.Cells(i, 1).Font.Color = RGB(0, 0, 255)
            End If
        End If
    Next i
    
    targetWindow.Activate
    targetWindowSheet.Activate
    originalWindow.Activate
    originalWindowSheet.Activate
    
    PrintExplanation targetSheet

    Application.ScreenUpdating = True
End Sub

Function NormalizeName(nm As String) As String
    NormalizeName = UCase(Trim(Replace(nm, " ", "")))
End Function

Function AlmostMatch(str1 As String, str2 As String) As Boolean
    Dim i As Long, j As Long
    Dim len1 As Long, len2 As Long
    Dim diffCount As Long, maxDiffCount As Long

    len1 = Len(str1)
    len2 = Len(str2)

    maxDiffCount = len1 * 0.2 ' 80% match
    If maxDiffCount < 2 Then maxDiffCount = 2

    If Abs(len1 - len2) > 1 Then
        AlmostMatch = False
        Exit Function
    End If

    i = 1
    j = 1
    diffCount = 0

    Do While i <= len1 And j <= len2
        If Mid(str1, i, 1) = Mid(str2, j, 1) Then
            i = i + 1
            j = j + 1
        Else
            diffCount = diffCount + 1
            If diffCount > maxDiffCount Then
                AlmostMatch = False
                Exit Function
            End If
            If len1 > len2 Then
                i = i + 1
            ElseIf len2 > len1 Then
                j = j + 1
            Else
                i = i + 1
                j = j + 1
            End If
        End If
    Loop

    If i <= len1 Or j <= len2 Then
        diffCount = diffCount + 1
    End If

    AlmostMatch = (diffCount <= maxDiffCount)
End Function

Function ExistsInDictionary(nm As String, dict As Object) As String
    Dim key As Variant
    ExistsInDictionary = ""
    For Each key In dict.keys
        If AlmostMatch(CStr(key), nm) Then
            ExistsInDictionary = CStr(key)
            Exit Function
        End If
    Next key
End Function

Private Sub PrintExplanation(ws As Worksheet)
    On Error Resume Next
    ws.Shapes("ColorExplanation").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 300, 50, 270, 120)
    shp.Name = "ColorExplanation"
    With shp
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Text = "Green: Name contains mistakes" & vbCrLf & _
                                     "Red: Name not in template range" & vbCrLf & _
                                     "Blue: Missing in target range"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .TextFrame2.MarginLeft = 6
        .TextFrame2.MarginTop = 6
    End With
    With shp.TextFrame2.TextRange
        .Paragraphs(1).Font.Fill.ForeColor.RGB = RGB(0, 255, 0)
        .Paragraphs(2).Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Paragraphs(3).Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
    End With
    ' Calculate visible center
    Dim visRange As Range
    Set visRange = ws.Parent.Windows(1).VisibleRange
    Dim left As Double, top As Double, width As Double, height As Double
    left = visRange.Cells(1, 1).left
    top = visRange.Cells(1, 1).top
    width = visRange.width
    height = visRange.height

    ' Center the shape within the visible window
    shp.left = left + (width - shp.width) / 2
    shp.top = top + (height - shp.height) / 2
End Sub

Public Function GetWindowOfSheet(ws As Worksheet) As Window
    Dim wnd As Window
    For Each wnd In ws.Parent.Windows   ' ws.Parent is the workbook
        If wnd.ActiveSheet Is ws Then
            Set GetWindowOfSheet = wnd
            Exit Function
        End If
    Next wnd
    ' If no window currently has this sheet active, return the first window
    Set GetWindowOfSheet = ws.Parent.Windows(1)
End Function
