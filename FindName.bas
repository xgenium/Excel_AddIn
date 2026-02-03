Attribute VB_Name = "Module3"
Public Sub SwitchWindowAndFind(control As IRibbonControl)
Attribute SwitchWindowAndFind.VB_ProcData.VB_Invoke_Func = "j\n14"

    Application.ScreenUpdating = False
    Dim OriginalCell As Range
    Dim TargetCell As Range
    Set OriginalCell = ActiveCell
    If OriginalCell.Value = "" Then
        MsgBox "Empty cell selected"
        Exit Sub
    End If
    
    If Windows.Count > 1 Then
        Windows(ActiveWindow.Index Mod Windows.Count + 1).Activate
    Else
        MsgBox "No other Excel window to switch to."
        Exit Sub
    End If
    
    Set TargetCell = FindName(ActiveSheet, OriginalCell.Value)
    If Not TargetCell Is Nothing Then
        Range(TargetCell.Address).Activate
    Else
        MsgBox "Not Found"
    End If
    
    Application.ScreenUpdating = True
End Sub

Function FindName(ws As Worksheet, Name As String) As Range
    Dim rng As Range
    Dim cell As Range
    Dim trimmedCellValue As String, trimmedName As String
    Dim isAlmostMatch As Boolean
    With ws
        Set rng = .UsedRange.Find(Name)
        If Not rng Is Nothing Then
            Set FindName = rng
            Exit Function
        Else
            For Each cell In .UsedRange
                trimmedCellValue = Replace(cell.Value, " ", "")
                trimmedName = Replace(Name, " ", "")
                If trimmedCellValue = trimmedName Or AlmostMatch(trimmedName, trimmedCellValue) Then
                    cell.Font.Color = RGB(0, 0, 255)
                    Set FindName = cell
                    Exit Function
                End If
            Next cell
        End If
    End With
End Function



Function AlmostMatch(str1 As String, str2 As String) As Boolean
    Dim i As Long, j As Long
    Dim len1 As Long, len2 As Long
    Dim diffCount As Long, maxDiffCount As Long
    
    len1 = Len(str1)
    len2 = Len(str2)
    
    maxDiffCount = len1 * 0.2 ' 80% match
    If maxDiffCount < 1 Then maxDiffCount = 2
    
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
