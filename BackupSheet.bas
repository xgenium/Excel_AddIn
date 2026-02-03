Attribute VB_Name = "Module4"
' Ensures a hidden backup sheet exists in the workbook containing the backup range.
Private Sub EnsureBackupSheetForWorkbook(ByRef wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("__Backup")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "__Backup"
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

' Back up a single range to the __Backup sheet of its parent workbook.
Public Sub SaveRangeBackup(rng As Range)
    Application.ScreenUpdating = False

    Dim wsBackup As Worksheet, wsOriginal As Worksheet
    Set wsOriginal = ActiveSheet
    Dim wb As Workbook
    Set wb = rng.Parent.Parent
    EnsureBackupSheetForWorkbook wb
    Set wsBackup = wb.Sheets("__Backup")
    wsBackup.Cells.Clear

    ' Metadata
    wsBackup.Range("A1").Value = wb.Name
    wsBackup.Range("A2").Value = rng.Worksheet.Name
    wsBackup.Range("A3").Value = rng.Address(False, False)
    wsBackup.Range("A4").Value = rng.Rows.Count

    ' Data (format + values)
    rng.Copy
    wsBackup.Range("A5").PasteSpecial xlPasteAll
    wsOriginal.Activate
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

' Back up two ranges, each to their own proper __Backup sheet in their workbook.
' Returns nothing, only side-effect in sheets.
Public Sub Save2RangeBackup(rng1 As Range, rng2 As Range)
    SaveRangeBackup rng1
    If rng1.Parent.Parent Is rng2.Parent.Parent Then
        ' For same workbook, append second region below first
        Dim wsBackup As Worksheet
        Set wsBackup = rng1.Parent.Parent.Sheets("__Backup")
        Dim rStart As Long
        rStart = wsBackup.Cells(wsBackup.Rows.Count, 1).End(xlUp).Row + 2
        wsBackup.Range("A" & rStart).Value = rng2.Parent.Parent.Name
        wsBackup.Range("A" & (rStart + 1)).Value = rng2.Worksheet.Name
        wsBackup.Range("A" & (rStart + 2)).Value = rng2.Address(False, False)
        rng2.Copy
        wsBackup.Range("A" & (rStart + 3)).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    Else
        SaveRangeBackup rng2
    End If
End Sub

' Undo backup for a given workbook (reverse the data for that workbook).
Private Sub UndoBackupForWorkbook(wb As Workbook)
    Dim originalSheet As Worksheet
    Set originalSheet = wb.ActiveSheet
    Dim wsBackup As Worksheet
    On Error Resume Next
    Set wsBackup = wb.Sheets("__Backup")
    On Error GoTo 0
    If wsBackup.Range("A1").Value = "" Then
        MsgBox "No backup metadata found in " & wb.Name
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim rngRestore As Range
    Dim rowsCount As Long, secondBlockRow As Long
    Set ws = wb.Worksheets(CStr(wsBackup.Range("A2").Value))
    Set rngRestore = ws.Range(wsBackup.Range("A3").Value)
    rowsCount = wsBackup.Range("A4").Value
    ws.Activate
    If ws.ProtectContents Then
        MsgBox "Worksheet " & ws.Name & " is protected. Cannot paste.", vbCritical
        Exit Sub
    End If
    wsBackup.Range("A5").Resize(rngRestore.Rows.Count, rngRestore.Columns.Count).Copy
    ' MsgBox "About to paste to workbook: " & wb.Name & ", sheet: " & ws.Name & ", range: " & rngRestore.Address
    rngRestore.PasteSpecial xlPasteAll
    
    secondBlockRow = 5 + rowsCount + 1
    If wsBackup.Range("A" & secondBlockRow) <> "" Then
        Set ws = wb.Worksheets(CStr(wsBackup.Range("A" & (secondBlockRow + 1)).Value))
        Set rngRestore = ws.Range(wsBackup.Range("A" & (secondBlockRow + 2)).Value)
        wsBackup.Range("A" & (secondBlockRow + 3)).Resize(rngRestore.Rows.Count, rngRestore.Columns.Count).Copy
        rngRestore.PasteSpecial xlPasteAll
    End If
    
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    wsBackup.Visible = xlSheetVisible
    wsBackup.Delete
    originalSheet.Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


End Sub


Public Sub UndoBackup(control As IRibbonControl)
    Dim wb As Workbook
    Dim anyFound As Boolean
    anyFound = False
    For Each wb In Application.Workbooks
        On Error Resume Next
        If wb.Sheets("__Backup") Is Nothing Then
            'Dim msg As String
            'msg = "No backup found in " & wb.Name
            'MsgBox msg, vbInformation
        Else
            ' MsgBox "Found backup for " & wb.Name
            UndoBackupForWorkbook wb
            anyFound = True
        End If
        On Error GoTo 0
    Next wb
    If Not anyFound Then
        MsgBox "No backups found in any open workbook.", vbInformation
        Exit Sub
    End If
    MsgBox "Undo complete"
End Sub
