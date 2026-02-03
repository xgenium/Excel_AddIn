Attribute VB_Name = "Module1"
Public Sub InsertInColoredCells(control As IRibbonControl)
    Dim targetRange As Range
    
    On Error Resume Next
    Set targetRange = Application.InputBox( _
    Prompt:="Select target range with colored cells", _
    Title:="Target Range Selection", Type:=8)
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each c In targetRange
        If c.DisplayFormat.Interior.ColorIndex <> -4142 And c.Value = "" Then ' No color
            c.Value = 1
        End If
    Next c
    
    Application.ScreenUpdating = True
End Sub


