Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim col As String
    col = "G" ' Specify your column here

    ' Check if the selected cell is in column A
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        ' Copy the contents of the selected cell
        Target.Copy
    End If

    ' Check if the selected cell is in the specified column (G)
    If Not Intersect(Target, Range(col & ":" & col)) Is Nothing Then
        ' If the cell is empty or contains "Done", toggle its content
        If Target.Value = "" Then
            Target.Value = "Done"
        ElseIf Target.Value = "Done" Then
            Target.Value = ""
        End If
    End If
End Sub



Sub RemoveMatchingPart()
    Dim rng As Range
    Dim cell As Range
    Dim sourceCol As String
    Dim targetCol As String
    Dim ws As Worksheet

    ' Auto-detect the active sheet
    Set ws = ActiveSheet

    ' Request the source and target columns
    sourceCol = InputBox("Enter the source column (e.g., 'A')", "Input needed", "A")
    targetCol = InputBox("Enter the target column (e.g., 'B')", "Input needed", "B")

    ' Define the range to check
    Set rng = ws.Range(targetCol & "2:" & targetCol & ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row)

    ' Check each cell in the range
    For Each cell In rng
        ' If source cell's value is found in target cell
        If InStr(cell.Value, ws.Cells(cell.Row, sourceCol).Value) > 0 Then
            ' Replace the first occurrence of source cell's value in target cell with an empty string
            cell.Value = Replace(cell.Value, ws.Cells(cell.Row, sourceCol).Value, "", 1, 1)
        End If
    Next cell

    ' Auto align left on the target column
    ws.Range(targetCol & ":" & targetCol).EntireColumn.HorizontalAlignment = xlLeft
End Sub


Sub FormatAndRenameColumns()
    Dim ws As Worksheet
    Dim col As Range
    Dim lastCol As Long

    ' Auto-detect the active sheet
    Set ws = ActiveSheet

    ' Delete columns named "body subtype" and "is terraformable"
    On Error Resume Next
    ws.Columns(ws.Rows(1).Find("body subtype").Column).Delete
    ws.Columns(ws.Rows(1).Find("is terraformable").Column).Delete
    On Error GoTo 0

    ' Get the last column number
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Format and rename certain columns
    For i = 1 To lastCol
        Set col = ws.Cells(1, i)
        If col.Value = "Distance To Arrival" Then
            ws.Range(col.Offset(1, 0), ws.Cells(ws.Rows.Count, i).End(xlUp)).NumberFormat = "#,##0.00"
            col.Value = "distance (LS)"
        ElseIf col.Value = "Estimated Scan Value" Or col.Value = "Estimated Mapping Value" Then
            ws.Range(col.Offset(1, 0), ws.Cells(ws.Rows.Count, i).End(xlUp)).NumberFormat = "#,##0"
            If col.Value = "Estimated Scan Value" Then
                col.Value = "Scan Value"
            Else
                col.Value = "Map Value"
            End If
        End If
    Next i
End Sub

Sub ApplyDarkMode()
    Dim ws As Worksheet

    ' Auto-detect the active sheet
    Set ws = ActiveSheet

    ' Select all cells
    ws.Cells.Select

    ' Change the background color to gray
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -1
        .PatternTintAndShade = 0
    End With

    ' Change the font color to white
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 1
    End With

    ' Apply white borders
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = RGB(255, 255, 255)
    End With

    ' Deselect cells
    ws.Cells(1, 1).Select
End Sub

