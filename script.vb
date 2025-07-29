Sub ExtractCampaignName()
    Dim selectedCell As Range
    Dim cellText As String
    Dim pipePositions(1 To 10) As Integer
    Dim pipeCount As Integer
    Dim i As Integer
    Dim pos As Integer
    Dim extractedText As String
    ' Check if a cell is selected
    If Selection.Cells.Count <> 1 Then
        MsgBox "Please select exactly one cell."
        Exit Sub
    End If
    Set selectedCell = Selection
    cellText = selectedCell.Value
    ' Find all pipe positions
    pos = 1
    pipeCount = 0
    Do While pos <= Len(cellText) And pipeCount < 10
        pos = InStr(pos, cellText, "|")
        If pos = 0 Then Exit Do
        pipeCount = pipeCount + 1
        pipePositions(pipeCount) = pos
        pos = pos + 1
    Loop
    ' Check if we have at least 6 pipes
    If pipeCount < 6 Then
        MsgBox "Need at least 6 pipe characters. Found only " & pipeCount & "."
        Exit Sub
    End If
    ' Extract text between 5th and 6th pipes
    extractedText = Mid(cellText, pipePositions(5) + 1, pipePositions(6) - pipePositions(5) - 1)
    ' Replace the cell content with the extracted text
    selectedCell.Value = extractedText
End Sub
