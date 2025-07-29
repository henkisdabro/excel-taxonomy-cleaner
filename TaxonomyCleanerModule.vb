'================================================================================
' EXCEL TAXONOMY CLEANER - Main Module
'================================================================================
' 
' INSTALLATION INSTRUCTIONS:
' 1. Open Excel VBA Editor (Alt + F11)
' 2. Right-click your project → Insert → Module
' 3. Copy and paste this entire code into the new module
' 4. Optionally create the UserForm (see TaxonomyCleanerForm.vb for instructions)
'
' USAGE:
' 1. Select cells with pipe-delimited text (e.g., "Marketing|Campaign|Q4|Social|Facebook")
' 2. Run the TaxonomyCleaner macro (assign to button or use Alt+F8)
' 3. Choose segment number to extract that specific part
'
' EXAMPLES:
' For text "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725":
' - Segment 1: "FY24_26" (1st segment)
' - Segment 3: "Tourism WA" (3rd segment) 
' - Segment 5: "Always On Remarketing" (5th segment)
' - Segment 9: "Conversions" (9th segment)
' - Activation ID: "DJTDOM060725" (after colon)
'================================================================================

' Global variables for UNDO functionality
Type UndoData
    CellAddress As String
    OriginalValue As String
End Type

Dim UndoArray() As UndoData
Dim UndoCount As Integer
Dim LastSegmentNumber As Integer

' Main macro to be called when button is pressed
Sub TaxonomyCleaner()
    ' Check if cells are selected
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select cells containing text before running this tool.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Check if selection contains text
    Dim hasText As Boolean
    hasText = False
    Dim cell As Range
    For Each cell In Selection
        If Len(Trim(cell.Value)) > 0 Then
            hasText = True
            Exit For
        End If
    Next cell
    
    If Not hasText Then
        MsgBox "Please select cells that contain text.", vbExclamation, "No Text Found"
        Exit Sub
    End If
    
    ' Try to show UserForm first, fallback to InputBox if form doesn't exist
    On Error GoTo UseInputBox
    TaxonomyCleanerForm.Show
    Exit Sub
    
UseInputBox:
    ' Fallback to simple input dialog if UserForm not created
    Call ShowSegmentSelector
End Sub

' Simple input dialog interface (fallback when UserForm doesn't exist)
Sub ShowSegmentSelector()
    Dim selectedSegment As String
    Dim validNumber As Integer
    
    ' Show clean, simple interface
    selectedSegment = InputBox("TAXONOMY CLEANER - Segment Extractor" & vbCrLf & vbCrLf & _
                              "This tool extracts specific segments from pipe-delimited data." & vbCrLf & vbCrLf & _
                              "EXAMPLE: 'FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725'" & vbCrLf & _
                              "  Segment 1 = FY24_26" & vbCrLf & _
                              "  Segment 3 = Tourism WA" & vbCrLf & _
                              "  Segment 5 = Always On Remarketing" & vbCrLf & _
                              "  Segment 9 = Conversions" & vbCrLf & _
                              "  A = DJTDOM060725 (Activation ID)" & vbCrLf & vbCrLf & _
                              "Enter segment number (1-9) or 'A' for Activation ID:", "Taxonomy Cleaner", "")
    
    ' Validate and execute
    If selectedSegment = "" Then Exit Sub ' User cancelled
    
    ' Check for Activation ID
    If UCase(Trim(selectedSegment)) = "A" Then
        Call ExtractActivationID
    ElseIf IsNumeric(selectedSegment) Then
        validNumber = CInt(selectedSegment)
        If validNumber >= 1 And validNumber <= 9 Then
            Call ExtractPipeSegment(validNumber)
        Else
            MsgBox "Please enter a number between 1 and 9, or 'A' for Activation ID.", vbExclamation, "Invalid Input"
        End If
    Else
        MsgBox "Please enter a valid number between 1 and 9, or 'A' for Activation ID.", vbExclamation, "Invalid Input"
    End If
End Sub

' Core function to extract specific segment from pipe-delimited text
Sub ExtractPipeSegment(segmentNumber As Integer)
    Dim cell As Range
    Dim cellText As String
    Dim extractedText As String
    Dim pipePositions(1 To 10) As Integer
    Dim pipeCount As Integer
    Dim pos As Integer
    Dim processedCount As Integer
    Dim i As Integer
    
    ' Initialize undo functionality
    UndoCount = 0
    LastSegmentNumber = segmentNumber
    ReDim UndoArray(1 To Selection.Cells.Count)
    
    ' Disable screen updating for better performance, then re-enable for visual update
    Application.ScreenUpdating = False
    
    processedCount = 0
    
    For Each cell In Selection
        On Error GoTo NextCell ' Skip any problematic cells
        cellText = CStr(cell.Value)
        
        ' Skip empty cells
        If Len(Trim(cellText)) = 0 Then
            GoTo NextCell
        End If
        
        ' Find all pipe positions
        pipeCount = 0
        pos = 1
        Do While pos <= Len(cellText) And pipeCount < 10
            pos = InStr(pos, cellText, "|")
            If pos = 0 Then Exit Do
            pipeCount = pipeCount + 1
            pipePositions(pipeCount) = pos
            pos = pos + 1
        Loop
        
        ' Extract the requested segment
        If segmentNumber = 1 Then
            ' First segment: from start to first pipe (or entire text if no pipes)
            If pipeCount >= 1 Then
                extractedText = Trim(Left(cellText, pipePositions(1) - 1))
            Else
                extractedText = Trim(cellText) ' No pipes, use entire text
            End If
            ' Store original value for undo before changing
            UndoCount = UndoCount + 1
            UndoArray(UndoCount).CellAddress = cell.Address
            UndoArray(UndoCount).OriginalValue = cellText
            
            cell.Value = extractedText
            processedCount = processedCount + 1
        ElseIf segmentNumber <= pipeCount + 1 Then
            ' Middle/end segments: between pipes
            Dim startPos As Integer
            Dim endPos As Integer
            
            If segmentNumber <= pipeCount Then
                ' Between two pipes
                startPos = pipePositions(segmentNumber - 1) + 1
                endPos = pipePositions(segmentNumber) - 1
            Else
                ' Last segment after final pipe (but stop at colon if present)
                startPos = pipePositions(pipeCount) + 1
                Dim colonPos As Integer
                colonPos = InStr(startPos, cellText, ":")
                If colonPos > 0 Then
                    endPos = colonPos - 1
                Else
                    endPos = Len(cellText)
                End If
            End If
            
            extractedText = Trim(Mid(cellText, startPos, endPos - startPos + 1))
            ' Store original value for undo before changing
            UndoCount = UndoCount + 1
            UndoArray(UndoCount).CellAddress = cell.Address
            UndoArray(UndoCount).OriginalValue = cellText
            
            cell.Value = extractedText
            processedCount = processedCount + 1
        End If
        ' If not enough segments, leave cell unchanged
        
NextCell:
        On Error GoTo 0 ' Reset error handling
    Next cell
    
    ' Re-enable screen updating to show all changes immediately
    Application.ScreenUpdating = True
    
    ' Show completion message with undo information
    If processedCount > 0 Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Successfully extracted segment " & segmentNumber & " from " & processedCount & " cell(s)!" & vbCrLf & vbCrLf & _
                       "Click OK to keep the dialog open (use Undo button if needed)" & vbCrLf & _
                       "Click Cancel to close the dialog", vbOKCancel + vbInformation, "Process Complete")
        
        ' Close the UserForm if user clicked Cancel
        If result = vbCancel Then
            On Error Resume Next
            Unload TaxonomyCleanerForm
            On Error GoTo 0
        End If
    Else
        MsgBox "No cells were processed. Make sure your selected cells have at least " & segmentNumber & " pipe-delimited segment(s).", vbExclamation, "No Changes Made"
        UndoCount = 0 ' Clear undo data if nothing was processed
    End If
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
End Sub

' Undo the last taxonomy cleaning operation
Sub UndoTaxonomyCleaning()
    Dim i As Integer
    Dim cell As Range
    Dim undoRange As Range
    
    ' Check if there's anything to undo
    If UndoCount = 0 Then
        MsgBox "No taxonomy cleaning operations to undo.", vbInformation, "Nothing to Undo"
        Exit Sub
    End If
    
    ' Confirm undo operation
    Dim operationType As String
    If LastSegmentNumber = 0 Then
        operationType = "Activation ID"
    Else
        operationType = "segment " & LastSegmentNumber
    End If
    
    If MsgBox("This will restore " & UndoCount & " cell(s) to their original values before " & operationType & " extraction." & vbCrLf & vbCrLf & _
              "Do you want to continue?", vbYesNo + vbQuestion, "Confirm Undo") = vbNo Then
        Exit Sub
    End If
    
    ' Disable screen updating for better performance, then re-enable for visual update
    Application.ScreenUpdating = False
    
    ' Restore original values
    For i = 1 To UndoCount
        Set cell = Range(UndoArray(i).CellAddress)
        cell.Value = UndoArray(i).OriginalValue
    Next i
    
    ' Re-enable screen updating to show all changes immediately
    Application.ScreenUpdating = True
    
    ' Clear undo data
    UndoCount = 0
    
    ' Show confirmation
    MsgBox "Successfully restored " & i - 1 & " cell(s) to their original values.", vbInformation, "Undo Complete"
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
End Sub

' Extract Activation ID (text after colon character)
Sub ExtractActivationID()
    Dim cell As Range
    Dim cellText As String
    Dim extractedText As String
    Dim colonPos As Integer
    Dim processedCount As Integer
    
    ' Initialize undo functionality
    UndoCount = 0
    LastSegmentNumber = 0 ' Special marker for Activation ID
    ReDim UndoArray(1 To Selection.Cells.Count)
    
    ' Disable screen updating for better performance, then re-enable for visual update
    Application.ScreenUpdating = False
    
    processedCount = 0
    
    For Each cell In Selection
        On Error GoTo NextCell ' Skip any problematic cells
        cellText = CStr(cell.Value)
        
        ' Skip empty cells
        If Len(Trim(cellText)) = 0 Then
            GoTo NextCell
        End If
        
        ' Find colon position
        colonPos = InStr(cellText, ":")
        
        If colonPos > 0 Then
            ' Extract text after colon (trim any spaces)
            extractedText = Trim(Mid(cellText, colonPos + 1))
            
            ' Store original value for undo before changing
            UndoCount = UndoCount + 1
            UndoArray(UndoCount).CellAddress = cell.Address
            UndoArray(UndoCount).OriginalValue = cellText
            
            cell.Value = extractedText
            processedCount = processedCount + 1
        End If
        ' If no colon found, leave cell unchanged
        
NextCell:
        On Error GoTo 0 ' Reset error handling
    Next cell
    
    ' Re-enable screen updating to show all changes immediately
    Application.ScreenUpdating = True
    
    ' Show completion message with undo information
    If processedCount > 0 Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Successfully extracted Activation ID from " & processedCount & " cell(s)!" & vbCrLf & vbCrLf & _
                       "Click OK to keep the dialog open (use Undo button if needed)" & vbCrLf & _
                       "Click Cancel to close the dialog", vbOKCancel + vbInformation, "Process Complete")
        
        ' Close the UserForm if user clicked Cancel
        If result = vbCancel Then
            On Error Resume Next
            Unload TaxonomyCleanerForm
            On Error GoTo 0
        End If
    Else
        MsgBox "No cells were processed. Make sure your selected cells contain colon (:) characters.", vbExclamation, "No Changes Made"
        UndoCount = 0 ' Clear undo data if nothing was processed
    End If
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
End Sub

' Test function to verify segment extraction works correctly
Sub TestSegmentExtraction()
    Dim testText As String
    
    testText = "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725"
    
    ' Create a test cell
    Range("A1").Value = testText
    Range("A1").Select
    
    MsgBox "Test data placed in A1. You can now test:" & vbCrLf & vbCrLf & _
           "• Segment 8 should extract: 'Facebook_Instagram'" & vbCrLf & _
           "• Segment 9 should extract: 'Conversions'" & vbCrLf & _
           "• Activation ID should extract: 'DJTDOM060725'" & vbCrLf & vbCrLf & _
           "Run TaxonomyCleaner to test these buttons!", vbInformation, "Test Setup Complete"
End Sub