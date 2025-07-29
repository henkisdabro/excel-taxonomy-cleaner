'================================================================================
' EXCEL TAXONOMY CLEANER - Enhanced Version
'================================================================================
' 
' QUICK START INSTRUCTIONS:
' 1. Copy this entire code into a VBA Module in Excel
' 2. Select cells with pipe-delimited text (e.g., "A|B|C|D|E")
' 3. Run the TaxonomyCleaner macro
' 4. Choose a number (1-8) to extract THAT SPECIFIC SEGMENT
'
' ADVANCED SETUP (Optional - for awesome button interface):
' - Create a UserForm named "TaxonomyCleanerForm" following instructions at bottom
' - Get 8 beautiful buttons instead of boring input dialog!
'
' EXAMPLES:
' For text "Marketing|Campaign|Q4|Social|Facebook|Brand|Active|2024":
' - Button/Number 1: "Marketing" (1st segment)
' - Button/Number 3: "Q4" (3rd segment)
' - Button/Number 5: "Facebook" (5th segment)
' - Button/Number 8: "2024" (8th segment)
'================================================================================

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
    
    ' Show the simple interface (for real buttons, create UserForm manually)
    TaxonomyCleanerForm.Show
End Sub


' Simple single-dialog interface for segment selection
Sub ShowSegmentSelector()
    Dim selectedSegment As String
    Dim validNumber As Integer
    
    ' Show clean, simple interface
    selectedSegment = InputBox("TAXONOMY CLEANER - Segment Extractor" & vbCrLf & vbCrLf & _
                              "This tool extracts specific segments from pipe-delimited data." & vbCrLf & vbCrLf & _
                              "EXAMPLE: For 'Marketing|Campaign|Q4|Social|Facebook|Brand|Active|2024'" & vbCrLf & _
                              "  Segment 1 = Marketing" & vbCrLf & _
                              "  Segment 3 = Q4" & vbCrLf & _
                              "  Segment 5 = Facebook" & vbCrLf & _
                              "  Segment 8 = 2024" & vbCrLf & vbCrLf & _
                              "Enter segment number (1-8):", "Taxonomy Cleaner", "")
    
    ' Validate and execute
    If selectedSegment = "" Then Exit Sub ' User cancelled
    
    If IsNumeric(selectedSegment) Then
        validNumber = CInt(selectedSegment)
        If validNumber >= 1 And validNumber <= 8 Then
            Call ExtractPipeSegment(validNumber)
        Else
            MsgBox "Please enter a number between 1 and 8.", vbExclamation, "Invalid Input"
        End If
    Else
        MsgBox "Please enter a valid number between 1 and 8.", vbExclamation, "Invalid Input"
    End If
End Sub

' Function to extract text between pipes (segment n)
Sub ExtractPipeSegment(segmentNumber As Integer)
    Dim cell As Range
    Dim cellText As String
    Dim extractedText As String
    Dim pipePositions(1 To 10) As Integer
    Dim pipeCount As Integer
    Dim pos As Integer
    Dim processedCount As Integer
    Dim i As Integer
    
    processedCount = 0
    
    For Each cell In Selection
        cellText = cell.Value
        
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
                extractedText = Left(cellText, pipePositions(1) - 1)
            Else
                extractedText = cellText ' No pipes, use entire text
            End If
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
                ' Last segment after final pipe
                startPos = pipePositions(pipeCount) + 1
                endPos = Len(cellText)
            End If
            
            extractedText = Mid(cellText, startPos, endPos - startPos + 1)
            cell.Value = extractedText
            processedCount = processedCount + 1
        End If
        ' If not enough segments, leave cell unchanged
        
NextCell:
    Next cell
    
    ' Show completion message ONLY (don't unload form here)
    If processedCount > 0 Then
        MsgBox "Successfully extracted segment " & segmentNumber & " from " & processedCount & " cell(s)!", vbInformation, "Process Complete"
    Else
        MsgBox "No cells were processed. Make sure your selected cells have at least " & segmentNumber & " pipe-delimited segment(s).", vbExclamation, "No Changes Made"
    End If
End Sub

'================================================================================
' OPTIONAL: ADVANCED USERFORM WITH 8 BUTTONS
'================================================================================
' For the ultimate experience, you can create a UserForm with actual buttons.
' This requires manual setup but gives you the real button interface you want.
'
' STEP 1: Create the UserForm
' - Open VBA Editor (Alt + F11)
' - Right-click your project → Insert → UserForm  
' - Name it "TaxonomyCleanerForm"
' - Set properties: Width=400, Height=300
'
' STEP 2: Add controls to the UserForm:
' 
' 1. Add a LABEL at the top:
'    - Name: lblInstructions
'    - Caption: "Select segment to extract from pipe-delimited data:"
'    - Position: Top center
'    - Size: Width=350, Height=40
'
' 2. Add 8 BUTTONS in two rows:
'    Row 1 (buttons 1-4): Y=80, Heights=40, Width=70 each
'    - btn1: X=30,  Caption="Segment 1"
'    - btn2: X=110, Caption="Segment 2" 
'    - btn3: X=190, Caption="Segment 3"
'    - btn4: X=270, Caption="Segment 4"
'    
'    Row 2 (buttons 5-8): Y=130, Heights=40, Width=70 each  
'    - btn5: X=30,  Caption="Segment 5"
'    - btn6: X=110, Caption="Segment 6"
'    - btn7: X=190, Caption="Segment 7" 
'    - btn8: X=270, Caption="Segment 8"
'
' 3. Add CANCEL button:
'    - btnCancel: X=160, Y=200, Width=80, Height=30, Caption="Cancel"
'
' STEP 3: Add this VBA code to the UserForm module:

Private Sub UserForm_Initialize()
    Me.Caption = "Taxonomy Cleaner - Segment Selector"
End Sub

Private Sub btn1_Click(): Call ExtractPipeSegment(1): Unload Me: End Sub
Private Sub btn2_Click(): Call ExtractPipeSegment(2): Unload Me: End Sub  
Private Sub btn3_Click(): Call ExtractPipeSegment(3): Unload Me: End Sub
Private Sub btn4_Click(): Call ExtractPipeSegment(4): Unload Me: End Sub
Private Sub btn5_Click(): Call ExtractPipeSegment(5): Unload Me: End Sub
Private Sub btn6_Click(): Call ExtractPipeSegment(6): Unload Me: End Sub
Private Sub btn7_Click(): Call ExtractPipeSegment(7): Unload Me: End Sub
Private Sub btn8_Click(): Call ExtractPipeSegment(8): Unload Me: End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub

' STEP 4: Update the main function to use the form:
' Replace "TaxonomyCleanerForm.Show" with "TaxonomyCleanerForm.Show"
'
'================================================================================
