'================================================================================
' EXCEL TAXONOMY EXTRACTOR - Main Module
'================================================================================
' 
' INSTALLATION INSTRUCTIONS:
' 1. Open Excel VBA Editor (Alt + F11)
' 2. Right-click your project → Insert → Module
' 3. Copy and paste this entire code into the new module
' 4. Optionally create the UserForm (see TaxonomyExtractorForm.vb for instructions)
'
' USAGE:
' 1. Select cells with pipe-delimited text (e.g., "Marketing|Campaign|Q4|Social|Facebook")
' 2. Run the TaxonomyExtractor macro (assign to button or use Alt+F8)
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

' Data structure to hold parsed segments from first selected cell
Type ParsedCellData
    OriginalText As String
    TruncatedDisplay As String
    SelectedCellCount As Long
    Segment1 As String
    Segment2 As String
    Segment3 As String
    Segment4 As String
    Segment5 As String
    Segment6 As String
    Segment7 As String
    Segment8 As String
    Segment9 As String
    ActivationID As String
End Type

' Multi-step undo operation structure
Type UndoOperation
    Description As String           ' "Extract Segment 3", "Extract Activation ID"
    CellChanges() As UndoData      ' Array of cell changes for this operation
    CellCount As Integer           ' Number of cells changed in this operation
    OperationId As Integer         ' Unique identifier for debugging
    Timestamp As Date              ' When operation was performed
End Type

' Multi-step undo stack (up to 10 operations)
Dim UndoStack(1 To 10) As UndoOperation
Public UndoOperationCount As Integer    ' Number of operations in stack
Dim NextOperationId As Integer          ' For assigning unique IDs
Public UndoInProgress As Boolean        ' Global flag to prevent rapid clicking

' Legacy variables for backward compatibility during transition
Dim UndoArray() As UndoData
Public UndoCount As Integer
Dim LastSegmentNumber As Integer

' Global variable to hold ribbon reference (optional)
Public myRibbon As Object

' Global variable for modeless form event handling
Public AppEvents As clsAppEvents


' Main macro to be called when button is pressed
Sub TaxonomyExtractor()
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
    
    ' Parse the first selected cell
    Dim firstCellContent As String
    firstCellContent = Selection.Cells(1).Value
    
    ' DEBUG: Show what we're parsing
    Debug.Print "TaxonomyExtractor: First cell content: " & firstCellContent
    
    Dim parsedData As ParsedCellData
    parsedData = ParseFirstCellData(firstCellContent, Selection.Cells.Count)
    
    ' DEBUG: Show parsed results
    Debug.Print "TaxonomyExtractor: Parsed data:"
    Debug.Print "  Original: " & parsedData.OriginalText
    Debug.Print "  Truncated: " & parsedData.TruncatedDisplay
    Debug.Print "  Segment1: " & parsedData.Segment1
    Debug.Print "  Segment2: " & parsedData.Segment2
    Debug.Print "  Segment3: " & parsedData.Segment3
    Debug.Print "  ActivationID: " & parsedData.ActivationID
    
    ' Show the UserForm and pass the parsed data
    Debug.Print "TaxonomyExtractor: Calling SetParsedData..."
    TaxonomyExtractorForm.SetParsedData parsedData
    Debug.Print "TaxonomyExtractor: Showing form..."
    TaxonomyExtractorForm.Show
End Sub

' Modeless version - allows interaction with Excel while form is open
Sub TaxonomyExtractorModeless()
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
    
    ' Initialize application events for selection tracking
    Set AppEvents = New clsAppEvents
    Set AppEvents.App = Application
    
    ' Parse the first selected cell
    Dim firstCellContent As String
    firstCellContent = Selection.Cells(1).Value
    
    Dim parsedData As ParsedCellData
    parsedData = ParseFirstCellData(firstCellContent, Selection.Cells.Count)
    
    ' Show the UserForm as modeless and pass the parsed data
    TaxonomyExtractorForm.SetParsedData parsedData
    TaxonomyExtractorForm.Show vbModeless
    
    ' Give the UserForm immediate focus and set focus to first available button
    On Error Resume Next
    ' Set focus to the first available segment button for immediate usability
    If TaxonomyExtractorForm.btn1.Enabled Then
        TaxonomyExtractorForm.btn1.SetFocus
    ElseIf TaxonomyExtractorForm.btnActivationID.Enabled Then
        TaxonomyExtractorForm.btnActivationID.SetFocus
    Else
        ' Fallback to Close button if no segments available
        TaxonomyExtractorForm.btnClose.SetFocus
    End If
    On Error GoTo 0
End Sub

' Simple input dialog interface (fallback when UserForm doesn't exist)
Sub ShowSegmentSelector()
    Dim selectedSegment As String
    Dim validNumber As Integer
    
    ' Show clean, simple interface
    selectedSegment = InputBox("TAXONOMY EXTRACTOR - Segment Extractor" & vbCrLf & vbCrLf & _
                              "This tool extracts specific segments from pipe-delimited data." & vbCrLf & vbCrLf & _
                              "EXAMPLE: 'FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725'" & vbCrLf & _
                              "  Segment 1 = FY24_26" & vbCrLf & _
                              "  Segment 3 = Tourism WA" & vbCrLf & _
                              "  Segment 5 = Always On Remarketing" & vbCrLf & _
                              "  Segment 9 = Conversions" & vbCrLf & _
                              "  A = DJTDOM060725 (Activation ID)" & vbCrLf & vbCrLf & _
                              "Enter segment number (1-9) or 'A' for Activation ID:", "Taxonomy Extractor", "")
    
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

' Parse first selected cell into individual segments
Function ParseFirstCellData(cellContent As String, selectedCellCount As Long) As ParsedCellData
    Dim result As ParsedCellData
    
    ' Store original text and cell count
    result.OriginalText = cellContent
    result.SelectedCellCount = selectedCellCount
    
    ' Create truncated display (12 chars + "...")
    If Len(cellContent) > 15 Then
        result.TruncatedDisplay = Left(cellContent, 12) & "..."
    Else
        result.TruncatedDisplay = cellContent
    End If
    
    ' Split by colon first to separate activation ID
    Dim colonParts() As String
    colonParts = Split(cellContent, ":")
    
    Dim mainContent As String
    mainContent = colonParts(0)
    
    ' Extract activation ID if colon exists
    If UBound(colonParts) > 0 Then
        result.ActivationID = Trim(colonParts(1))
    Else
        result.ActivationID = ""
    End If
    
    ' Only parse segments if there are actual pipe characters
    ' Without pipes, this is not taxonomy data and all segments should be empty
    If InStr(mainContent, "|") > 0 Then
        ' Split main content by pipes
        Dim segments() As String
        segments = Split(mainContent, "|")
        
        ' Assign segments (with bounds checking)
        If UBound(segments) >= 0 Then result.Segment1 = Trim(segments(0))
        If UBound(segments) >= 1 Then result.Segment2 = Trim(segments(1))
        If UBound(segments) >= 2 Then result.Segment3 = Trim(segments(2))
        If UBound(segments) >= 3 Then result.Segment4 = Trim(segments(3))
        If UBound(segments) >= 4 Then result.Segment5 = Trim(segments(4))
        If UBound(segments) >= 5 Then result.Segment6 = Trim(segments(5))
        If UBound(segments) >= 6 Then result.Segment7 = Trim(segments(6))
        If UBound(segments) >= 7 Then result.Segment8 = Trim(segments(7))
        If UBound(segments) >= 8 Then result.Segment9 = Trim(segments(8))
    Else
        ' No pipes found - leave all segments empty (they default to empty strings)
        ' This will cause all buttons to show "N/A" as intended
    End If
    
    ParseFirstCellData = result
End Function

' Extract the first targeting pattern found in text (for button caption display)
Function ExtractTargetingPattern(inputText As String) As String
    Dim regex As Object
    Dim matches As Object
    
    On Error GoTo ErrorHandler
    
    ' Create regex object for pattern matching
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False  ' Only find first match
        .Pattern = "\^[^^]+\^"  ' Matches ^any characters except caret^ (no optional space for display)
    End With
    
    ' Find the first match
    Set matches = regex.Execute(inputText)
    
    If matches.Count > 0 Then
        ExtractTargetingPattern = matches(0).Value
    Else
        ExtractTargetingPattern = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractTargetingPattern = ""
End Function

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
    Dim colonPos As Integer
    
    ' Add operation to undo stack BEFORE making changes
    Call AddUndoOperation("Extract Segment " & segmentNumber)
    
    ' Legacy variables for backward compatibility
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
        If segmentNumber <= pipeCount + 1 Then
            Dim startPos As Integer
            Dim endPos As Integer
            
            If segmentNumber = 1 Then
                ' First segment: from start to first pipe (or entire text if no pipes)
                If pipeCount >= 1 Then
                    startPos = 1
                    endPos = pipePositions(1) - 1
                Else
                    ' No pipes, check for colon
                    startPos = 1
                    colonPos = InStr(cellText, ":")
                    If colonPos > 0 Then
                        endPos = colonPos - 1
                    Else
                        endPos = Len(cellText)
                    End If
                End If
            ElseIf segmentNumber <= pipeCount Then
                ' Middle segments: between two pipes
                startPos = pipePositions(segmentNumber - 1) + 1
                endPos = pipePositions(segmentNumber) - 1
            Else
                ' Last segment: after final pipe, but before colon if present
                startPos = pipePositions(pipeCount) + 1
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
    
    ' Silent operation - only show errors when nothing processed
    If processedCount = 0 Then
        ' Only show error if nothing was processed
        MsgBox "No cells were processed. Make sure your selected cells have at least " & segmentNumber & " pipe-delimited segment(s).", vbExclamation, "No Changes Made"
        UndoCount = 0 ' Clear undo data if nothing was processed
    End If
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
    
    ' Refresh modeless UserForm if it's open (v1.4.0 enhancement)
    Call RefreshModelessFormIfOpen
End Sub

' Undo the last taxonomy cleaning operation
Sub UndoTaxonomyCleaning()
    ' Legacy function redirected to new multi-step undo system
    Call UndoLastOperation
End Sub

' Extract Activation ID (text after colon character)
Sub ExtractActivationID()
    Dim cell As Range
    Dim cellText As String
    Dim extractedText As String
    Dim colonPos As Integer
    Dim processedCount As Integer
    
    ' Add operation to undo stack BEFORE making changes
    Call AddUndoOperation("Extract Activation ID")
    
    ' Legacy variables for backward compatibility
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
    
    ' Silent operation - only show errors when nothing processed
    If processedCount = 0 Then
        ' Only show error if nothing was processed
        MsgBox "No cells were processed. Make sure your selected cells contain colon (:) characters.", vbExclamation, "No Changes Made"
        UndoCount = 0 ' Clear undo data if nothing was processed
    End If
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
    
    ' Refresh modeless UserForm if it's open (v1.4.0 enhancement)
    Call RefreshModelessFormIfOpen
End Sub

' Clean targeting acronyms (removes text in format ^ABC^ with optional trailing space)
Sub CleanTargetingAcronyms()
    Dim cell As Range
    Dim cellText As String
    Dim cleanedText As String
    Dim processedCount As Integer
    Dim regex As Object
    
    ' Add operation to undo stack BEFORE making changes
    Call AddUndoOperation("Clean Targeting Acronyms")
    
    ' Legacy variables for backward compatibility
    UndoCount = 0
    LastSegmentNumber = -1 ' Special marker for targeting acronym cleaning
    ReDim UndoArray(1 To Selection.Cells.Count)
    
    ' Create regex object for pattern matching
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "\^[^^]+\^ ?" ' Matches ^any characters except caret^ with optional trailing space
    End With
    
    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    
    processedCount = 0
    
    For Each cell In Selection
        On Error GoTo NextCell ' Skip any problematic cells
        cellText = CStr(cell.Value)
        
        ' Skip empty cells
        If Len(Trim(cellText)) = 0 Then
            GoTo NextCell
        End If
        
        ' Check if cell contains targeting acronym pattern
        If regex.Test(cellText) Then
            ' Store original value for undo before changing
            UndoCount = UndoCount + 1
            UndoArray(UndoCount).CellAddress = cell.Address
            UndoArray(UndoCount).OriginalValue = cellText
            
            ' Remove all targeting acronym patterns
            cleanedText = regex.Replace(cellText, "")
            cell.Value = cleanedText
            processedCount = processedCount + 1
        End If
        
NextCell:
        On Error GoTo 0 ' Reset error handling
    Next cell
    
    ' Re-enable screen updating to show all changes immediately
    Application.ScreenUpdating = True
    
    ' Silent operation - only show errors when nothing processed
    If processedCount = 0 Then
        ' Only show error if nothing was processed
        MsgBox "No cells were processed. Make sure your selected cells contain targeting acronyms in format ^ABC^.", vbExclamation, "No Changes Made"
        UndoCount = 0 ' Clear undo data if nothing was processed
    End If
    
    ' Ensure screen updating is always re-enabled
    Application.ScreenUpdating = True
    
    ' Refresh modeless UserForm if it's open
    Call RefreshModelessFormIfOpen
End Sub

' Refresh modeless UserForm after extraction (v1.4.0 UX enhancement)
Sub RefreshModelessFormIfOpen()
    On Error GoTo ErrorHandler
    
    ' Only refresh if UserForm exists and is visible (modeless mode)
    If Not TaxonomyExtractorForm Is Nothing Then
        If TaxonomyExtractorForm.Visible Then
            ' Get current selection and update form with new content
            If Selection.Cells.Count > 0 Then
                Dim firstCellContent As String
                firstCellContent = Selection.Cells(1).Value
                
                ' Parse the updated cell content (now likely a single value, no pipes)
                Dim updatedData As ParsedCellData
                updatedData = ParseFirstCellData(firstCellContent, Selection.Cells.Count)
                
                ' Update the form with new data
                TaxonomyExtractorForm.SetParsedData updatedData
                
                Debug.Print "RefreshModelessFormIfOpen: Updated form after extraction with: " & firstCellContent
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Silent error handling - don't interrupt user workflow
    Debug.Print "RefreshModelessFormIfOpen Error: " & Err.Description
End Sub

' Test function to verify segment extraction works correctly
Sub TestSegmentExtraction()
    Dim testText As String
    
    testText = "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725"
    
    ' Create test cells
    Range("A1").Value = testText
    Range("A2").Value = "Test|Without|Colon|Data" ' Test without colon
    Range("A1:A2").Select
    
    MsgBox "Test data placed in A1:A2. COMPILATION ERROR FIXED!" & vbCrLf & vbCrLf & _
           "A1 (with colon):" & vbCrLf & _
           "• Segment 8 should extract: 'Facebook_Instagram'" & vbCrLf & _
           "• Segment 9 should extract: 'Conversions'" & vbCrLf & _
           "• Activation ID should extract: 'DJTDOM060725'" & vbCrLf & vbCrLf & _
           "A2 (without colon):" & vbCrLf & _
           "• Segment 4 should extract: 'Data'" & vbCrLf & _
           "• Activation ID should show 'no colon' message" & vbCrLf & vbCrLf & _
           "Run TaxonomyExtractor to test these buttons!", vbInformation, "Test Setup Complete"
End Sub

' Quick test of Activation ID extraction directly
Sub TestActivationIDDirect()
    Range("A1").Value = "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725"
    Range("A1").Select
    Call ExtractActivationID
End Sub

' Test simple positioning - centers UserForm in Excel window
Sub TestSimplePositioning()
    ' Create test data
    Range("B2").Value = "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725"
    Range("B2").Select
    
    MsgBox "SIMPLE POSITIONING TEST:" & vbCrLf & vbCrLf & _
           "The UserForm will now appear centered within the Excel window." & vbCrLf & _
           "This uses Excel's Application.Left, .Top, .Width, and .Height properties" & vbCrLf & _
           "to calculate the center position reliably.", _
           vbInformation, "Simple Positioning Test"
    
    ' Launch the form to test positioning
    Call TaxonomyExtractor
End Sub

' Test targeting acronym cleaning functionality with smart button behavior and expanded patterns
Sub TestTargetingAcronymCleaning()
    ' Create different types of test data to show smart button behavior with various patterns
    Range("A1").Value = "FY24_26|Q1-4|Tourism WA|WA|Marketing:ABC123"  ' Taxonomy data with pipes
    Range("A2").Value = "^AT^ testing string"                            ' Simple letters pattern
    Range("A3").Value = "^ACX123^Acxiom Targeting"                      ' Letters + numbers pattern 
    Range("A4").Value = "^FB_Campaign^ Facebook data"                   ' Letters + underscore pattern
    Range("A5").Value = "^Multi-Word^ test content"                     ' Multi-word with hyphen pattern
    Range("A6").Value = "^123ABC^ numeric start pattern"               ' Numbers + letters pattern
    Range("A7").Value = "No acronyms here"                              ' Regular text (no pattern)
    Range("A8").Value = "Regular taxonomy data without targeting"        ' Regular text (no pattern)
    
    MsgBox "EXPANDED TARGETING PATTERN TEST:" & vbCrLf & vbCrLf & _
           "Test data placed in A1:A8. Select each row to see smart button visibility:" & vbCrLf & vbCrLf & _
           "A1: 'FY24_26|Q1-4|...' → Trim button HIDDEN (has pipes)" & vbCrLf & _
           "A2: '^AT^ testing string' → Trim button VISIBLE: ^AT^" & vbCrLf & _
           "A3: '^ACX123^Acxiom...' → Trim button VISIBLE: ^ACX123^" & vbCrLf & _
           "A4: '^FB_Campaign^ Facebook...' → Trim button VISIBLE: ^FB_Campaign^" & vbCrLf & _
           "A5: '^Multi-Word^ test...' → Trim button VISIBLE: ^Multi-Word^" & vbCrLf & _
           "A6: '^123ABC^ numeric...' → Trim button VISIBLE: ^123ABC^" & vbCrLf & _
           "A7: 'No acronyms here' → Trim button HIDDEN (no pattern)" & vbCrLf & _
           "A8: 'Regular taxonomy...' → Trim button HIDDEN (no pattern)" & vbCrLf & vbCrLf & _
           "Button overlays Segment 1 and only appears when needed! Use MODELESS mode!", _
           vbInformation, "Expanded Targeting Pattern Test"
    
    ' Select first row and launch modeless form for easy testing
    Range("A1").Select
    Call TaxonomyExtractorModeless
End Sub

'================================================================================
' RIBBON CALLBACK FUNCTIONS
'================================================================================
' These functions are called by the CustomUI ribbon buttons embedded in the XLAM file.
' DO NOT MODIFY the function names - they must match the onAction attributes in customUI.xml

' Ribbon callback function - called when IPG Taxonomy Extractor ribbon button is clicked
Public Sub RibbonTaxonomyExtractor(control As Object)
    On Error GoTo ErrorHandler
    
    ' Call the modeless extractor function (v1.4.0 - superior user experience)
    TaxonomyExtractorModeless
    Exit Sub
    
ErrorHandler:
    MsgBox "Error launching IPG Taxonomy Extractor: " & Err.Description, vbCritical, "IPG Taxonomy Extractor v1.6.0"
End Sub

' Ribbon callback function - called when IPG Taxonomy Extractor (Modeless) ribbon button is clicked
Public Sub RibbonTaxonomyExtractorModeless(control As Object)
    On Error GoTo ErrorHandler
    
    ' Call the modeless extractor function
    TaxonomyExtractorModeless
    Exit Sub
    
ErrorHandler:
    MsgBox "Error launching IPG Taxonomy Extractor (Modeless): " & Err.Description, vbCritical, "IPG Taxonomy Extractor v1.6.0"
End Sub

' Cleanup function for modeless form - called when UserForm is closed
Public Sub CleanupModelessEvents()
    On Error Resume Next
    If Not AppEvents Is Nothing Then
        AppEvents.Cleanup
        Set AppEvents = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' MULTI-STEP UNDO SYSTEM
' ============================================================================

' Add a new operation to the undo stack, capturing current cell states BEFORE changes
Public Sub AddUndoOperation(description As String)
    On Error GoTo ErrorHandler
    
    ' Increment operation count, managing 10-operation limit
    If UndoOperationCount >= 10 Then
        ' Remove oldest operation (shift array left)
        Dim i As Integer
        For i = 1 To 9
            UndoStack(i) = UndoStack(i + 1)
        Next i
        UndoOperationCount = 9
    End If
    
    UndoOperationCount = UndoOperationCount + 1
    NextOperationId = NextOperationId + 1
    
    ' Initialize the new operation
    With UndoStack(UndoOperationCount)
        .Description = description
        .OperationId = NextOperationId
        .Timestamp = Now
        .CellCount = 0
        
        ' Capture current state of all selected cells BEFORE any changes
        ReDim .CellChanges(1 To Selection.Cells.Count)
        
        Dim cell As Range
        Dim cellIndex As Integer
        cellIndex = 0
        
        For Each cell In Selection
            If Len(cell.Value) > 0 Then  ' Only capture cells with content
                cellIndex = cellIndex + 1
                .CellChanges(cellIndex).CellAddress = cell.Address
                .CellChanges(cellIndex).OriginalValue = cell.Value
                .CellCount = .CellCount + 1
            End If
        Next cell
        
        ' Resize array to actual count
        If .CellCount > 0 Then
            ReDim Preserve .CellChanges(1 To .CellCount)
        End If
    End With
    
    Debug.Print "AddUndoOperation: Added '" & description & "' (ID: " & NextOperationId & ") affecting " & UndoStack(UndoOperationCount).CellCount & " cells"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "AddUndoOperation Error: " & Err.Description
    ' Don't let undo system errors interrupt operations
End Sub

' Undo the most recent operation (LIFO - Last In, First Out)
Public Sub UndoLastOperation()
    On Error GoTo ErrorHandler
    
    ' Check if there are operations to undo
    If UndoOperationCount = 0 Then
        MsgBox "No operations to undo.", vbInformation, "Nothing to Undo"
        Exit Sub
    End If
    
    ' Button is already disabled by the click handler
    ' Just proceed with the undo operation
    On Error GoTo ErrorHandler
    
    ' Get the most recent operation
    Dim currentOp As UndoOperation
    currentOp = UndoStack(UndoOperationCount)
    
    Debug.Print "UndoLastOperation: Undoing '" & currentOp.Description & "' (ID: " & currentOp.OperationId & ")"
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' Restore all cells from this operation
    Dim i As Integer
    Dim cell As Range
    
    For i = 1 To currentOp.CellCount
        On Error Resume Next  ' Handle invalid cell references gracefully
        Set cell = Range(currentOp.CellChanges(i).CellAddress)
        If Not cell Is Nothing Then
            cell.Value = currentOp.CellChanges(i).OriginalValue
            Debug.Print "  Restored " & cell.Address & " to: " & currentOp.CellChanges(i).OriginalValue
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' Remove this operation from the stack
    UndoOperationCount = UndoOperationCount - 1
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    Debug.Print "UndoLastOperation: Complete. Operations remaining: " & UndoOperationCount
    
    ' Refresh the UI to show updated state
    Call RefreshModelessFormIfOpen
    
    ' Re-enable undo button after brief delay to prevent rapid clicking issues
    Call ReenableUndoButtonAfterDelay
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "UndoLastOperation Error: " & Err.Description
    MsgBox "Error during undo operation: " & Err.Description, vbCritical, "Undo Error"
    ' Clear flags and re-enable undo button immediately if error occurred
    UndoInProgress = False
    On Error Resume Next
    If Not TaxonomyExtractorForm Is Nothing Then
        If TaxonomyExtractorForm.Visible Then
            TaxonomyExtractorForm.Tag = ""
            TaxonomyExtractorForm.btnUndo.Enabled = True
            Call TaxonomyExtractorForm.UpdateUndoButtonState
        End If
    End If
    On Error GoTo 0
End Sub

' Re-enable undo button after a brief delay to prevent rapid clicking issues
Private Sub ReenableUndoButtonAfterDelay()
    ' Use a simple timer approach with Application.Wait for 500ms delay
    Dim startTime As Double
    startTime = Timer
    
    ' Brief delay to allow cell operations to complete and provide visual feedback
    Application.Wait Now + TimeValue("00:00:01")  ' 1 second delay for better responsiveness
    
    ' Re-enable the undo button and restore proper state
    On Error Resume Next
    If Not TaxonomyExtractorForm Is Nothing Then
        If TaxonomyExtractorForm.Visible Then
            ' Clear processing flags
            TaxonomyExtractorForm.Tag = ""
            UndoInProgress = False
            ' Re-enable the button first
            TaxonomyExtractorForm.btnUndo.Enabled = True
            ' Update button state will handle caption and colors correctly
            Call TaxonomyExtractorForm.UpdateUndoButtonState
            ' Restore focus to undo button after re-enabling
            If TaxonomyExtractorForm.btnUndo.Enabled Then
                TaxonomyExtractorForm.btnUndo.SetFocus
            End If
            Debug.Print "ReenableUndoButtonAfterDelay: Focus restored to Undo button"
            Debug.Print "ReenableUndoButtonAfterDelay: Undo button re-enabled"
        End If
    End If
    On Error GoTo 0
End Sub

' Clear all operations from the undo stack
Public Sub ClearUndoStack()
    UndoOperationCount = 0
    NextOperationId = 0
    Debug.Print "ClearUndoStack: All undo operations cleared"
End Sub

' Get description of the most recent operation (for UI display)
Public Function GetTopUndoDescription() As String
    If UndoOperationCount > 0 Then
        GetTopUndoDescription = UndoStack(UndoOperationCount).Description
    Else
        GetTopUndoDescription = ""
    End If
End Function

' Optional: Ribbon load callback - called when the ribbon is loaded
Public Sub RibbonOnLoad(ribbon As Object)
    Set myRibbon = ribbon
End Sub