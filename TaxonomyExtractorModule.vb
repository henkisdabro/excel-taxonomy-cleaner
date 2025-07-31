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

Dim UndoArray() As UndoData
Dim UndoCount As Integer
Dim LastSegmentNumber As Integer

' Global variable to hold ribbon reference (optional)
Public myRibbon As IRibbonUI

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
    parsedData = ParseFirstCellData(firstCellContent)
    
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
Function ParseFirstCellData(cellContent As String) As ParsedCellData
    Dim result As ParsedCellData
    
    ' Store original text
    result.OriginalText = cellContent
    
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
    
    ParseFirstCellData = result
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
    
    ' Silent undo operation - no confirmation dialog
    
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
    
    ' Silent completion - no success message
    
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
    
    ' Silent operation - only show errors when nothing processed
    If processedCount = 0 Then
        ' Only show error if nothing was processed
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

'================================================================================
' RIBBON CALLBACK FUNCTIONS
'================================================================================
' These functions are called by the CustomUI ribbon buttons embedded in the XLAM file.
' DO NOT MODIFY the function names - they must match the onAction attributes in customUI.xml

' Ribbon callback function - called when IPG Taxonomy Extractor ribbon button is clicked
Public Sub RibbonTaxonomyExtractor(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' Call the main extractor function
    TaxonomyExtractor
    Exit Sub
    
ErrorHandler:
    MsgBox "Error launching IPG Taxonomy Extractor: " & Err.Description, vbCritical, "IPG Taxonomy Extractor v1.2.0"
End Sub

' Optional: Ribbon load callback - called when the ribbon is loaded
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub