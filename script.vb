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
    
    ' Show the user interface (with fallback for missing UserForm)
    On Error GoTo UseInputBox
    TaxonomyCleanerForm.Show
    Exit Sub
    
UseInputBox:
    ' Simple button-style interface using message boxes
    Call ShowSegmentSelector
End Sub

' Creates a button-style interface for segment selection
Sub ShowSegmentSelector()
    Dim response As VbMsgBoxResult
    Dim selectedSegment As Integer
    
    ' Show welcome and instructions
    response = MsgBox("🚀 TAXONOMY CLEANER MAGIC! 🚀" & vbCrLf & vbCrLf & _
                     "✨ Ready to extract segments from your pipe-delimited data! ✨" & vbCrLf & vbCrLf & _
                     "🎯 Example: 'Marketing|Campaign|Q4|Social|Facebook|Brand|Active|2024'" & vbCrLf & _
                     "   • Segment 1 = 'Marketing'" & vbCrLf & _
                     "   • Segment 3 = 'Q4'" & vbCrLf & _
                     "   • Segment 5 = 'Facebook'" & vbCrLf & vbCrLf & _
                     "💡 Ready to choose your segment?", vbOKCancel + vbInformation, "🎪 Welcome to Segment Paradise! 🎪")
    
    If response = vbCancel Then Exit Sub
    
    ' Show segment selection buttons in groups
    response = MsgBox("🎯 SELECT YOUR SEGMENT (Part 1 of 2)" & vbCrLf & vbCrLf & _
                     "Choose which segment to extract:" & vbCrLf & vbCrLf & _
                     "✅ Click YES for segments 1-4" & vbCrLf & _
                     "✅ Click NO for segments 5-8" & vbCrLf & _
                     "❌ Click CANCEL to exit", vbYesNoCancel + vbQuestion, "🎪 Segment Selector 🎪")
    
    If response = vbCancel Then Exit Sub
    
    If response = vbYes Then
        ' Segments 1-4
        response = MsgBox("🎯 SEGMENTS 1-4 SELECTION" & vbCrLf & vbCrLf & _
                         "Choose your segment:" & vbCrLf & vbCrLf & _
                         "✅ YES = Show segments 1 & 2" & vbCrLf & _
                         "✅ NO = Show segments 3 & 4" & vbCrLf & _
                         "❌ CANCEL = Go back", vbYesNoCancel + vbQuestion, "🔥 Segments 1-4 🔥")
                         
        If response = vbCancel Then Call ShowSegmentSelector: Exit Sub
        
        If response = vbYes Then
            ' Segments 1 & 2
            response = MsgBox("🎯 FINAL CHOICE - Segments 1 & 2" & vbCrLf & vbCrLf & _
                             "✅ YES = Extract SEGMENT 1 (first part)" & vbCrLf & _
                             "✅ NO = Extract SEGMENT 2 (second part)", vbYesNo + vbInformation, "🎊 Almost There! 🎊")
            selectedSegment = IIf(response = vbYes, 1, 2)
        Else
            ' Segments 3 & 4
            response = MsgBox("🎯 FINAL CHOICE - Segments 3 & 4" & vbCrLf & vbCrLf & _
                             "✅ YES = Extract SEGMENT 3 (third part)" & vbCrLf & _
                             "✅ NO = Extract SEGMENT 4 (fourth part)", vbYesNo + vbInformation, "🎊 Almost There! 🎊")
            selectedSegment = IIf(response = vbYes, 3, 4)
        End If
    Else
        ' Segments 5-8
        response = MsgBox("🎯 SEGMENTS 5-8 SELECTION" & vbCrLf & vbCrLf & _
                         "Choose your segment:" & vbCrLf & vbCrLf & _
                         "✅ YES = Show segments 5 & 6" & vbCrLf & _
                         "✅ NO = Show segments 7 & 8" & vbCrLf & _
                         "❌ CANCEL = Go back", vbYesNoCancel + vbQuestion, "🔥 Segments 5-8 🔥")
                         
        If response = vbCancel Then Call ShowSegmentSelector: Exit Sub
        
        If response = vbYes Then
            ' Segments 5 & 6
            response = MsgBox("🎯 FINAL CHOICE - Segments 5 & 6" & vbCrLf & vbCrLf & _
                             "✅ YES = Extract SEGMENT 5 (fifth part)" & vbCrLf & _
                             "✅ NO = Extract SEGMENT 6 (sixth part)", vbYesNo + vbInformation, "🎊 Almost There! 🎊")
            selectedSegment = IIf(response = vbYes, 5, 6)
        Else
            ' Segments 7 & 8
            response = MsgBox("🎯 FINAL CHOICE - Segments 7 & 8" & vbCrLf & vbCrLf & _
                             "✅ YES = Extract SEGMENT 7 (seventh part)" & vbCrLf & _
                             "✅ NO = Extract SEGMENT 8 (eighth part)", vbYesNo + vbInformation, "🎊 Almost There! 🎊")
            selectedSegment = IIf(response = vbYes, 7, 8)
        End If
    End If
    
    ' Execute the extraction
    Call ExtractPipeSegment(selectedSegment)
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
    
    ' Show completion message
    If processedCount > 0 Then
        MsgBox "🎉 Successfully extracted segment " & segmentNumber & " from " & processedCount & " cell(s)!", vbInformation, "Mission Accomplished!"
    Else
        MsgBox "⚠️ No cells were processed. Make sure your selected cells have at least " & segmentNumber & " pipe-delimited segment(s).", vbExclamation, "Nothing to Extract"
    End If
    
    ' Close the form (if it exists)
    On Error Resume Next
    Unload TaxonomyCleanerForm
    On Error GoTo 0
End Sub

'================================================================================
' USERFORM CODE - TaxonomyCleanerForm
' Instructions: Create a UserForm named "TaxonomyCleanerForm" and add the code below
'================================================================================

' UserForm_Initialize - This code goes in the UserForm module
Private Sub UserForm_Initialize()
    ' Set form properties
    Me.Caption = "Taxonomy Cleaner Tool"
    Me.Width = 400
    Me.Height = 350
    
    ' Welcome message (add a Label control named "lblWelcome")
    ' lblWelcome.Caption = "Welcome to the Taxonomy Cleaner Tool!" & vbCrLf & vbCrLf & _
    '                     "This tool extracts text from pipe-delimited data in your selected cells." & vbCrLf & _
    '                     "Choose the number below to extract all text BEFORE that pipe position:" & vbCrLf & vbCrLf & _
    '                     "Example: For 'A|B|C|D|E', button 3 extracts 'A|B|C'"
End Sub

' Button click handlers - Add these to the UserForm module
Private Sub btn1_Click()
    Call ExtractPipeSegment(1)
End Sub

Private Sub btn2_Click()
    Call ExtractPipeSegment(2)
End Sub

Private Sub btn3_Click()
    Call ExtractPipeSegment(3)
End Sub

Private Sub btn4_Click()
    Call ExtractPipeSegment(4)
End Sub

Private Sub btn5_Click()
    Call ExtractPipeSegment(5)
End Sub

Private Sub btn6_Click()
    Call ExtractPipeSegment(6)
End Sub

Private Sub btn7_Click()
    Call ExtractPipeSegment(7)
End Sub

Private Sub btn8_Click()
    Call ExtractPipeSegment(8)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

'================================================================================
' USERFORM DESIGN INSTRUCTIONS
'================================================================================
' 
' To create the UserForm in Excel VBA:
' 1. Open VBA Editor (Alt + F11)
' 2. Right-click on your project → Insert → UserForm
' 3. Name the UserForm "TaxonomyCleanerForm"
' 4. Add the following controls:
'
' LABEL (lblWelcome):
'   - Position: Top of form
'   - Caption: "Welcome to the Taxonomy Cleaner Tool!
'              
'              This tool extracts text from pipe-delimited data in your selected cells.
'              Choose the number below to extract all text BEFORE that pipe position:
'              
'              Example: For 'A|B|C|D|E', button 3 extracts 'A|B|C'"
'   - WordWrap: True
'   - Size: Width=350, Height=120
'
' BUTTONS (btn1 through btn8):
'   - Create 8 CommandButton controls
'   - Name them: btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8
'   - Caption: "1", "2", "3", "4", "5", "6", "7", "8" respectively
'   - Arrange in 2 rows of 4 buttons
'   - Size: Width=60, Height=30 each
'
' CANCEL BUTTON (btnCancel):
'   - Name: btnCancel
'   - Caption: "Cancel"
'   - Position: Bottom center
'   - Size: Width=80, Height=30
'
'================================================================================
