'================================================================================
' EXCEL TAXONOMY EXTRACTOR - UserForm Code (TaxonomyExtractorForm)
'================================================================================
' 
' This file contains the VBA code for the working UserForm with 9 segment buttons + Activation ID button.
' This is the code that should be placed in your UserForm named "TaxonomyExtractorForm"

' Windows API functions for screen metrics (must be at module level)
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
'
' USERFORM SETUP (Already completed if your form is working):
' ==========================================================
' 
' UserForm Name: TaxonomyExtractorForm
' UserForm Properties:
'    - Width: 480
'    - Height: 320
'    - Caption: "Taxonomy Extractor - Select Segment"
'
' VERSION MANAGEMENT:
' ===================
' The UserForm_Initialize() code below sets the title to include version number (v1.1.0).
' IMPORTANT: Increment version number on each significant code update:
'   - v1.0.0: Initial release with 9 buttons + Activation ID
'   - v1.1.0: Added smart data preview and dynamic button captions
'   - v1.2.0: Enhanced undo functionality
'   - v1.6.0: Multi-step undo system with professional UX enhancements
' This helps track which version users are running in their Excel environment.
'
' CONTROLS (MUST be named exactly as shown):
' ===========================================
' 
' REQUIRED CONTROLS FOR SMART INTERFACE:
' 
' 1. INSTRUCTION LABEL (for data preview):
'    - Control Type: Label
'    - Name: lblInstructions
'    - Caption: "Select cells and click segment button"
'    - Position: Top of form (X: 12, Y: 12)
'    - Size: Width: 450, Height: 24
'    - Font: Calibri, 10pt
'    - Important: This label will show truncated data preview automatically
'
' 2. CELL COUNT LABEL (optional - shows number of selected cells):
'    - Control Type: Label
'    - Name: lblCellCount
'    - Caption: "Processing: 1 cells"
'    - Position: Below instructions (X: 12, Y: 36)
'    - Size: Width: 200, Height: 18
'    - Font: Calibri, 9pt
'    - Important: Access via cellData.SelectedCellCount
'
' 3. SEGMENT BUTTONS (9 buttons for segments 1-9):
'    - Control Type: CommandButton
'    - Names: btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9
'    - Default Captions: "Segment 1", "Segment 2", etc.
'    - Note: Captions will update automatically to show segment previews
'    - Suggested Layout: 3 rows x 3 columns
'      Row 1: btn1, btn2, btn3 (Y: 50)
'      Row 2: btn4, btn5, btn6 (Y: 90)
'      Row 3: btn7, btn8, btn9 (Y: 130)
'    - Size: Width: 140, Height: 30
'    - Spacing: X positions: 12, 164, 316
'
' 4. ACTIVATION ID BUTTON:
'    - Control Type: CommandButton
'    - Name: btnActivationID
'    - Default Caption: "Activation ID"
'    - Note: Caption will update to show "ID: [preview]"
'    - Position: X: 12, Y: 180
'    - Size: Width: 140, Height: 30
'
' 5. TARGETING ACRONYM CLEAN CONTROLS (overlaid on Segment 1 button):
'    - Control Type: CommandButton
'    - Name: TargetingAcronymCleanButton
'    - Caption: "Trim ^ABC^" (dynamically updated to show actual pattern)
'    - Position: Same as btn1 (overlaid on top)
'    - Size: Same as btn1
'    - Visible: Only when no pipes present AND caret pattern found
'    - Note: Removes text in format ^any_characters^ with optional trailing space
'    
'    - Control Type: Label (optional companion label)
'    - Name: AcronymCleanLabel
'    - Caption: Any descriptive text you want
'    - Position: Near TargetingAcronymCleanButton
'    - Visible: Same logic as button (only when targeting pattern detected)
'
' 6. ACTION BUTTONS:
'    - Control Type: CommandButton (2 buttons)
'    - Names: btnUndo, btnClose
'    - Captions: "Undo Last", "Close"
'    - Position: X: 12, Y: 220 (btnUndo), X: 240, Y: 220 (btnClose)
'    - Size: Width: 68, Height: 30
'    - Undo Button: Dynamically enabled/disabled with operation count (e.g., "Undo Last (3)")
'
' 7. UNDO CAPACITY WARNING LABEL:
'    - Control Type: Label
'    - Name: lblUndoWarning
'    - Caption: "⚠️ Undo limit reached (10/10). Oldest operations will be removed."
'    - Position: X: 12, Y: 252 (below Undo button)
'    - Size: Width: 350, Height: 20
'    - Font: 8pt, ForeColor: RGB(200, 100, 0) (orange warning color)
'    - Visible: Only when UndoOperationCount = 10
'
' LAYOUT SUMMARY:
' - Form dimensions: 480 x 280
' - lblInstructions shows: "Selected: [complete original text - no truncation]"
' - Segment buttons show: "1: [12 chars]", "2: [12 chars]", etc. (or "N/A" if missing)
' - ID button shows: "ID: [full activation ID]" (or "ID: N/A" if missing)
' - All previews update automatically when UserForm opens
'
' VBA CODE FOR THE USERFORM:
' ==========================
' After creating all controls above, copy and paste this code into the UserForm module (double-click TaxonomyExtractorForm):
' 
' IMPORTANT: The code below expects the exact control names listed above.
' If lblInstructions doesn't exist, the line lblInstructions.Caption will cause an error.
' Make sure all controls are created and named correctly before adding this code.

Private cellData As ParsedCellData

Public Sub SetParsedData(parsedData As ParsedCellData)
    cellData = parsedData
    
    ' DEBUG: Show what data was received
    Debug.Print "SetParsedData called with:"
    Debug.Print "  Original: " & cellData.OriginalText
    Debug.Print "  Truncated: " & cellData.TruncatedDisplay
    Debug.Print "  Segment1: " & cellData.Segment1
    Debug.Print "  Segment2: " & cellData.Segment2
    Debug.Print "  Segment3: " & cellData.Segment3
    Debug.Print "  ActivationID: " & cellData.ActivationID
    
    ' Update the interface immediately after receiving data
    UpdateInterface
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "IPG Mediabrands Taxonomy Extractor v1.6.0"
    
    ' DEBUG: Check if cellData has been populated
    Debug.Print "UserForm_Initialize called"
    Debug.Print "  cellData.OriginalText length: " & Len(cellData.OriginalText)
    
    ' Apply simple, reliable positioning - center within Excel window
    ApplyOptimalPositioning
    
    ' Note: SetParsedData will handle interface updates
    ' Don't try to update interface here as cellData may not be set yet
End Sub

Private Sub ApplyOptimalPositioning()
    ' Center the form within Excel's main window on the correct monitor
    ' Respects the UserForm's design-time Width and Height properties
    ' Includes bounds checking to prevent off-screen positioning
    
    On Error GoTo CenterOnScreen
    
    ' Get Excel application window position and size (this is the main Excel window)
    Dim excelLeft As Long, excelTop As Long, excelWidth As Long, excelHeight As Long
    excelLeft = Application.Left
    excelTop = Application.Top
    excelWidth = Application.Width
    excelHeight = Application.Height
    
    ' Get the form's design-time dimensions (don't override them)
    Dim formWidth As Long, formHeight As Long
    formWidth = Me.Width
    formHeight = Me.Height
    
    ' Calculate center position within Excel window
    Dim centerLeft As Long, centerTop As Long
    centerLeft = excelLeft + (excelWidth - formWidth) / 2
    centerTop = excelTop + (excelHeight - formHeight) / 2
    
    ' Get system metrics for bounds checking
    Dim screenWidth As Long, screenHeight As Long
    screenWidth = GetSystemMetrics(0)  ' SM_CXSCREEN - primary monitor width
    screenHeight = GetSystemMetrics(1) ' SM_CYSCREEN - primary monitor height
    
    ' For multi-monitor setups, we need virtual screen dimensions
    Dim virtualLeft As Long, virtualTop As Long, virtualWidth As Long, virtualHeight As Long
    virtualLeft = GetSystemMetrics(76)   ' SM_XVIRTUALSCREEN
    virtualTop = GetSystemMetrics(77)    ' SM_YVIRTUALSCREEN  
    virtualWidth = GetSystemMetrics(78)  ' SM_CXVIRTUALSCREEN
    virtualHeight = GetSystemMetrics(79) ' SM_CYVIRTUALSCREEN
    
    ' If virtual screen metrics failed, use primary screen
    If virtualWidth = 0 Or virtualHeight = 0 Then
        virtualLeft = 0
        virtualTop = 0
        virtualWidth = screenWidth
        virtualHeight = screenHeight
    End If
    
    ' Bounds checking to ensure window stays within virtual screen bounds
    Dim minMargin As Long
    minMargin = 50  ' Minimum distance from screen edges
    
    ' Ensure the window stays within the virtual screen bounds
    If centerLeft < virtualLeft + minMargin Then
        centerLeft = virtualLeft + minMargin
    ElseIf centerLeft + formWidth > virtualLeft + virtualWidth - minMargin Then
        centerLeft = virtualLeft + virtualWidth - formWidth - minMargin
    End If
    
    If centerTop < virtualTop + minMargin Then
        centerTop = virtualTop + minMargin
    ElseIf centerTop + formHeight > virtualTop + virtualHeight - minMargin Then
        centerTop = virtualTop + virtualHeight - formHeight - minMargin
    End If
    
    ' Final sanity check - if position is still invalid, fall back
    If centerLeft < virtualLeft Or centerTop < virtualTop Or _
       centerLeft + formWidth > virtualLeft + virtualWidth Or _
       centerTop + formHeight > virtualTop + virtualHeight Then
        GoTo CenterOnScreen
    End If
    
    ' Apply the calculated position
    Me.StartUpPosition = 0  ' Manual positioning
    Me.Left = centerLeft
    Me.Top = centerTop
    
    Debug.Print "ApplyOptimalPositioning: Positioned at (" & centerLeft & ", " & centerTop & ") within Excel window (" & excelLeft & ", " & excelTop & ", " & excelWidth & "x" & excelHeight & ") on virtual screen (" & virtualLeft & ", " & virtualTop & ", " & virtualWidth & "x" & virtualHeight & ")"
    Exit Sub
    
CenterOnScreen:
    ' Fallback: center on primary screen
    Debug.Print "ApplyOptimalPositioning: Using fallback - center on primary screen"
    Me.StartUpPosition = 1  ' Center on screen
End Sub

Private Sub UpdateInterface()
    ' DEBUG: Confirm this method is called
    Debug.Print "UpdateInterface called"
    
    ' Preserve current focus to prevent Close button from stealing focus
    Dim currentFocus As Control
    On Error Resume Next
    Set currentFocus = Me.ActiveControl
    On Error GoTo 0
    
    ' Set the main label to show the entire string (no truncation)
    If Len(cellData.OriginalText) > 0 Then
        lblInstructions.Caption = "Selected: " & cellData.OriginalText
        Debug.Print "  Updated lblInstructions to show full text: " & lblInstructions.Caption
    Else
        lblInstructions.Caption = "Selected: [No data]"
        Debug.Print "  No original text data available"
    End If
    
    ' Optional: Update cell count label if it exists
    On Error Resume Next
    If cellData.SelectedCellCount = 1 Then
        lblCellCount.Caption = "Processing: 1 cell"
    Else
        lblCellCount.Caption = "Processing: " & cellData.SelectedCellCount & " cells"
    End If
    Debug.Print "  Cell count: " & cellData.SelectedCellCount
    On Error GoTo 0
    
    ' Update button captions with segment previews
    UpdateButtonCaptions
    
    ' Update Undo button state based on available undo data
    UpdateUndoButtonState
    
    ' Restore focus to prevent unwanted focus changes
    On Error Resume Next
    If Not currentFocus Is Nothing Then
        ' Only restore focus if the control is still enabled and visible
        If currentFocus.Enabled And currentFocus.Visible Then
            currentFocus.SetFocus
            Debug.Print "  Restored focus to: " & currentFocus.Name
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateButtonCaptions()
    ' DEBUG: Confirm this method is called and show segment data
    Debug.Print "UpdateButtonCaptions called"
    Debug.Print "  Segment data available:"
    Debug.Print "    Segment1: '" & cellData.Segment1 & "' (length: " & Len(cellData.Segment1) & ")"
    Debug.Print "    Segment2: '" & cellData.Segment2 & "' (length: " & Len(cellData.Segment2) & ")"
    Debug.Print "    Segment3: '" & cellData.Segment3 & "' (length: " & Len(cellData.Segment3) & ")"
    Debug.Print "    ActivationID: '" & cellData.ActivationID & "' (length: " & Len(cellData.ActivationID) & ")"
    
    ' Update button captions with hybrid approach: disable buttons and grey out text for unavailable segments
    If Len(cellData.Segment1) > 0 Then 
        btn1.Enabled = True
        btn1.Caption = "1: " & Left(cellData.Segment1, 12)
        btn1.ForeColor = RGB(0, 0, 0)  ' Black text for available segments
        Debug.Print "  Updated btn1 to: " & btn1.Caption & " (enabled)"
    Else
        btn1.Enabled = False
        btn1.Caption = "1: N/A"
        btn1.ForeColor = RGB(128, 128, 128)  ' Grey text for unavailable segments
        Debug.Print "  Updated btn1 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment2) > 0 Then 
        btn2.Enabled = True
        btn2.Caption = "2: " & Left(cellData.Segment2, 12)
        btn2.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn2 to: " & btn2.Caption & " (enabled)"
    Else
        btn2.Enabled = False
        btn2.Caption = "2: N/A"
        btn2.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn2 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment3) > 0 Then 
        btn3.Enabled = True
        btn3.Caption = "3: " & Left(cellData.Segment3, 12)
        btn3.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn3 to: " & btn3.Caption & " (enabled)"
    Else
        btn3.Enabled = False
        btn3.Caption = "3: N/A"
        btn3.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn3 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment4) > 0 Then 
        btn4.Enabled = True
        btn4.Caption = "4: " & Left(cellData.Segment4, 12)
        btn4.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn4 to: " & btn4.Caption & " (enabled)"
    Else
        btn4.Enabled = False
        btn4.Caption = "4: N/A"
        btn4.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn4 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment5) > 0 Then 
        btn5.Enabled = True
        btn5.Caption = "5: " & Left(cellData.Segment5, 12)
        btn5.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn5 to: " & btn5.Caption & " (enabled)"
    Else
        btn5.Enabled = False
        btn5.Caption = "5: N/A"
        btn5.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn5 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment6) > 0 Then 
        btn6.Enabled = True
        btn6.Caption = "6: " & Left(cellData.Segment6, 12)
        btn6.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn6 to: " & btn6.Caption & " (enabled)"
    Else
        btn6.Enabled = False
        btn6.Caption = "6: N/A"
        btn6.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn6 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment7) > 0 Then 
        btn7.Enabled = True
        btn7.Caption = "7: " & Left(cellData.Segment7, 12)
        btn7.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn7 to: " & btn7.Caption & " (enabled)"
    Else
        btn7.Enabled = False
        btn7.Caption = "7: N/A"
        btn7.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn7 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment8) > 0 Then 
        btn8.Enabled = True
        btn8.Caption = "8: " & Left(cellData.Segment8, 12)
        btn8.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn8 to: " & btn8.Caption & " (enabled)"
    Else
        btn8.Enabled = False
        btn8.Caption = "8: N/A"
        btn8.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn8 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.Segment9) > 0 Then 
        btn9.Enabled = True
        btn9.Caption = "9: " & Left(cellData.Segment9, 12)
        btn9.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btn9 to: " & btn9.Caption & " (enabled)"
    Else
        btn9.Enabled = False
        btn9.Caption = "9: N/A"
        btn9.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btn9 to N/A (disabled and greyed)"
    End If
    
    If Len(cellData.ActivationID) > 0 Then 
        btnActivationID.Enabled = True
        ' Show full activation ID since they're always exactly 12 characters
        btnActivationID.Caption = "ID: " & cellData.ActivationID
        btnActivationID.ForeColor = RGB(0, 0, 0)
        Debug.Print "  Updated btnActivationID to: " & btnActivationID.Caption & " (enabled)"
    Else
        btnActivationID.Enabled = False
        btnActivationID.Caption = "ID: N/A"
        btnActivationID.ForeColor = RGB(128, 128, 128)
        Debug.Print "  Updated btnActivationID to N/A (disabled and greyed)"
    End If
    
    ' Handle TargetingAcronymCleanButton - only visible when no pipes present and caret pattern exists
    On Error Resume Next
    If InStr(cellData.OriginalText, "|") > 0 Then
        ' Data has pipes - this is taxonomy data, hide targeting button and label
        TargetingAcronymCleanButton.Visible = False
        AcronymCleanLabel.Visible = False
        Debug.Print "  Hidden TargetingAcronymCleanButton and AcronymCleanLabel (pipes present)"
    Else
        ' No pipes - check for targeting acronym pattern
        Dim targetingPattern As String
        targetingPattern = ExtractTargetingPattern(cellData.OriginalText)
        
        If Len(targetingPattern) > 0 Then
            ' Found targeting pattern - show button and label with pattern
            TargetingAcronymCleanButton.Visible = True
            TargetingAcronymCleanButton.Enabled = True
            TargetingAcronymCleanButton.Caption = "Trim: " & targetingPattern
            TargetingAcronymCleanButton.ForeColor = RGB(0, 0, 0)
            AcronymCleanLabel.Visible = True
            Debug.Print "  Shown TargetingAcronymCleanButton and AcronymCleanLabel: " & TargetingAcronymCleanButton.Caption
        Else
            ' No targeting pattern found - hide button and label
            TargetingAcronymCleanButton.Visible = False
            AcronymCleanLabel.Visible = False
            Debug.Print "  Hidden TargetingAcronymCleanButton and AcronymCleanLabel (no pattern)"
        End If
    End If
    On Error GoTo 0
    
    Debug.Print "UpdateButtonCaptions completed"
End Sub

' New method for modeless operation - called when user changes selection
Public Sub UpdateForNewSelection(target As Range)
    On Error GoTo ErrorHandler
    
    ' Only update if the selection contains valid taxonomy data
    If target.Cells.Count > 0 And Len(Trim(target.Cells(1).Value)) > 0 Then
        Dim firstCellContent As String
        firstCellContent = target.Cells(1).Value
        
        ' Check if it looks like taxonomy data (contains pipes)
        If InStr(firstCellContent, "|") > 0 Then
            ' Parse the new data
            Dim newParsedData As ParsedCellData
            newParsedData = ParseFirstCellData(firstCellContent, target.Cells.Count)
            
            ' Update our internal data
            cellData = newParsedData
            
            ' Refresh the interface
            UpdateInterface
            
            Debug.Print "UpdateForNewSelection: Updated form for new selection: " & firstCellContent
        Else
            ' Not taxonomy data, but provide feedback
            lblInstructions.Caption = "Selected: " & firstCellContent & " (no pipe-delimited data)"
            Debug.Print "UpdateForNewSelection: Selected data has no pipes, not updating buttons"
        End If
    Else
        ' Empty selection
        lblInstructions.Caption = "Selected: (empty selection)"
        Debug.Print "UpdateForNewSelection: Empty selection"
    End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "UpdateForNewSelection Error: " & Err.Description
    ' Don't show message box in modeless mode - would interrupt user workflow
End Sub

Public Sub UpdateUndoButtonState()
    ' Update Undo button appearance and caption based on multi-step undo stack
    ' Shows dynamic count of available undo operations and manages warning label
    
    Debug.Print "UpdateUndoButtonState called - UndoOperationCount: " & TaxonomyExtractorModule.UndoOperationCount
    
    ' Don't update button appearance if it's currently processing
    If Me.Tag = "UNDO_PROCESSING" Then
        Debug.Print "  Skipping button update - undo is processing"
        Exit Sub
    End If
    
    If TaxonomyExtractorModule.UndoOperationCount = 0 Then
        ' Disable button with grey appearance when no undo operations available
        btnUndo.Enabled = False
        btnUndo.ForeColor = RGB(128, 128, 128)  ' Grey text for disabled state
        btnUndo.Caption = "Undo Last"
        Debug.Print "  Undo button disabled (no operations available)"
    ElseIf TaxonomyExtractorModule.UndoOperationCount = 1 Then
        ' Enable button with normal appearance for single operation
        btnUndo.Enabled = True
        btnUndo.ForeColor = RGB(0, 0, 0)  ' Black text for enabled state
        btnUndo.BackColor = RGB(240, 240, 240)  ' Normal button background
        btnUndo.Caption = "Undo Last (1)"
        Debug.Print "  Undo button enabled (1 operation available)"
    Else
        ' Enable button with operation count for multiple operations
        btnUndo.Enabled = True
        btnUndo.ForeColor = RGB(0, 0, 0)  ' Black text for enabled state
        btnUndo.BackColor = RGB(240, 240, 240)  ' Normal button background
        btnUndo.Caption = "Undo Last (" & TaxonomyExtractorModule.UndoOperationCount & ")"
        Debug.Print "  Undo button enabled (" & TaxonomyExtractorModule.UndoOperationCount & " operations available)"
    End If
    
    ' Handle capacity warning label visibility
    On Error Resume Next  ' Handle case where label doesn't exist yet
    If TaxonomyExtractorModule.UndoOperationCount = 10 Then
        lblUndoWarning.Visible = True
        Debug.Print "  Undo warning label shown (capacity reached)"
    Else
        lblUndoWarning.Visible = False
        Debug.Print "  Undo warning label hidden"
    End If
    On Error GoTo 0
End Sub

' Helper function to restore focus to a specific button after extraction operations
Private Sub RestoreFocusToClickedButton(targetButton As Control)
    On Error Resume Next
    ' Small delay to let extraction operations complete
    DoEvents
    ' Restore focus to the button that was clicked
    If targetButton.Enabled And targetButton.Visible Then
        targetButton.SetFocus
        Debug.Print "RestoreFocusToClickedButton: Focus restored to " & targetButton.Name
    End If
    On Error GoTo 0
End Sub

Private Sub btn1_Click(): Call ExtractPipeSegment(1): Call RestoreFocusToClickedButton(btn1): End Sub
Private Sub btn2_Click(): Call ExtractPipeSegment(2): Call RestoreFocusToClickedButton(btn2): End Sub  
Private Sub btn3_Click(): Call ExtractPipeSegment(3): Call RestoreFocusToClickedButton(btn3): End Sub
Private Sub btn4_Click(): Call ExtractPipeSegment(4): Call RestoreFocusToClickedButton(btn4): End Sub
Private Sub btn5_Click(): Call ExtractPipeSegment(5): Call RestoreFocusToClickedButton(btn5): End Sub
Private Sub btn6_Click(): Call ExtractPipeSegment(6): Call RestoreFocusToClickedButton(btn6): End Sub
Private Sub btn7_Click(): Call ExtractPipeSegment(7): Call RestoreFocusToClickedButton(btn7): End Sub
Private Sub btn8_Click(): Call ExtractPipeSegment(8): Call RestoreFocusToClickedButton(btn8): End Sub
Private Sub btn9_Click(): Call ExtractPipeSegment(9): Call RestoreFocusToClickedButton(btn9): End Sub
Private Sub btnActivationID_Click(): Call ExtractActivationID: Call RestoreFocusToClickedButton(btnActivationID): End Sub
Private Sub TargetingAcronymCleanButton_Click(): Call CleanTargetingAcronyms: Call RestoreFocusToClickedButton(TargetingAcronymCleanButton): End Sub
Private Sub btnUndo_Click()
    ' Check global flag first - exit immediately if undo already in progress
    If TaxonomyExtractorModule.UndoInProgress Then
        Debug.Print "btnUndo_Click: Blocked - undo already in progress"
        Exit Sub
    End If
    
    ' Set global flag and disable events immediately
    TaxonomyExtractorModule.UndoInProgress = True
    Application.EnableEvents = False
    
    ' Immediately disable button to prevent rapid clicking
    btnUndo.Enabled = False
    btnUndo.Caption = "Processing..."
    btnUndo.ForeColor = RGB(128, 128, 128)
    Me.Tag = "UNDO_PROCESSING"
    
    ' Re-enable events for normal operation
    Application.EnableEvents = True
    
    ' Process pending events to ensure visual update
    DoEvents
    
    ' Proceed with undo operation
    Call UndoTaxonomyCleaning
End Sub
Private Sub btnClose_Click(): Unload Me: End Sub

' Cleanup when form is terminated (important for modeless operation)
Private Sub UserForm_Terminate()
    ' Cleanup application events if this was used in modeless mode
    Call CleanupModelessEvents
    Debug.Print "UserForm_Terminate: Cleaned up modeless events"
End Sub

' BENEFITS OF THIS SMART USERFORM:
' =================================
' - Clean, professional interface with 9 clearly labeled buttons + Activation ID button
' - SMART DATA PREVIEW: Shows truncated view of your actual selected data
' - DYNAMIC BUTTON CAPTIONS: Buttons show previews of what each segment contains
' - CONTEXT-AWARE: Interface adapts to show your real data content
' - SMART POSITIONING: Centers form within Excel window for optimal placement
'   • Simple, reliable positioning using Excel's window properties
'   • Always appears in the center of Excel's application window
'   • Falls back to screen center if positioning fails
' - No typing required - just click the segment you want
' - Immediate visual feedback of both data content and extraction results
' - Much faster workflow for frequent use with live previews
' - Looks and feels like a proper Excel tool
' - Built-in UNDO button to reverse the last operation
' - Custom undo functionality with instant operation (Excel's built-in Undo doesn't work with VBA changes)

' FUNCTIONALITY:
' ==============
' - Segments 1-9: Extract specific pipe-delimited segments
' - Activation ID: Extract text after colon character
' - Trim ^ABC^: Remove targeting acronyms in format ^any_characters^ with optional trailing space (only visible when pattern detected)
' - Undo Last: Multi-step undo with LIFO behavior, shows operation count (e.g., "Undo Last (3)")
' - Close: Close the dialog
'
' EXAMPLE DATA:
' =============
' For text: "FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725"
' - Segment 1 → "FY24_26"
' - Segment 3 → "Tourism WA" 
' - Segment 5 → "Always On Remarketing"
' - Segment 9 → "Conversions"
' - Activation ID → "DJTDOM060725"
'
' For targeting acronym text: "^AT^ testing string", "^ACX123^Acxiom Targeting", or "^FB_Campaign^ Facebook data"
' - Clean ^ABC^ → "testing string", "Acxiom Targeting", or "Facebook data" (removes any ^pattern^ with optional trailing space)
'
' QUICK SETUP GUIDE FOR SMART INTERFACE:
' =======================================
' 1. Insert UserForm → Name it "TaxonomyExtractorForm"
' 2. Add Label → Name: "lblInstructions" (for data preview)
' 3. Add 9 CommandButtons → Names: "btn1" through "btn9" (for segments)
' 4. Add 1 CommandButton → Name: "btnActivationID" (for activation ID)
' 5. Add 2 CommandButtons → Names: "btnUndo", "btnClose"
' 6. Copy VBA code above into UserForm module
' 7. Test with sample data: "FY24_26|Q1-4|Tourism WA|WA|Marketing:ABC123"
' 8. Label should show "Selected: FY24_26|Q1-4|Tourism WA|WA|Marketing:ABC123" (complete text)
' 9. Buttons should show "1: FY24_26", "2: Q1-4", "3: Tourism WA", etc. (12 chars each)
' 10. ID button should show "ID: ABC123" (full activation ID)
' 11. Missing segments/ID will show "N/A" with disabled state and grey text (e.g., "7: N/A" greyed out)
'
'================================================================================