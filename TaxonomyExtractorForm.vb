'================================================================================
' EXCEL TAXONOMY EXTRACTOR - UserForm Code (TaxonomyCleanerForm_2)
'================================================================================
' 
' This file contains the VBA code for the working UserForm with 9 segment buttons + Activation ID button.
' This is the code that should be placed in your UserForm named "TaxonomyCleanerForm_2"
'
' USERFORM SETUP (Already completed if your form is working):
' ==========================================================
' 
' UserForm Name: TaxonomyCleanerForm_2
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
'   - v1.2.0: [Future updates - increment as needed]
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
' 2. SEGMENT BUTTONS (9 buttons for segments 1-9):
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
' 3. ACTIVATION ID BUTTON:
'    - Control Type: CommandButton
'    - Name: btnActivationID
'    - Default Caption: "Activation ID"
'    - Note: Caption will update to show "ID: [preview]"
'    - Position: X: 12, Y: 180
'    - Size: Width: 140, Height: 30
'
' 4. ACTION BUTTONS:
'    - Control Type: CommandButton (3 buttons)
'    - Names: btnUndo, btnCancel, btnClose
'    - Captions: "Undo Last", "Cancel", "Close"
'    - Position: X: 164, Y: 180 (btnUndo), X: 238, Y: 180 (btnCancel), X: 316, Y: 180 (btnClose)
'    - Size: Width: 68, Height: 30
'
' LAYOUT SUMMARY:
' - Form dimensions: 480 x 250
' - lblInstructions shows: "Selected: [12 chars]..."
' - Segment buttons show: "1: [8 chars]", "2: [8 chars]", etc.
' - ID button shows: "ID: [6 chars]"
' - All previews update automatically when UserForm opens
'
' VBA CODE FOR THE USERFORM:
' ==========================
' After creating all controls above, copy and paste this code into the UserForm module (double-click TaxonomyCleanerForm_2):
' 
' IMPORTANT: The code below expects the exact control names listed above.
' If lblInstructions doesn't exist, the line lblInstructions.Caption will cause an error.
' Make sure all controls are created and named correctly before adding this code.

Private cellData As ParsedCellData

Public Sub SetParsedData(parsedData As ParsedCellData)
    cellData = parsedData
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Taxonomy Extractor v1.1.0 - Segment Selector"
    
    ' Set the main label to show truncated content
    lblInstructions.Caption = "Selected: " & cellData.TruncatedDisplay
    
    ' Optional: Update button captions with segment previews
    UpdateButtonCaptions
End Sub

Private Sub UpdateButtonCaptions()
    ' Update button captions to show what each segment contains
    If Len(cellData.Segment1) > 0 Then btn1.Caption = "1: " & Left(cellData.Segment1, 8)
    If Len(cellData.Segment2) > 0 Then btn2.Caption = "2: " & Left(cellData.Segment2, 8)
    If Len(cellData.Segment3) > 0 Then btn3.Caption = "3: " & Left(cellData.Segment3, 8)
    If Len(cellData.Segment4) > 0 Then btn4.Caption = "4: " & Left(cellData.Segment4, 8)
    If Len(cellData.Segment5) > 0 Then btn5.Caption = "5: " & Left(cellData.Segment5, 8)
    If Len(cellData.Segment6) > 0 Then btn6.Caption = "6: " & Left(cellData.Segment6, 8)
    If Len(cellData.Segment7) > 0 Then btn7.Caption = "7: " & Left(cellData.Segment7, 8)
    If Len(cellData.Segment8) > 0 Then btn8.Caption = "8: " & Left(cellData.Segment8, 8)
    If Len(cellData.Segment9) > 0 Then btn9.Caption = "9: " & Left(cellData.Segment9, 8)
    If Len(cellData.ActivationID) > 0 Then btnActivationID.Caption = "ID: " & Left(cellData.ActivationID, 6)
End Sub

Private Sub btn1_Click(): Call ExtractPipeSegment(1): End Sub
Private Sub btn2_Click(): Call ExtractPipeSegment(2): End Sub  
Private Sub btn3_Click(): Call ExtractPipeSegment(3): End Sub
Private Sub btn4_Click(): Call ExtractPipeSegment(4): End Sub
Private Sub btn5_Click(): Call ExtractPipeSegment(5): End Sub
Private Sub btn6_Click(): Call ExtractPipeSegment(6): End Sub
Private Sub btn7_Click(): Call ExtractPipeSegment(7): End Sub
Private Sub btn8_Click(): Call ExtractPipeSegment(8): End Sub
Private Sub btn9_Click(): Call ExtractPipeSegment(9): End Sub
Private Sub btnActivationID_Click(): Call ExtractActivationID: End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub
Private Sub btnUndo_Click(): Call UndoTaxonomyCleaning: End Sub
Private Sub btnClose_Click(): Unload Me: End Sub

' BENEFITS OF THIS SMART USERFORM:
' =================================
' - Clean, professional interface with 9 clearly labeled buttons + Activation ID button
' - SMART DATA PREVIEW: Shows truncated view of your actual selected data
' - DYNAMIC BUTTON CAPTIONS: Buttons show previews of what each segment contains
' - CONTEXT-AWARE: Interface adapts to show your real data content
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
' - Undo Last: Restore original values before extraction
' - Cancel/Close: Close the dialog
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
' QUICK SETUP GUIDE FOR SMART INTERFACE:
' =======================================
' 1. Insert UserForm → Name it "TaxonomyCleanerForm_2"
' 2. Add Label → Name: "lblInstructions" (for data preview)
' 3. Add 9 CommandButtons → Names: "btn1" through "btn9" (for segments)
' 4. Add 1 CommandButton → Name: "btnActivationID" (for activation ID)
' 5. Add 3 CommandButtons → Names: "btnUndo", "btnCancel", "btnClose"
' 6. Copy VBA code above into UserForm module
' 7. Test with sample data: "FY24_26|Q1-4|Tourism WA|WA|Marketing:ABC123"
' 8. Label should show "Selected: FY24_26|Q1-4..."
' 9. Buttons should show "1: FY24_26", "2: Q1-4", "3: Tourism", etc.
' 10. ID button should show "ID: ABC123"
'
'================================================================================