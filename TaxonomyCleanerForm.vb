'================================================================================
' EXCEL TAXONOMY CLEANER - UserForm Code
'================================================================================
' 
' This file contains the code and instructions for creating the UserForm with 8 buttons.
' The UserForm provides a much better experience than the simple InputBox interface.
'
' SETUP INSTRUCTIONS:
' ===================
' 
' STEP 1: Create the UserForm
' ---------------------------
' 1. Open Excel VBA Editor (Alt + F11)
' 2. Right-click your project → Insert → UserForm  
' 3. Name the UserForm "TaxonomyCleanerForm" (important!)
' 4. Set UserForm properties:
'    - Width: 420
'    - Height: 280
'    - Caption: "Taxonomy Cleaner - Select Segment"
'
' STEP 2: Add Controls to the UserForm
' ------------------------------------
' 
' A) Add INSTRUCTION LABEL:
'    - Control: Label
'    - Name: lblInstructions
'    - Caption: "Click a button to extract that segment from your pipe-delimited data:"
'    - Position: Left=20, Top=20
'    - Size: Width=380, Height=40
'    - TextAlign: Center (2)
'    - Font: Size=10, Bold=True
'
' B) Add SEGMENT BUTTONS (arrange in 2 rows of 4):
'    
'    ROW 1 (Top row - segments 1-4):
'    - btn1: Left=25,  Top=70,  Width=80, Height=35, Caption="Segment 1"
'    - btn2: Left=120, Top=70,  Width=80, Height=35, Caption="Segment 2"
'    - btn3: Left=215, Top=70,  Width=80, Height=35, Caption="Segment 3" 
'    - btn4: Left=310, Top=70,  Width=80, Height=35, Caption="Segment 4"
'    
'    ROW 2 (Bottom row - segments 5-8):
'    - btn5: Left=25,  Top=120, Width=80, Height=35, Caption="Segment 5"
'    - btn6: Left=120, Top=120, Width=80, Height=35, Caption="Segment 6"
'    - btn7: Left=215, Top=120, Width=80, Height=35, Caption="Segment 7"
'    - btn8: Left=310, Top=120, Width=80, Height=35, Caption="Segment 8"
'    
'    Set all buttons to: Font Size=10, Bold=True
'
' C) Add ACTION BUTTONS:
'    - btnCancel: Name="btnCancel", Caption="Cancel", Left=120, Top=180, Width=80, Height=30, Font Size=10
'    - btnUndo: Name="btnUndo", Caption="Undo Last", Left=220, Top=180, Width=80, Height=30, Font Size=10
'
' STEP 3: Add VBA Code to UserForm
' --------------------------------
' Copy and paste the code below into the UserForm module (double-click the UserForm):

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
Private Sub btnUndo_Click(): Call UndoTaxonomyCleaning: Unload Me: End Sub

' STEP 4: Test the UserForm
' -------------------------
' 1. Close the VBA Editor
' 2. Select some cells with pipe-delimited text
' 3. Run the TaxonomyCleaner macro 
' 4. You should now see the beautiful 8-button interface!
'
' TROUBLESHOOTING:
' ================
' - If you get "Object Required" error: The UserForm name must be exactly "TaxonomyCleanerForm"
' - If buttons don't work: Make sure the ExtractPipeSegment function exists in your main module
' - If layout looks wrong: Double-check the control positions and sizes above
'
' BENEFITS OF THE USERFORM:
' =========================
' - Clean, professional interface with 8 clearly labeled buttons
' - No typing required - just click the segment you want
' - Immediate visual feedback 
' - Much faster workflow for frequent use
' - Looks and feels like a proper Excel tool
' - Built-in UNDO button to reverse the last operation
' - Custom undo functionality (Excel's built-in Undo doesn't work with VBA changes)

' UNDO FUNCTIONALITY:
' ===================
' - Every extraction operation stores the original values automatically
' - Click "Undo Last" button to restore previous values
' - Can also run UndoTaxonomyCleaning macro manually
' - Undo data is cleared after each new extraction operation
' - Confirmation dialog prevents accidental undo operations
'
'================================================================================