'================================================================================
' EXCEL TAXONOMY CLEANER - UserForm Code
'================================================================================
' 
' This file contains the code and instructions for creating the UserForm with 9 segment buttons + Activation ID button.
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
'    - Width: 480
'    - Height: 320
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
'    - Size: Width=440, Height=40
'    - TextAlign: Center (2)
'    - Font: Size=10, Bold=True
'
' B) Add SEGMENT BUTTONS (arrange in 3 rows):
'    
'    ROW 1 (segments 1-3):
'    - btn1: Left=25,  Top=70,  Width=80, Height=35, Caption="Segment 1"
'    - btn2: Left=120, Top=70,  Width=80, Height=35, Caption="Segment 2"
'    - btn3: Left=215, Top=70,  Width=80, Height=35, Caption="Segment 3" 
'    - btn4: Left=310, Top=70,  Width=80, Height=35, Caption="Segment 4"
'    
'    ROW 2 (segments 4-6):
'    - btn5: Left=25,  Top=110, Width=80, Height=35, Caption="Segment 5"
'    - btn6: Left=120, Top=110, Width=80, Height=35, Caption="Segment 6"
'    - btn7: Left=215, Top=110, Width=80, Height=35, Caption="Segment 7"
'    - btn8: Left=310, Top=110, Width=80, Height=35, Caption="Segment 8"
'    
'    ROW 3 (segment 9 + Activation ID):
'    - btn9: Left=25,  Top=150, Width=80, Height=35, Caption="Segment 9"
'    - btnActivationID: Left=120, Top=150, Width=120, Height=35, Caption="Activation ID"
'    
'    Set all buttons to: Font Size=10, Bold=True
'
' C) Add ACTION BUTTONS:
'    - btnCancel: Name="btnCancel", Caption="Cancel", Left=100, Top=200, Width=60, Height=30, Font Size=10
'    - btnUndo: Name="btnUndo", Caption="Undo Last", Left=170, Top=200, Width=80, Height=30, Font Size=10
'    - btnClose: Name="btnClose", Caption="Close", Left=260, Top=200, Width=60, Height=30, Font Size=10
'
' STEP 3: Add VBA Code to UserForm
' --------------------------------
' Copy and paste the code below into the UserForm module (double-click the UserForm):

Private Sub UserForm_Initialize()
    Me.Caption = "Taxonomy Cleaner - Segment Selector"
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