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
' CONTROLS:
' - 9 segment buttons (btn1 through btn9)
' - 1 activation ID button (btnActivationID)
' - 3 action buttons (btnCancel, btnUndo, btnClose)
' - 1 instruction label (lblInstructions)
'
' VBA CODE FOR THE USERFORM:
' ==========================
' Copy and paste this code into the UserForm module (double-click TaxonomyCleanerForm_2):

Private Sub UserForm_Initialize()
    Me.Caption = "Taxonomy Extractor - Segment Selector"
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

' BENEFITS OF THIS USERFORM:
' ==========================
' - Clean, professional interface with 9 clearly labeled buttons + Activation ID button
' - No typing required - just click the segment you want
' - Immediate visual feedback 
' - Much faster workflow for frequent use
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
'================================================================================