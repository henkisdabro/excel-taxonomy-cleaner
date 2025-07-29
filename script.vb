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
    
    ' Show the user interface
    TaxonomyCleanerForm.Show
End Sub

' Function to extract text before the nth pipe for a range of cells
Sub ExtractBeforePipe(pipeNumber As Integer)
    Dim cell As Range
    Dim cellText As String
    Dim extractedText As String
    Dim pipeCount As Integer
    Dim pos As Integer
    Dim processedCount As Integer
    
    processedCount = 0
    
    For Each cell In Selection
        cellText = cell.Value
        
        ' Skip empty cells
        If Len(Trim(cellText)) = 0 Then
            GoTo NextCell
        End If
        
        ' Count pipes and find position of nth pipe
        pipeCount = 0
        pos = 1
        
        Do While pos <= Len(cellText)
            pos = InStr(pos, cellText, "|")
            If pos = 0 Then Exit Do
            pipeCount = pipeCount + 1
            If pipeCount = pipeNumber Then
                ' Extract text before the nth pipe
                extractedText = Left(cellText, pos - 1)
                cell.Value = extractedText
                processedCount = processedCount + 1
                GoTo NextCell
            End If
            pos = pos + 1
        Loop
        
        ' If we don't have enough pipes, keep original text or warn user
        If pipeCount < pipeNumber Then
            ' If requesting pipe 1 and no pipes exist, use entire text
            If pipeNumber = 1 Then
                processedCount = processedCount + 1
            End If
            ' For other cases, leave cell unchanged
        End If
        
NextCell:
    Next cell
    
    ' Show completion message
    If processedCount > 0 Then
        MsgBox "Successfully processed " & processedCount & " cell(s). Text before pipe " & pipeNumber & " has been extracted.", vbInformation, "Process Complete"
    Else
        MsgBox "No cells were processed. Make sure your selected cells contain pipe-delimited text with at least " & pipeNumber & " pipe character(s).", vbExclamation, "No Changes Made"
    End If
    
    ' Close the form
    Unload TaxonomyCleanerForm
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
    Call ExtractBeforePipe(1)
End Sub

Private Sub btn2_Click()
    Call ExtractBeforePipe(2)
End Sub

Private Sub btn3_Click()
    Call ExtractBeforePipe(3)
End Sub

Private Sub btn4_Click()
    Call ExtractBeforePipe(4)
End Sub

Private Sub btn5_Click()
    Call ExtractBeforePipe(5)
End Sub

Private Sub btn6_Click()
    Call ExtractBeforePipe(6)
End Sub

Private Sub btn7_Click()
    Call ExtractBeforePipe(7)
End Sub

Private Sub btn8_Click()
    Call ExtractBeforePipe(8)
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
