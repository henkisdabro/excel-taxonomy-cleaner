' clsAppEvents - Application Event Handler for Modeless UserForm
' Part of IPG Mediabrands Taxonomy Extractor v1.4.0
' 
' This class module handles Excel application events to enable real-time
' UserForm updates when users change their cell selection while the 
' TaxonomyExtractorForm is open in modeless mode.
'
' IMPORTANT: Copy only this code into your class module - do NOT include
' any VERSION, BEGIN, END, or Attribute lines that appear in exported files.

Public WithEvents App As Application

' Event handler for worksheet selection changes
' Fires whenever user selects different cells while modeless form is open
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Only process if our UserForm is visible and modeless
    If Not TaxonomyExtractorForm Is Nothing Then
        If TaxonomyExtractorForm.Visible Then
            ' Temporarily disable events to prevent recursive loops
            Application.EnableEvents = False
            
            ' Update the UserForm with new selection
            Call TaxonomyExtractorForm.UpdateForNewSelection(Target)
            
            ' Re-enable events
            Application.EnableEvents = True
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Ensure events are re-enabled even if error occurs
    Application.EnableEvents = True
    
    ' Log error but don't show message box (would interrupt user workflow)
    Debug.Print "clsAppEvents.App_SheetSelectionChange Error: " & Err.Description
End Sub

' Event handler for workbook activation changes
' Helps maintain proper form behavior when switching between workbooks
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error GoTo ErrorHandler
    
    ' Ensure our UserForm remains accessible when switching workbooks
    If Not TaxonomyExtractorForm Is Nothing Then
        If TaxonomyExtractorForm.Visible Then
            ' Keep form on top and responsive
            TaxonomyExtractorForm.SetFocus
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "clsAppEvents.App_WorkbookActivate Error: " & Err.Description
End Sub

' Event handler for application deactivation
' Manages form behavior when user switches to other applications
Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    On Error GoTo ErrorHandler
    
    ' Optional: Could implement logic to handle Excel losing focus
    ' For now, just ensure events remain properly configured
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "clsAppEvents.App_WindowDeactivate Error: " & Err.Description
End Sub

' Cleanup method to be called when shutting down event monitoring
Public Sub Cleanup()
    On Error Resume Next
    Set App = Nothing
End Sub