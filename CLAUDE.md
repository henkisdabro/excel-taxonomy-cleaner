# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview - Version 1.5.0

This is an advanced VBA (Visual Basic for Applications) utility for Excel that provides flexible extraction of specific segments from pipe-delimited taxonomy data with activation ID support. The tool features a professional user interface with 9 segment buttons plus activation ID extraction, custom undo functionality, and Excel Add-in deployment capability.

**Version 1.5.0 introduces smart targeting acronym trimming with an intelligent overlay button that appears only when targeting patterns (^ABC^) are detected in cells without pipes, enabling seamless removal of text like ^AT^, ^ACX123^, ^FB_Campaign^ with full undo support and modeless operation integration.**

## Architecture

### Core Functionality
- **Enhanced Range Support**: Works with single cells or multiple selected cell ranges
- **Professional UserForm**: Custom interface with 9 segment buttons + Activation ID button
- **Modeless Operation** (v1.4.0): Form stays open while Excel remains fully interactive
- **Real-time Updates** (v1.4.0): Button content updates automatically when selection changes
- **Smart Data Preview**: Displays full selected data with dynamic button captions
- **Selected Cell Count** (v1.4.0): Shows number of cells that will be processed
- **Dynamic Button Captions**: Shows preview of each segment content on buttons
- **Context-Aware Parsing**: Automatically parses first selected cell into individual segments
- **Pipe Validation** (v1.4.0): Requires pipe characters for segment display, shows "N/A" for single values
- **Post-Extraction Refresh** (v1.4.0): Buttons immediately update after extraction operations
- **Smart Positioning**: Centers UserForm within Excel window while respecting design-time dimensions
- **Flexible Text Processing**: Extracts specific segments (1-9) or activation IDs from pipe-delimited data
- **Batch Processing**: Processes multiple cells simultaneously with progress feedback
- **Custom Undo System**: Full undo functionality since Excel's built-in undo doesn't work with VBA
- **Excel Add-in Ready**: Can be packaged as .xlam for distribution

### Key Components
- **Main Entry Point** (`TaxonomyExtractor`): Validates selection, parses first cell, and launches the user interface (modal)
- **Modeless Entry Point** (`TaxonomyExtractorModeless`): New v1.4.0 - launches modeless form with real-time updates
- **Application Events Handler** (`clsAppEvents`): New v1.4.0 - monitors Excel selection changes for modeless operation
- **Data Parser** (`ParseFirstCellData`): Parses selected cell into individual segments and activation ID
- **Data Structure** (`ParsedCellData`): Type definition holding all parsed segments and display text
- **Segment Extraction** (`ExtractPipeSegment`): Extracts specific segments (1-9) from pipe-delimited text
- **Activation ID Extraction** (`ExtractActivationID`): Extracts text after colon characters
- **Undo System** (`UndoTaxonomyCleaning`): Custom undo functionality with automatic value storage
- **User Interface** (`TaxonomyExtractorForm`): Professional 9-button interface with dynamic content preview
  - **Real-time Updates**: New v1.4.0 - `UpdateForNewSelection` method for modeless operation
  - **Automatic Cleanup**: Proper event management and memory cleanup on form termination
- **Smart Positioning System**: Simple, reliable UserForm positioning for optimal placement
  - `ApplyOptimalPositioning`: Centers form within Excel window using Application properties
  - Respects UserForm design-time Width and Height settings
  - Falls back to screen center if positioning fails
- **Comprehensive Validation**: Checks for text content, proper selections, and data format

## Development Environment

### Requirements
- Microsoft Excel with VBA support enabled
- No external dependencies or package management

### File Structure
- **TaxonomyExtractorModule.vb**: Main VBA module with core functionality and undo system
- **TaxonomyExtractorForm.vb**: UserForm code and detailed setup instructions
- **install.ps1**: PowerShell installation script for GitHub one-liner deployment
- **RIBBON_SOLUTION.md**: Complete guide for embedding CustomUI ribbon buttons in XLAM files
- **DEPLOYMENT_CHECKLIST.md**: Production deployment guide and testing procedures
- **ADDON_INSTRUCTIONS.md**: Manual Excel Add-in creation guide (legacy approach)
- **README.md**: User-friendly quick start guide with modern installation methods

### Testing the VBA Code
1. Open Microsoft Excel
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new module (`Insert > Module`)
4. Copy code from `TaxonomyExtractorModule.vb` into the module
5. Create UserForm following instructions in `TaxonomyExtractorForm.vb`
6. Close the VBA Editor and test with sample pipe-delimited data

### Usage Workflow

#### Modeless Interface (New v1.4.0 - Continuous Operation)
1. Select one or more cells containing pipe-delimited text with activation IDs
2. Run the `TaxonomyExtractorModeless` macro - UserForm appears and stays open
3. Interface shows full preview of your data with segment button captions
4. **Excel remains interactive** - you can click and select different cells
5. Form automatically updates when you select new cells with taxonomy data
6. Click any segment button (1-9) or "Activation ID" - processes currently selected cells
7. **No need to reopen form** - continue selecting new ranges and extracting
8. Use "Undo Last" button to reverse the last operation
9. Use "Close" button when finished - perfect for batch processing multiple ranges

#### Traditional Modal Interface (UserForm with 9 Buttons)
1. Select one or more cells containing pipe-delimited text with activation IDs
2. Run the `TaxonomyExtractor` macro - UserForm appears with 9 segment buttons
3. Interface shows truncated preview of your data (12 chars + "...")
4. Button captions display preview of each segment content  
5. Click any segment button (1-9) or "Activation ID" - all cells process immediately and silently
6. No success dialogs or confirmations - extraction happens instantly
7. Use "Undo Last" button to reverse the last operation without confirmation
8. Use "Close" button when finished - perfect for rapid experimentation

#### Fallback Interface (InputBox)
1. Select one or more cells containing pipe-delimited text
2. If UserForm doesn't exist, InputBox appears automatically
3. Enter segment number (1-9) or 'A' for Activation ID
4. All cells process immediately and silently (no success messages)
5. Run `UndoTaxonomyCleaning` macro to reverse if needed

## Code Structure

### Main Functions

#### `TaxonomyExtractor()`
- Entry point macro that validates cell selection (modal operation)
- Checks for text content in selected cells
- Parses first selected cell into individual segments using `ParseFirstCellData`
- Passes parsed data to UserForm via `SetParsedData` method
- Shows TaxonomyExtractorForm (UserForm with buttons) if it exists
- Falls back to InputBox interface if UserForm not created

#### `TaxonomyExtractorModeless()` (New v1.4.0)
- Entry point macro for modeless operation - allows continuous Excel interaction
- Validates cell selection and initializes application event monitoring
- Parses first selected cell and passes data to UserForm
- Shows UserForm as modeless (`vbModeless`) and keeps Excel active
- Enables real-time updates when user changes selection
- Perfect for batch processing multiple ranges without reopening form

#### `ParseFirstCellData(cellContent As String, selectedCellCount As Long) As ParsedCellData` (Updated v1.4.0)
- Parses pipe-delimited text into individual segments (1-9) and activation ID
- Creates truncated display text (12 characters + "...")
- Includes selected cell count for UserForm display
- Validates pipe characters - requires pipes for segment parsing (v1.4.0 improvement)
- Handles missing segments gracefully with bounds checking
- Returns structured data for UserForm consumption

#### `ExtractPipeSegment(segmentNumber As Integer)` (Enhanced v1.4.0)
- Core extraction logic for segments 1-9
- Handles segment extraction with colon delimiter support
- Silent operation - only shows errors if no cells processed
- Stores original values for undo functionality
- Automatically refreshes modeless UserForm after extraction (v1.4.0)

#### `ExtractActivationID()` (Enhanced v1.4.0)
- Specialized function for extracting activation IDs (text after colon)
- Silent operation - only shows errors if no cells processed
- Integrates with undo system
- Automatically refreshes modeless UserForm after extraction (v1.4.0)

#### `UndoTaxonomyCleaning()` (Enhanced v1.4.0)
- Custom undo system that works with VBA changes
- Restores original cell values before extraction
- Silent operation - no confirmation dialogs needed
- Automatically refreshes modeless UserForm after undo (v1.4.0)

#### UserForm Event Handlers
- `SetParsedData(parsedData As ParsedCellData)` - Receives parsed cell data from main module
- `UserForm_Initialize()` - Sets up interface with data preview and dynamic button captions
- `UpdateButtonCaptions()` - Updates button text to show segment content previews
- `UpdateForNewSelection(target As Range)` - New v1.4.0 - handles real-time selection changes in modeless mode
- `UserForm_Terminate()` - New v1.4.0 - cleanup event management when form closes
- 9 segment button handlers (btn1_Click through btn9_Click)
- Activation ID button handler (btnActivationID_Click)
- Undo, Cancel, and Close button handlers

#### Application Events and Cleanup (New v1.4.0)
- `clsAppEvents.App_SheetSelectionChange()` - Monitors Excel selection changes and updates UserForm
- `RefreshModelessFormIfOpen()` - Refreshes UserForm after extraction operations
- `CleanupModelessEvents()` - Proper cleanup of application event handlers
- Prevents memory leaks and ensures stable operation across multiple form uses

### Error Handling
- **No Selection**: Prompts user to select cells before running
- **No Text Content**: Validates that selected cells contain text
- **Insufficient Segments**: Processes cells with available segments, reports results
- **Processing Summary**: Silent operation except for error conditions
- **Undo Protection**: Immediate undo operation without confirmation
- **Loop Protection**: Error handling ensures all selected cells get processed

## Data Format Expectations and Examples

### IPG Interact Taxonomy Format
The utility is specifically designed to work with the taxonomy format outputted from the **IPG Interact Taxonomy tool**. This format is used consistently across Campaign names, Insertion Order names, Ad group names, Line item names, and Ad names.

The utility works with pipe-delimited data with activation IDs in this format:
```
segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID
```

**Real Example**:
```
FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
```

**Extraction Results**:
- **Segment 1**: `FY24_26`
- **Segment 3**: `Tourism WA`
- **Segment 5**: `Always On Remarketing`
- **Segment 8**: `Facebook_Instagram`
- **Segment 9**: `Conversions` (text after 8th pipe, before colon)
- **Activation ID**: `DJTDOM060725` (text after colon)

## Deployment Options

### Option 1: GitHub PowerShell Installation (Recommended)
**User Experience:**
```powershell
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex
```

**What it does:**
1. Downloads latest XLAM from GitHub Releases automatically
2. Installs to native Excel AddIns folder (`%APPDATA%\Microsoft\AddIns`)
3. Configures security settings (trusted location + file unblocking)
4. Registers with Excel for automatic loading
5. Works without admin rights
6. Creates desktop instructions file

**Uninstall:**
Users can uninstall via Excel's built-in interface: File → Options → Add-ins → Excel Add-ins → Go → Uncheck the add-in

**Requirements:**
- XLAM file must include embedded CustomUI ribbon XML (see `RIBBON_SOLUTION.md`)
- XLAM file uploaded as GitHub Release asset
- PowerShell script hosted in repository root

### Option 2: Basic Workbook (Development/Testing)
1. Copy code from `TaxonomyExtractorModule.vb` into an Excel VBA module
2. Create UserForm following instructions in `TaxonomyExtractorForm.vb`
3. Save workbook as `.xlsm` (macro-enabled) format
4. Manual ribbon button creation or use InputBox fallback

### Option 3: Manual XLAM Installation (Legacy)
1. Follow Option 2 setup
2. Save as Excel Add-in (`.xlam`) format
3. Manually install via File > Options > Add-ins
4. Follow complete instructions in `ADDON_INSTRUCTIONS.md`

### Recommended Modern Workflow
1. **Development**: Use Option 2 for coding and testing
2. **Ribbon Integration**: Embed CustomUI XML using Custom UI Editor
3. **Distribution**: Use Option 1 for professional deployment
4. **Maintenance**: Update XLAM file in GitHub Releases, users auto-update

### Advanced Features
- **Modeless Operation** (New v1.4.0): Keep form open while interacting with Excel
- **Real-time Updates** (New v1.4.0): Form content updates automatically when selecting new cells
- **Continuous Workflow** (New v1.4.0): Process multiple ranges without reopening form
- **Application Event Management** (New v1.4.0): Proper monitoring and cleanup of Excel events
- **Custom Undo System**: Works where Excel's built-in Undo cannot (VBA changes)
- **Silent Operation**: No confirmation dialogs or success messages - immediate action
- **Rapid Experimentation**: Instant extraction with one-click undo for quick testing
- **Professional Workflow**: Extract → Review → Undo → Extract again → Close (all silent)
- **Context-Aware Interface**: Shows your actual data content in truncated form
- **Dynamic Button Previews**: See what each segment contains before extracting
- **Smart Data Parsing**: Automatically breaks down first selected cell into individual segments
- **GitHub-based Distribution**: PowerShell one-liner installation with auto-updates
- **Native AddIns Integration**: Installs to optimal Excel folder for best compatibility
- **CustomUI Ribbon Support**: Embedded ribbon buttons that survive distribution
- **Security-Aware Deployment**: Automatic trusted location setup and file unblocking
- **Activation ID Support**: Extract unique identifiers from colon-delimited text
- **9 Segment Support**: Handle complex taxonomy structures up to 9 segments

## Technical Notes

### Modern Distribution Architecture
- **PowerShell Script**: `install.ps1` handles GitHub Releases API integration
- **Native Folder**: Uses `%APPDATA%\Microsoft\AddIns` for optimal Excel compatibility
- **Security Configuration**: Automatic trusted location setup and file unblocking
- **Registry Integration**: Registers add-in for automatic Excel loading
- **GitHub Releases**: XLAM file distributed as binary release asset

### Ribbon Button Integration
- **CustomUI XML**: Must be embedded in XLAM file using Custom UI Editor
- **IPG Branding**: Button appears as "IPG Taxonomy Extractor" in "IPG Tools" group on Home tab
- **Callback Functions**: Ribbon buttons call `RibbonTaxonomyExtractor` function in the module
- **Pre-embedded**: Ribbon callbacks already included in `TaxonomyExtractorModule.vb`
- **Not Automated**: PowerShell script cannot create ribbon buttons - must be pre-embedded
- **Professional UI**: Embedded ribbon survives distribution and provides native Excel integration

### UserForm Naming
- UserForm must be named exactly `TaxonomyExtractorForm` in the Excel VBA project
- This name is referenced in the module code for proper functionality
- UserForm must have `SetParsedData` method to receive parsed cell data
- If UserForm doesn't exist, code automatically falls back to InputBox interface

### Data Structure Requirements
- `ParsedCellData` type must be defined in the main module
- Contains individual segment variables (Segment1-Segment9, ActivationID)
- Includes original text and truncated display text for UI purposes

### Undo System Implementation
- Stores original cell values before any extraction
- Custom implementation required because Excel's Undo doesn't work with VBA changes
- Silent operation for rapid workflow without interruptions
- Undo data cleared after each new extraction operation

### Error Recovery
- Robust error handling prevents crashes during batch processing
- Screen updating control for better performance and visual feedback
- Graceful fallback from UserForm to InputBox if form doesn't exist

## Version Management Guidelines

### ⚠️ CRITICAL VERSION UPDATE CHECKLIST

**Every time a version is incremented, Claude Code must update ALL these locations:**

#### 🎯 PRIMARY VERSION LOCATIONS (Must be updated simultaneously):

1. **install.ps1** - CRITICAL for downloads to work:
   - Line 1: `# Excel Taxonomy Cleaner v1.X.X - One-Click Installation Script`
   - Line 14: `$AddInName = "ipg_taxonomy_extractor_addonv1.X.X.xlam"`
   - Line 15: `$DisplayName = "Excel Taxonomy Cleaner v1.X.X"`
   - Line ~172: `Excel Taxonomy Cleaner v1.X.X - Installation Complete!`
   - Line ~204: `🎉 Excel Taxonomy Cleaner v1.X.X is now installed!`
   - Line ~237: `Excel Taxonomy Cleaner v1.X.X - Installer`

2. **TaxonomyExtractorForm.vb**:
   - Line ~101: `Me.Caption = "IPG Mediabrands Taxonomy Extractor v1.X.X"`

3. **TaxonomyExtractorModule.vb**:
   - Line ~449: `MsgBox "Error launching IPG Taxonomy Extractor: " & Err.Description, vbCritical, "IPG Taxonomy Extractor v1.X.X"`
   - Line ~461: `MsgBox "Error launching IPG Taxonomy Extractor (Modeless): " & Err.Description, vbCritical, "IPG Taxonomy Extractor v1.X.X"`

4. **README.md**:
   - Line 1: `# Excel Taxonomy Extractor v1.X.X`

5. **CLAUDE.md**:
   - Line 5: `## Project Overview - Version 1.X.X`
   - Line 9: Version introduction text

### 🚨 FAILURE CONSEQUENCES

**If install.ps1 is not updated:**
- Users cannot download the new version (404 errors)
- Installation script references wrong file names
- Completely breaks the user installation experience

**If VBA files are not updated:**
- Users see old version numbers in interface
- Error messages show incorrect version info
- Confusion about which version is actually installed

### 📋 VERSION UPDATE WORKFLOW

**When releasing a new version:**
1. **FIRST**: Update install.ps1 with new filename and all version references
2. **SECOND**: Update all VBA version strings
3. **THIRD**: Update documentation files
4. **VERIFY**: Search entire repository for old version numbers
5. **TEST**: Ensure install.ps1 references match planned release asset filename

### 🔍 SEARCH COMMANDS FOR VERIFICATION

Before releasing, run these searches to find missed version references:
```bash
grep -r "v1\.[0-9]\.[0-9]" .
grep -r "1\.[0-9]\.[0-9]" .
```

### 📦 RELEASE ASSET NAMING

**XLAM file must be named exactly as specified in install.ps1:**
- install.ps1 Line 14: `$AddInName = "ipg_taxonomy_extractor_addonv1.X.X.xlam"`
- GitHub Release asset must have this exact filename
- Any mismatch breaks the installer

### When to Increment Version Numbers

#### Major Version (X.0.0)
- **Breaking Changes**: Modifications that change existing functionality or require user action
- **Architecture Changes**: Fundamental changes to how the tool works
- **UI Overhauls**: Complete interface redesigns

#### Minor Version (X.Y.0) 
- **New Features**: Adding new functionality like modeless operation (v1.3.0 → v1.4.0)
- **Significant Enhancements**: Major improvements to existing features
- **Performance Improvements**: Notable speed or efficiency gains
- **UI Improvements**: Enhanced user interface elements

#### Patch Version (X.Y.Z)
- **Bug Fixes**: Corrections to existing functionality  
- **Small Improvements**: Minor tweaks and refinements
- **Documentation Updates**: Significant documentation improvements

### Automation Reminder

Claude Code should proactively suggest version increments when:
- Implementing new features (like the modeless system)
- Making significant improvements 
- Fixing important bugs
- Completing feature branches

Always provide a comprehensive commit message that explains the changes and increment reasoning.

## Development Lessons Learned - v1.3.0

### Key Improvements Made
1. **Smart Data Preview System**: 
   - Implemented truncated display (12 chars + "...") for better UX
   - Shows actual user data rather than generic text
   - Helps users understand what they're working with

2. **Dynamic Button Interface**:
   - Button captions now show actual segment content from user's data
   - Real-time preview of what each extraction will produce
   - Context-aware interface that adapts to user's specific data

3. **Enhanced Data Parsing**:
   - Added `ParseFirstCellData` function for smart analysis
   - Automatic breakdown of first selected cell into individual segments
   - Structured data passing between module and UserForm

4. **Professional UI Polish**:
   - Removed unnecessary confirmation dialogs
   - Silent operation for smooth workflow
   - Better integration between main module and UserForm

### Technical Architecture Insights

#### Data Flow Pattern
```
User Selection → ParseFirstCellData → SetParsedData → UserForm Display → Button Click → Extraction
```

#### Key VBA Patterns Used
- **Type Definitions**: `ParsedCellData` for structured data passing
- **Method Injection**: `SetParsedData` method on UserForm for loose coupling
- **Graceful Degradation**: Automatic fallback to InputBox if UserForm missing
- **Screen Updating Control**: Performance optimization during batch operations

#### UI/UX Best Practices Applied
- **Context Awareness**: Show user's actual data, not generic examples
- **Progressive Disclosure**: Preview before action
- **Silent Operation**: No unnecessary confirmations
- **Immediate Feedback**: Instant extraction with visible results
- **Easy Reversal**: One-click undo system

### Development Workflow Insights

#### Version Evolution
- **v1.0**: Basic functionality with InputBox
- **v1.1**: Added UserForm with static buttons
- **v1.2**: Enhanced with dynamic content, smart preview, and GitHub PowerShell distribution

#### Testing Approach
- Always test with real taxonomy data, not just simple examples
- Verify graceful handling of missing segments
- Test batch processing with mixed data types
- Validate undo functionality across different scenarios

#### Code Organization Principles
- **Single Responsibility**: Each function has one clear purpose
- **Data Encapsulation**: Use Type definitions for complex data structures
- **Error Boundary**: Centralized error handling prevents cascading failures
- **Performance Awareness**: Screen updating control and memory management

### Future Enhancement Opportunities
1. **Configuration Storage**: Save user preferences between sessions
2. **Custom Delimiters**: Support for different separator characters
3. **Export Capabilities**: Direct export to CSV or other formats
4. **Batch Templates**: Predefined extraction patterns
5. **Advanced Validation**: Content format checking and suggestions

### Best Practices for VBA Development
1. **Always provide fallback interfaces** (InputBox when UserForm fails)
2. **Use Type definitions** for complex data structures
3. **Implement custom undo** when Excel's built-in won't work
4. **Show real user data** in previews, not generic examples
5. **Make operations silent** - avoid confirmation dialogs
6. **Test with edge cases** - missing segments, empty cells, etc.
7. **Optimize for batch processing** - screen updating control essential

## Interactive Installer Architecture (v1.5.0+)

### 🎯 Enhanced Version Management Requirements

**With the new interactive installer interface, version updates require additional attention:**

#### 📍 CRITICAL VERSION LOCATIONS (Enhanced List):
1. **Core Files** (Original):
   - install.ps1 Line 1: `# Excel Taxonomy Extractor v1.X.X - One-Click Installation Script`
   - install.ps1 Line 14: `$AddInName = "ipg_taxonomy_extractor_addonv1.X.X.xlam"`
   - install.ps1 Line 15: `$DisplayName = "Excel Taxonomy Extractor AddIn v1.X.X"`
   - TaxonomyExtractorForm.vb Line ~101: `Me.Caption = "IPG Mediabrands Taxonomy Extractor v1.X.X"`
   - TaxonomyExtractorModule.vb error messages with version numbers
   - README.md Line 1: `# Excel Taxonomy Extractor v1.X.X`
   - CLAUDE.md Line 5: `## Project Overview - Version 1.X.X`

2. **Interactive Interface** (New in v1.5.0):
   - install.ps1 Line ~553: `"Initializing IPG Taxonomy Extractor AddIn Installer..."`
   - install.ps1 Line ~558: `"🎯 Ready to install Excel Taxonomy Extractor AddIn v1.X.X?"`
   - install.ps1 Line ~513: `"🎉 INSTALLATION COMPLETE!" "...AddIn v1.X.X is ready to use"`

#### 🚨 INTERACTIVE INSTALLER CRITICAL REQUIREMENTS:

**Frame Alignment Standards:**
- All UI frames must use exactly 79-character width: `╔═══...═══╗` (79 chars)
- Content within frames must use `.PadRight(77)` for proper alignment
- Progress bar component already handles proper padding
- Never use hardcoded spaces for alignment

**Progress System Integrity:**
- 9 installation steps = 0%, 11%, 22%, 33%, 44%, 55%, 66%, 77%, 88%, 100%
- Only one `Update-ProgressDisplay` call per step (via `Update-StepStatus`)
- No duplicate or manual progress display calls

**Terminology Consistency:**
- Always "Excel Taxonomy Extractor AddIn" (not "Cleaner")
- Always "AddIn" (not "add-in" or "Add-in")
- Consistent capitalization throughout all user-facing text

**Registry Management:**
- Automatically remove orphaned registry keys from previous versions
- Prevent duplicate registry entries when re-installing same version
- Use filename-only format with quotes: `"ipg_taxonomy_extractor_addonv1.X.X.xlam"`

#### 🎨 VISUAL CONSISTENCY STANDARDS:

**Color Scheme:**
- DarkCyan: Progress frames and headers
- Green: Interactive prompts and success messages
- Cyan: ASCII logo display
- Yellow: Status updates and warnings
- Red: Error messages

**ASCII Logo Protection:**
- Logo content is user-maintained - never modify the ASCII art
- Only frame the logo, never change the characters or layout
- Preserve exact spacing and positioning

#### ⚠️ VERSION UPDATE TESTING PROTOCOL:

**Required Tests for Each Version:**
1. **Clean Install Test**: Fresh installation from PowerShell one-liner
2. **Upgrade Test**: Install previous version, then upgrade to new version
3. **Same Version Test**: Install same version twice (should prevent duplicates)
4. **Frame Alignment**: Verify all UI frames display correctly
5. **Progress Flow**: Confirm percentages progress smoothly 0→100%
6. **Registry Cleanup**: Check old entries removed, new entry created
7. **Terminology Check**: All text uses "AddIn" and "Extractor" consistently

**Visual Verification:**
- [ ] ASCII logo displays correctly without corruption
- [ ] All frames have consistent borders and alignment
- [ ] Progress bars fit properly within frames
- [ ] No text overflow or truncation in UI elements
- [ ] Color scheme matches specification

**Functional Verification:**
- [ ] Registry entries for old versions are removed
- [ ] New registry entry created with proper format
- [ ] No duplicate entries when re-installing same version
- [ ] Installation progress shows smooth 0-100% progression
- [ ] All version references are consistent across installer

### 🔧 Development Workflow for Interactive Installer

When modifying the installer interface:

1. **Test Frame Alignment**: Use consistent `.PadRight(77)` for all content
2. **Verify Progress Logic**: Ensure no duplicate `Update-ProgressDisplay` calls
3. **Check Terminology**: Search for "Cleaner" or "add-in" and replace consistently
4. **Validate Colors**: Maintain DarkCyan/Green/Cyan color scheme
5. **Protect ASCII Logo**: Never modify user-maintained logo content
6. **Test All Scenarios**: Clean install, upgrade, same version re-install

This interactive installer represents a significant UX improvement and requires careful attention to these enhanced version management practices.
