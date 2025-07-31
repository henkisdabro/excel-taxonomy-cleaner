# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview - Version 1.2.0

This is an advanced VBA (Visual Basic for Applications) utility for Excel that provides flexible extraction of specific segments from pipe-delimited taxonomy data with activation ID support. The tool features a professional user interface with 9 segment buttons plus activation ID extraction, custom undo functionality, and Excel Add-in deployment capability.

**Version 1.2.0 introduces enhanced UI capabilities with smart data preview, dynamic button captions, context-aware parsing, and modern GitHub-based PowerShell deployment for a truly professional user experience.**

## Architecture

### Core Functionality
- **Enhanced Range Support**: Works with single cells or multiple selected cell ranges
- **Professional UserForm**: Custom interface with 9 segment buttons + Activation ID button
- **Smart Data Preview**: Displays truncated view (12 chars + "...") of selected data
- **Dynamic Button Captions**: Shows preview of each segment content on buttons
- **Context-Aware Parsing**: Automatically parses first selected cell into individual segments
- **Flexible Text Processing**: Extracts specific segments (1-9) or activation IDs from pipe-delimited data
- **Batch Processing**: Processes multiple cells simultaneously with progress feedback
- **Custom Undo System**: Full undo functionality since Excel's built-in undo doesn't work with VBA
- **Excel Add-in Ready**: Can be packaged as .xlam for distribution

### Key Components
- **Main Entry Point** (`TaxonomyExtractor`): Validates selection, parses first cell, and launches the user interface
- **Data Parser** (`ParseFirstCellData`): Parses selected cell into individual segments and activation ID
- **Data Structure** (`ParsedCellData`): Type definition holding all parsed segments and display text
- **Segment Extraction** (`ExtractPipeSegment`): Extracts specific segments (1-9) from pipe-delimited text
- **Activation ID Extraction** (`ExtractActivationID`): Extracts text after colon characters
- **Undo System** (`UndoTaxonomyCleaning`): Custom undo functionality with automatic value storage
- **User Interface** (`TaxonomyExtractorForm`): Professional 9-button interface with dynamic content preview
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

#### Professional Interface (UserForm with 9 Buttons)
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
- Entry point macro that validates cell selection
- Checks for text content in selected cells
- Parses first selected cell into individual segments using `ParseFirstCellData`
- Passes parsed data to UserForm via `SetParsedData` method
- Shows TaxonomyExtractorForm (UserForm with buttons) if it exists
- Falls back to InputBox interface if UserForm not created

#### `ParseFirstCellData(cellContent As String) As ParsedCellData`
- Parses pipe-delimited text into individual segments (1-9) and activation ID
- Creates truncated display text (12 characters + "...")
- Handles missing segments gracefully with bounds checking
- Returns structured data for UserForm consumption

#### `ExtractPipeSegment(segmentNumber As Integer)`
- Core extraction logic for segments 1-9
- Handles segment extraction with colon delimiter support
- Silent operation - only shows errors if no cells processed
- Stores original values for undo functionality

#### `ExtractActivationID()`
- Specialized function for extracting activation IDs (text after colon)
- Silent operation - only shows errors if no cells processed
- Integrates with undo system

#### `UndoTaxonomyCleaning()`
- Custom undo system that works with VBA changes
- Restores original cell values before extraction
- Silent operation - no confirmation dialogs needed

#### UserForm Event Handlers
- `SetParsedData(parsedData As ParsedCellData)` - Receives parsed cell data from main module
- `UserForm_Initialize()` - Sets up interface with data preview and dynamic button captions
- `UpdateButtonCaptions()` - Updates button text to show segment content previews
- 9 segment button handlers (btn1_Click through btn9_Click)
- Activation ID button handler (btnActivationID_Click)
- Undo, Cancel, and Close button handlers

### Error Handling
- **No Selection**: Prompts user to select cells before running
- **No Text Content**: Validates that selected cells contain text
- **Insufficient Segments**: Processes cells with available segments, reports results
- **Processing Summary**: Silent operation except for error conditions
- **Undo Protection**: Immediate undo operation without confirmation
- **Loop Protection**: Error handling ensures all selected cells get processed

## Data Format Expectations and Examples

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

## Development Lessons Learned - v1.2.0

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
