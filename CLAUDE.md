# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an advanced VBA (Visual Basic for Applications) utility for Excel that provides flexible extraction of specific segments from pipe-delimited taxonomy data with activation ID support. The tool features a professional user interface with 9 segment buttons plus activation ID extraction, custom undo functionality, and Excel Add-in deployment capability.

## Architecture

### Core Functionality
- **Enhanced Range Support**: Works with single cells or multiple selected cell ranges
- **Professional UserForm**: Custom interface with 9 segment buttons + Activation ID button
- **Flexible Text Processing**: Extracts specific segments (1-9) or activation IDs from pipe-delimited data
- **Batch Processing**: Processes multiple cells simultaneously with progress feedback
- **Custom Undo System**: Full undo functionality since Excel's built-in undo doesn't work with VBA
- **Excel Add-in Ready**: Can be packaged as .xlam for distribution

### Key Components
- **Main Entry Point** (`TaxonomyExtractor`): Validates selection and launches the user interface
- **Segment Extraction** (`ExtractPipeSegment`): Extracts specific segments (1-9) from pipe-delimited text
- **Activation ID Extraction** (`ExtractActivationID`): Extracts text after colon characters
- **Undo System** (`UndoTaxonomyCleaning`): Custom undo functionality with automatic value storage
- **User Interface** (`TaxonomyExtractorForm`): Professional 9-button interface with undo controls
- **Comprehensive Validation**: Checks for text content, proper selections, and data format

## Development Environment

### Requirements
- Microsoft Excel with VBA support enabled
- No external dependencies or package management

### File Structure
- **TaxonomyExtractorModule.vb**: Main VBA module with core functionality and undo system
- **TaxonomyExtractorForm.vb**: UserForm code and detailed setup instructions
- **ADDON_INSTRUCTIONS.md**: Complete guide for creating and installing Excel Add-in
- **README.md**: User-friendly quick start guide

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
3. Click any segment button (1-9) or "Activation ID" - all cells process immediately and silently
4. No success dialogs or confirmations - extraction happens instantly
5. Use "Undo Last" button to reverse the last operation without confirmation
6. Use "Close" button when finished - perfect for rapid experimentation

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
- Shows TaxonomyExtractorForm (UserForm with buttons) if it exists
- Falls back to InputBox interface if UserForm not created

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

### Option 1: Basic Workbook (Quick Setup)
1. Copy code from `TaxonomyExtractorModule.vb` into an Excel VBA module
2. Create UserForm following instructions in `TaxonomyExtractorForm.vb`
3. Save workbook as `.xlsm` (macro-enabled) format
4. Assign `TaxonomyExtractor` to a ribbon button

### Option 2: Excel Add-in (Professional Distribution)
1. Follow Option 1 setup
2. Save as Excel Add-in (`.xlam`) format
3. Install via File > Options > Add-ins
4. Available in all Excel workbooks automatically
5. Follow complete instructions in `ADDON_INSTRUCTIONS.md`

### Recommended Setup
- Use Excel Add-in format for professional deployment
- Create ribbon button for easy access
- Test with sample data before production use
- Take advantage of custom undo system for safe experimentation
- UserForm provides much better user experience than InputBox fallback

### Advanced Features
- **Custom Undo System**: Works where Excel's built-in Undo cannot (VBA changes)
- **Silent Operation**: No confirmation dialogs or success messages - immediate action
- **Rapid Experimentation**: Instant extraction with one-click undo for quick testing
- **Professional Workflow**: Extract → Review → Undo → Extract again → Close (all silent)
- **Add-in Distribution**: Package as .xlam for easy sharing and installation
- **Activation ID Support**: Extract unique identifiers from colon-delimited text
- **9 Segment Support**: Handle complex taxonomy structures up to 9 segments

## Technical Notes

### UserForm Naming
- UserForm must be named exactly `TaxonomyCleanerForm_2` in the Excel VBA project
- This name is referenced in the module code for proper functionality
- If UserForm doesn't exist, code automatically falls back to InputBox interface

### Undo System Implementation
- Stores original cell values before any extraction
- Custom implementation required because Excel's Undo doesn't work with VBA changes
- Silent operation for rapid workflow without interruptions
- Undo data cleared after each new extraction operation

### Error Recovery
- Robust error handling prevents crashes during batch processing
- Screen updating control for better performance and visual feedback
- Graceful fallback from UserForm to InputBox if form doesn't exist
