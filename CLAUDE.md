# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview - Version 1.6.0

Advanced VBA utility for Excel providing flexible extraction of specific segments from pipe-delimited taxonomy data with activation ID support. Features professional UI with 9 segment buttons, modeless operation, real-time updates, multi-step undo functionality, and Excel Add-in deployment.

**Version 1.6.0 introduces a comprehensive multi-step undo system supporting up to 10 sequential operations with LIFO behavior, dynamic button captions showing operation counts, enhanced focus management, and professional processing feedback during undo operations.**

## Architecture & Key Components

### Core Features
- **Enhanced Range Support**: Single cells or multiple selected ranges
- **Modeless Operation** (v1.4.0): Form stays open while Excel remains interactive
- **Real-time Updates** (v1.4.0): Button content updates automatically when selection changes
- **Smart Data Preview**: Full selected data display with dynamic button captions
- **Context-Aware Parsing**: Automatically parses first selected cell into individual segments
- **Pipe Validation** (v1.4.0): Requires pipes for segment display, shows "N/A" for single values
- **Custom Undo System**: Works where Excel's built-in undo cannot (VBA changes)
- **Professional UI**: Smart positioning, dynamic previews, silent operation

### Main Components
- **`TaxonomyExtractor`**: Modal entry point, validates selection and launches UI
- **`TaxonomyExtractorModeless`** (v1.4.0): Modeless entry with real-time updates
- **`clsAppEvents`** (v1.4.0): Monitors Excel selection changes for modeless operation
- **`ParseFirstCellData`**: Parses cells into segments and activation ID
- **`ExtractPipeSegment`** / **`ExtractActivationID`**: Core extraction logic
- **`UndoTaxonomyCleaning`**: Custom undo with automatic value storage
- **`TaxonomyExtractorForm`**: Professional 9-button interface with dynamic content

## Development Environment

### Requirements
- Microsoft Excel with VBA support enabled
- No external dependencies

### File Structure
- **TaxonomyExtractorModule.vb**: Main VBA module with core functionality
- **TaxonomyExtractorForm.vb**: UserForm code and setup instructions
- **install.ps1**: PowerShell installation script for GitHub deployment
- **RIBBON_SOLUTION.md**: CustomUI ribbon embedding guide
- **DEPLOYMENT_CHECKLIST.md**: Production deployment guide

### Testing Workflow
1. Open Excel → Alt + F11 → Insert Module
2. Copy code from `TaxonomyExtractorModule.vb`
3. Create UserForm per `TaxonomyExtractorForm.vb` instructions
4. Test with sample pipe-delimited data

### Usage Workflows

#### Modeless Interface (Recommended - v1.4.0+)
1. Select cells with pipe-delimited text → Run `TaxonomyExtractorModeless`
2. Form stays open, Excel remains interactive
3. Select different cells → Form auto-updates with new data
4. Click segment buttons → Processes currently selected cells
5. Continue selecting and extracting without reopening
6. Use "Undo Last" for instant reversal

#### Traditional Modal Interface
1. Select cells → Run `TaxonomyExtractor` → UserForm appears
2. See data preview and segment button previews
3. Click segment button → Instant extraction, no confirmations
4. Use "Undo Last" for reversal → Close when finished

## Data Format Support

### IPG Interact Taxonomy Format
Designed for taxonomy format from **IPG Interact Taxonomy tool** used across Campaign names, Insertion Order names, Ad group names, Line item names, and Ad names.

**Format**: `segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID`

**Example**: `FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725`

**Extraction Results**:
- Segment 1: `FY24_26` | Segment 3: `Tourism WA` | Segment 5: `Always On Remarketing`
- Segment 8: `Facebook_Instagram` | Segment 9: `Conversions` | Activation ID: `DJTDOM060725`

## Deployment Options

### Option 1: GitHub PowerShell Installation (Recommended)
```powershell
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex
```
- Downloads latest XLAM from GitHub Releases
- Installs to `%APPDATA%\Microsoft\AddIns`
- Configures security (trusted location + unblocking)
- Registers with Excel for auto-loading
- Works without admin rights

### Option 2: Manual VBA Setup (Development)
1. Copy `TaxonomyExtractorModule.vb` into Excel VBA module
2. Create UserForm per `TaxonomyExtractorForm.vb` instructions
3. Save as `.xlsm` format

### Option 3: Manual XLAM Installation (Legacy)
1. Complete Option 2 → Save as `.xlam`
2. Install via File → Options → Add-ins

## Technical Implementation

### UserForm Requirements
- Must be named `TaxonomyExtractorForm`
- Must have `SetParsedData(parsedData As ParsedCellData)` method
- Falls back to InputBox if UserForm missing

### Data Structure
- `ParsedCellData` type with Segment1-Segment9, ActivationID variables
- Includes original text and truncated display text

### Event Management (v1.4.0)
- `clsAppEvents.App_SheetSelectionChange()`: Monitors selection changes
- `RefreshModelessFormIfOpen()`: Updates UI after operations
- `CleanupModelessEvents()`: Prevents memory leaks

## Version Management Guidelines

### ⚠️ CRITICAL VERSION UPDATE CHECKLIST

**Every version increment MUST update ALL these locations:**

#### 🎯 PRIMARY VERSION LOCATIONS:
1. **install.ps1**:
   - Line 1: `# Excel Taxonomy Cleaner v1.X.X - One-Click Installation Script`
   - Line 14: `$AddInName = "ipg_taxonomy_extractor_addonv1.X.X.xlam"`
   - Line 15: `$DisplayName = "Excel Taxonomy Cleaner v1.X.X"`
   - Line ~208: ASCII logo header with version
   - Line ~585: Installation prompt with version

2. **TaxonomyExtractorForm.vb**: 
   - Line ~132: `Me.Caption = "IPG Mediabrands Taxonomy Extractor v1.X.X"`

3. **TaxonomyExtractorModule.vb**: 
   - Line ~673: Error message version string
   - Line ~685: Error message version string (Modeless)

4. **README.md**: Line 1: `# Excel Taxonomy Extractor v1.X.X`

5. **CLAUDE.md**: Line 5: `## Project Overview - Version 1.X.X`

6. **clsAppEvents.vb**: 
   - Line 2: `' Part of IPG Mediabrands Taxonomy Extractor v1.X.X`

7. **DEPLOYMENT_CHECKLIST.md**:
   - Line 1: Title with version
   - GitHub release command examples

#### 🚨 FAILURE CONSEQUENCES:
- **install.ps1 not updated**: Users get 404 errors, broken installation
- **VBA files not updated**: Incorrect version display, user confusion

#### 📋 VERSION UPDATE WORKFLOW:
1. Update install.ps1 filename and all version references
2. Update all VBA version strings
3. Update documentation files
4. Search repo for old version numbers: `grep -r "v1\.[0-9]\.[0-9]" .`
5. Ensure install.ps1 matches planned GitHub Release asset filename

### Version Increment Guidelines
- **Major (X.0.0)**: Breaking changes, architecture changes, UI overhauls
- **Minor (X.Y.0)**: New features, significant enhancements, performance improvements
- **Patch (X.Y.Z)**: Bug fixes, small improvements, documentation updates

## Interactive Installer Architecture (v1.5.0+)

### 🎯 Enhanced Version Requirements

**Additional v1.5.0 Interactive Interface Locations:**
- install.ps1 Line ~553: `"Initializing IPG Taxonomy Extractor AddIn Installer..."`
- install.ps1 Line ~558: `"🎯 Ready to install Excel Taxonomy Extractor AddIn v1.X.X?"`
- install.ps1 Line ~513: `"🎉 INSTALLATION COMPLETE!" "...AddIn v1.X.X is ready to use"`

### 🚨 CRITICAL INSTALLER REQUIREMENTS:

**Frame Alignment**: All UI frames 79-character width, content uses `.PadRight(77)`
**Progress System**: 9 steps = 0%→100%, one `Update-ProgressDisplay` per step
**Terminology**: Always "Excel Taxonomy Extractor AddIn" (not "Cleaner")
**Registry Management**: Auto-remove old versions, prevent duplicates

**Color Scheme**: DarkCyan (progress), Green (prompts), Cyan (logo), Yellow (status), Red (errors)
**ASCII Logo**: Never modify user-maintained logo content

### 🔧 Testing Protocol for Each Version:
1. Clean install, upgrade test, same version test
2. Frame alignment, progress flow, registry cleanup
3. Terminology consistency verification

## Development Best Practices & Lessons Learned

### Key VBA Patterns
- **Type Definitions**: `ParsedCellData` for structured data passing
- **Method Injection**: `SetParsedData` for loose coupling
- **Graceful Degradation**: InputBox fallback if UserForm missing
- **Screen Updating Control**: Performance optimization during batch operations

### UI/UX Principles Applied
- **Context Awareness**: Show user's actual data, not generic examples
- **Progressive Disclosure**: Preview before action
- **Silent Operation**: No unnecessary confirmations
- **Immediate Feedback**: Instant extraction with visible results
- **Easy Reversal**: One-click undo system

### Error Handling Strategy
- **No Selection**: Prompt user to select cells
- **No Text Content**: Validate cells contain text
- **Insufficient Segments**: Process available segments, report results
- **Loop Protection**: Ensure all selected cells get processed

### Testing Approach
- Test with real taxonomy data, not simple examples
- Verify graceful handling of missing segments
- Test batch processing with mixed data types
- Validate undo functionality across scenarios

### Future Enhancement Opportunities
1. Configuration storage for user preferences
2. Custom delimiter support
3. Direct export capabilities (CSV/other formats)
4. Batch templates for predefined patterns
5. Advanced validation and content suggestions

### Code Organization Principles
- **Single Responsibility**: Each function has one clear purpose
- **Data Encapsulation**: Type definitions for complex structures
- **Error Boundary**: Centralized error handling prevents cascades
- **Performance Awareness**: Memory management and screen updating control

## Modern Distribution Architecture

### PowerShell Integration
- GitHub Releases API integration
- Native Excel AddIns folder installation
- Automatic security configuration
- Registry integration for auto-loading

### Ribbon Button Integration
- CustomUI XML embedded via Custom UI Editor
- IPG branding: "IPG Taxonomy Extractor" in "IPG Tools" group
- Callback functions: `RibbonTaxonomyExtractor`
- Professional UI that survives distribution

### Security & Compatibility
- Trusted location setup
- File unblocking
- Optimal Excel folder compatibility
- Works without admin rights