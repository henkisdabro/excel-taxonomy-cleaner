# Excel Taxonomy Extractor

A professional VBA utility for extracting specific segments from pipe-delimited taxonomy data in Excel cells, with support for activation ID extraction and custom undo functionality.

## Quick Start

1. **Copy the main code**: Use `TaxonomyExtractorModule.vb` - copy this code into an Excel VBA module
2. **Create the UserForm**: Follow instructions in `TaxonomyExtractorForm.vb` to create the 9-button interface
3. **Run macro**: Execute `TaxonomyExtractor` macro (assign to ribbon button)
4. **Choose option**: Click segment button (1-9) or Activation ID button

## Example

For text: `FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725`

- **Segment 1** → `FY24_26`
- **Segment 3** → `Tourism WA` 
- **Segment 5** → `Always On Remarketing`
- **Segment 8** → `Facebook_Instagram`
- **Segment 9** → `Conversions` (text after 8th pipe, before colon)
- **Activation ID** → `DJTDOM060725` (text after colon)

## Files

- **`TaxonomyExtractorModule.vb`** - Main VBA code with all functionality
- **`TaxonomyExtractorForm.vb`** - UserForm setup instructions and button code
- **`ADDON_INSTRUCTIONS.md`** - Complete guide for creating Excel Add-in
- **`CLAUDE.md`** - Development documentation

## Interface Options

### Professional UserForm (Recommended)
- Beautiful interface with 9 segment buttons + Activation ID button
- **Smart Label Display**: Shows truncated preview of selected data (12 chars + "...")
- **Dynamic Button Captions**: Buttons show preview of each segment content
- **Context-Aware Interface**: Adapts to your selected data automatically
- Built-in "Undo Last" button for quick reversal
- Keep dialog open to experiment with different extractions
- Click buttons instead of typing numbers
- Follow setup instructions in `TaxonomyExtractorForm.vb`

### Basic InputBox (Fallback)
- Simple text input dialog
- Works immediately if UserForm not created
- Type 1-9 to select segment, or 'A' for Activation ID
- Run `UndoTaxonomyCleaning` macro to undo changes

## Key Features

### 🎯 Flexible Extraction
- **9 Segments**: Extract any of the first 9 pipe-delimited segments
- **Activation IDs**: Extract unique identifiers after colon characters
- **Batch Processing**: Works with single cells or multiple selected ranges
- **Smart Parsing**: Handles missing segments gracefully
- **Live Preview**: See segment content before extraction

### 🔄 Custom Undo System
Since Excel's built-in Undo doesn't work with VBA changes, this tool includes:
- **Automatic**: Original values stored before each extraction
- **UserForm**: Click "Undo Last" button 
- **Manual**: Run `UndoTaxonomyCleaning` macro
- **Instant**: Silent operation without confirmation dialogs
- **Smart**: Undo data cleared after each new operation

### 📦 Excel Add-in Ready
- **Professional Distribution**: Save as .xlam file for easy sharing
- **Universal Access**: Available in all Excel workbooks once installed
- **Ribbon Integration**: Add custom button for instant access
- **Follow instructions**: Complete guide in `ADDON_INSTRUCTIONS.md`

## Installation

### Quick Setup (Basic)
1. Open Excel → Alt+F11 (VBA Editor)
2. Right-click project → Insert → Module  
3. Copy code from `TaxonomyExtractorModule.vb`
4. Save as `.xlsm` file
5. Ready to use with InputBox interface!

### Professional Setup (Recommended)
1. Follow Quick Setup above
2. Create UserForm using `TaxonomyExtractorForm.vb` instructions
3. Get beautiful 9-button interface with built-in undo
4. Assign `TaxonomyExtractor` to ribbon button

### Excel Add-in (Advanced)
1. Complete Professional Setup
2. Save as Excel Add-in (.xlam) format
3. Install via File → Options → Add-ins
4. Available in ALL Excel workbooks automatically
5. Perfect for team distribution

## Usage Workflow

### With UserForm Interface
1. **Select cells** with pipe-delimited data
2. **Click ribbon button** (or run `TaxonomyExtractor`)
3. **See your data preview** - label shows truncated content, buttons show segment previews
4. **Click segment button** (1-9) or "Activation ID" - extraction happens instantly
5. **Review results** - keep dialog open for more extractions
6. **Experiment freely** - use "Undo Last" button for instant reversal
7. **Click "Close"** when finished

### With InputBox Interface
1. **Select cells** with pipe-delimited data  
2. **Run `TaxonomyExtractor`** macro
3. **Type segment number** (1-9) or 'A' for Activation ID
4. **Results applied** immediately and silently
5. **Run `UndoTaxonomyCleaning`** to reverse if needed

## Data Format Support

### Standard Format
```
segment1|segment2|segment3|segment4|segment5|segment6|segment7|segment8|segment9:activationID
```

### Real-World Example
```
FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
```

### Edge Cases Handled
- **Missing segments**: Processes available segments, skips others
- **No colons**: Segment 9 extracts to end of text
- **Empty cells**: Skipped automatically
- **Mixed data**: Each cell processed independently

## Technical Highlights

- **Robust Error Handling**: Prevents crashes during batch processing
- **Performance Optimized**: Screen updating control for smooth operation
- **Memory Efficient**: Automatic cleanup of undo data
- **Silent Operation**: No interruptions - only error messages when needed
- **Professional UI**: Looks and feels like built-in Excel tools
- **Context-Aware Display**: Interface adapts to show your actual data content

## Perfect For

- **Marketing Teams**: Extract campaign segments from taxonomy strings
- **Data Analysts**: Parse structured pipe-delimited datasets  
- **Business Users**: No programming required - just click buttons
- **IT Departments**: Deploy as add-in for organization-wide use
- **Anyone**: Working with complex delimited data structures

Ready to streamline your taxonomy data extraction! 🚀

## Support

See `CLAUDE.md` for detailed development documentation and `ADDON_INSTRUCTIONS.md` for complete add-in creation guide.