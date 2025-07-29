# Excel Taxonomy Cleaner

A VBA utility for extracting specific segments from pipe-delimited text in Excel cells.

## Quick Start

1. **Copy the main code**: Use `TaxonomyCleanerModule.vb` - copy this code into an Excel VBA module
2. **Select cells**: Choose cells containing pipe-delimited text with activation IDs
3. **Run macro**: Execute `TaxonomyCleaner` macro to extract specific segments
4. **Choose option**: Pick segment (1-9) or activation ID

## Example

For text: `FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725`

- Segment 1 → `FY24_26`
- Segment 3 → `Tourism WA` 
- Segment 5 → `Always On Remarketing`
- Segment 9 → `Conversions`
- Activation ID → `DJTDOM060725`

## Files

- **`TaxonomyCleanerModule.vb`** - Main VBA code (copy this into Excel)
- **`TaxonomyCleanerForm.vb`** - Optional UserForm for 8-button interface
- **`CLAUDE.md`** - Development documentation

## Interface Options

### Basic (InputBox)
- Simple text input dialog
- Works immediately after copying the module code
- Type 1-9 to select segment, or 'A' for Activation ID
- Run `UndoTaxonomyCleaning` macro to undo changes

### Advanced (UserForm) 
- Professional interface with 9 segment buttons + Activation ID button
- Requires following setup instructions in `TaxonomyCleanerForm.vb`
- Click buttons instead of typing numbers
- Built-in "Undo Last" button for quick reversal

## Undo Functionality

Since Excel's built-in Undo doesn't work with VBA changes, this tool includes custom undo:

- **Automatic**: Original values stored before each extraction
- **UserForm**: Click "Undo Last" button 
- **Manual**: Run `UndoTaxonomyCleaning` macro
- **Safe**: Confirmation dialog prevents accidents
- **Smart**: Undo data cleared after each new operation

## Installation

1. Open Excel → Alt+F11 (VBA Editor)
2. Right-click project → Insert → Module  
3. Copy code from `TaxonomyCleanerModule.vb`
4. Optional: Create UserForm using `TaxonomyCleanerForm.vb` instructions
5. Save as `.xlsm` file

Ready to clean your taxonomy data! 🚀