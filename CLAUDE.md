# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an advanced VBA (Visual Basic for Applications) utility for Excel that provides flexible extraction of text segments from pipe-delimited taxonomy data. The tool features a user-friendly interface allowing users to extract text before any specified pipe position across multiple selected cells.

## Architecture

### Core Functionality
- **Enhanced Range Support**: Works with single cells or multiple selected cell ranges
- **User Interface**: Custom UserForm with numbered buttons (1-8) for different extraction options
- **Flexible Text Processing**: Extracts text before the nth pipe position based on user selection
- **Batch Processing**: Processes multiple cells simultaneously with progress feedback

### Key Components
- **Main Entry Point** (`TaxonomyCleaner`): Validates selection and launches the user interface
- **Extraction Engine** (`ExtractBeforePipe`): Processes selected cells based on pipe position
- **User Interface** (`TaxonomyCleanerForm`): Provides intuitive button-based selection interface
- **Comprehensive Validation**: Checks for text content, proper selections, and pipe availability

## Development Environment

### Requirements
- Microsoft Excel with VBA support enabled
- No external dependencies or package management

### Testing the VBA Code
1. Open Microsoft Excel
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new module (`Insert > Module`)
4. Copy the main subroutines from `script.vb` into the module
5. Create a UserForm named "TaxonomyCleanerForm" following the design instructions in the script
6. Add the UserForm event handlers to the form module
7. Close the VBA Editor and test with sample pipe-delimited data

### Usage Workflow
1. Select one or more cells containing pipe-delimited text
2. Run the `TaxonomyCleaner` macro (or assign it to a button)
3. Choose from numbered buttons (1-8) to extract text before that pipe position
4. Review the processed results and completion message

## Code Structure

### Main Functions

#### `TaxonomyCleaner()`
- Entry point macro that validates cell selection
- Checks for text content in selected cells
- Launches the UserForm interface for user interaction

#### `ExtractBeforePipe(pipeNumber As Integer)`
- Core extraction logic that processes all selected cells
- Extracts text before the specified pipe position
- Provides detailed feedback on processing results
- Handles edge cases (empty cells, insufficient pipes)

#### UserForm Event Handlers
- 8 button click handlers (btn1_Click through btn8_Click)
- Each button calls `ExtractBeforePipe` with corresponding pipe number
- Cancel button for user to exit without processing

### Error Handling
- **No Selection**: Prompts user to select cells before running
- **No Text Content**: Validates that selected cells contain text
- **Insufficient Pipes**: Processes cells with available pipes, reports results
- **Processing Summary**: Shows count of successfully processed cells

## Data Format Expectations and Examples

The utility works with pipe-delimited data in this format:
```
field1|field2|field3|field4|field5|field6|field7|field8
```

**Button Examples** (for input: `A|B|C|D|E|F|G|H`):
- **Button 1**: Extracts `A` (before 1st pipe)
- **Button 2**: Extracts `A|B` (before 2nd pipe)
- **Button 3**: Extracts `A|B|C` (before 3rd pipe)
- **Button 4**: Extracts `A|B|C|D` (before 4th pipe)
- **Button 5**: Extracts `A|B|C|D|E` (before 5th pipe)
- And so on...

## Deployment

This enhanced VBA utility requires more setup than a simple macro:

### Installation Steps
1. Copy the main subroutines into an Excel VBA module
2. Create the UserForm following the detailed design instructions in `script.vb`
3. Add the UserForm event handler code to the form module  
4. Create a macro button or assign the `TaxonomyCleaner` macro to a keyboard shortcut
5. Save the workbook as `.xlsm` (macro-enabled) format to preserve all components

### Recommended Setup
- Assign `TaxonomyCleaner` to a ribbon button for easy access
- Test with sample pipe-delimited data before production use
- Consider creating a backup of data before batch processing large ranges