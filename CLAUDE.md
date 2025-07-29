# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a VBA (Visual Basic for Applications) utility for Excel that extracts campaign names from structured data containing pipe-delimited fields. The tool is designed to parse taxonomy strings and extract specific segments for data cleaning purposes.

## Architecture

### Core Functionality
- **Single File Structure**: The entire utility is contained in `script.vb`
- **Excel Integration**: Designed to run as a macro within Microsoft Excel
- **Text Processing**: Parses pipe-delimited strings to extract the 6th field (between 5th and 6th pipes)
- **In-Place Editing**: Modifies the selected cell content directly

### Key Components
- **Cell Selection Validation**: Ensures exactly one cell is selected before processing
- **Pipe Position Detection**: Dynamically finds all pipe character positions in the text
- **Text Extraction Logic**: Extracts content between specific pipe delimiters
- **Error Handling**: Provides user feedback for insufficient pipe characters or invalid selections

## Development Environment

### Requirements
- Microsoft Excel with VBA support enabled
- No external dependencies or package management

### Testing the VBA Code
1. Open Microsoft Excel
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new module (`Insert > Module`)
4. Copy the code from `script.vb` into the module
5. Close the VBA Editor
6. Test with sample data containing pipe-delimited text

### Usage Workflow
1. Select a cell containing pipe-delimited text (minimum 6 pipes required)
2. Run the `ExtractCampaignName` macro
3. The cell content will be replaced with the extracted campaign name

## Code Structure

### Main Function: `ExtractCampaignName()`
- Validates single cell selection
- Parses pipe positions using string search
- Extracts text segment between positions 5 and 6
- Replaces original cell content with extracted text

### Error Conditions
- Multiple or no cell selection: Shows "Please select exactly one cell" message
- Insufficient pipes: Shows "Need at least 6 pipe characters" with count

## Data Format Expectations

The utility expects input data in this format:
```
field1|field2|field3|field4|field5|CAMPAIGN_NAME|field7|...
```

Where `CAMPAIGN_NAME` is the target content to be extracted (6th field).

## Deployment

This is a standalone VBA script that doesn't require traditional deployment. Users need to:
1. Copy the code into their Excel VBA environment
2. Save the workbook as `.xlsm` (macro-enabled) format if they want to preserve the macro