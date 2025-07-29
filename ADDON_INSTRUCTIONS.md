# Excel Taxonomy Extractor Add-in Installation

## What You Get
- **Professional UserForm** with 9 segment buttons + Activation ID button
- **Ribbon button** for easy access
- **Custom undo functionality** (Excel's undo doesn't work with VBA)
- **Works in any Excel workbook** once installed

## Installation Steps

### 1. Create the Add-in File
1. **Open your working Excel file** (the one with TaxonomyExtractor)
2. **File** → **Save As**
3. **File Type**: Choose **"Excel Add-in (*.xlam)"**
4. **File Name**: `TaxonomyExtractor.xlam`
5. **Location**: Use the default Add-ins folder Excel suggests
6. **Save**

### 2. Install the Add-in
1. **File** → **Options** → **Add-ins**
2. **Manage**: Select "Excel Add-ins" → **Go**
3. **Browse** → Find your `TaxonomyExtractor.xlam` file
4. **Check the box** next to "TaxonomyExtractor" 
5. **OK**

### 3. Verify Installation
- **Open any Excel workbook**
- **Look for your ribbon button** (should appear on ribbon)
- **Test**: Select cells with pipe-delimited data and click the button

## Usage

### Example Data
```
FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725
```

### Extract Segments
1. **Select cells** containing pipe-delimited text
2. **Click your ribbon button** (or run `TaxonomyExtractor` macro)
3. **Click desired segment button**:
   - **Segment 1** → `FY24_26`
   - **Segment 3** → `Tourism WA`
   - **Segment 5** → `Always On Remarketing`
   - **Segment 9** → `Conversions`
   - **Activation ID** → `DJTDOM060725`

### Undo Changes
- **Click "Undo Last"** button in the dialog
- **Or run** `UndoTaxonomyCleaning` macro manually

## File Components

Your add-in contains:
- **`TaxonomyCleanerModule_FIXED.vb`** - Main VBA code
- **`TaxonomyCleanerForm_2`** - UserForm with buttons
- **Ribbon customization** - Your custom button

## Distribution

To share with others:
1. **Copy the `.xlam` file** to their computer
2. **They follow steps 2-3** above to install
3. **Works immediately** in all their Excel workbooks

## Troubleshooting

### Add-in not appearing
- Check **File** → **Options** → **Add-ins** → Make sure it's checked
- Try **Browse** to find the `.xlam` file again

### Button not working
- Make sure UserForm is named exactly `TaxonomyCleanerForm_2`
- Verify all button names match the VBA code

### "File not found" error
- The module code is already updated for `TaxonomyCleanerForm_2`
- Make sure the UserForm exists in the add-in

## Benefits of Add-in Format

✅ **Available in all workbooks** - no need to copy code each time  
✅ **Professional deployment** - easy to install and share  
✅ **Automatic loading** - appears whenever Excel starts  
✅ **Centralized updates** - update once, works everywhere  
✅ **Clean ribbon integration** - your button appears automatically  

🚀 **Your taxonomy extraction tool is now a professional Excel add-in!**