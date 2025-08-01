# Excel Add-in Ribbon Solution - IPG Taxonomy Extractor

This document provides the complete solution for creating the **IPG Taxonomy Extractor** ribbon button that survives XLAM distribution and appears in the IPG Tools group on Excel's Home tab.

## Problem Analysis

XLAM add-ins work perfectly with VBA macros, but ribbon buttons disappear during distribution. This is solved by embedding CustomUI XML directly into the XLAM file structure using the Office RibbonX Editor.

## ✅ **Working Solution: Office RibbonX Editor**

This is the **proven method that works** for embedding ribbon buttons in XLAM files.

### Step 1: Download Office RibbonX Editor

1. **Download** from: https://github.com/fernandreu/office-ribbonx-editor/releases
2. **Install** the Office RibbonX Editor application
3. **This is the tool that actually works** for embedding ribbon buttons in XLAM files

### Step 2: Open XLAM in Office RibbonX Editor

1. **Launch** Office RibbonX Editor
2. **Open** your `ipg_taxonomy_extractor_addonvX.Y.Z.xlam` file in the editor
3. **Insert** → Office 2010+ Custom UI Part
4. **Replace** the default XML with this IPG-branded CustomUI:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="IPGToolsGroup" 
               label="IPG Tools" 
               insertAfterMso="GroupEditingExcel">
          <button id="IPGTaxonomyExtractorButton" 
                  imageMso="FontColorPicker"
                  size="large"
                  label="IPG Taxonomy Extractor"
                  screentip="IPG Taxonomy Extractor"
                  supertip="Extract specific segments from pipe-delimited IPG taxonomy data. Features 9 segment buttons plus activation ID extraction with modeless operation."
                  onAction="TaxonomyExtractorModule.RibbonTaxonomyExtractor" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Alternative icon options with letter-based imagery:**
- `imageMso="FontColorPicker"` - "A" with color picker (recommended)
- `imageMso="FontDialog"` - Font dialog icon with "Aa"
- `imageMso="AutoSumAlpha"` - Alpha symbol
- `imageMso="TextEffectsGallery"` - Stylized "A"
- `imageMso="CharacterSpacing"` - Text formatting icon
- `imageMso="Font"` - Simple font icon

### Step 3: Validate and Save

1. **Validate** the XML using the editor's "Validate" button
2. **Ensure no errors** - any XML errors will prevent the ribbon from loading
3. **Save** the file - this embeds the CustomUI XML into your XLAM structure
4. **Close** the Office RibbonX Editor

### Step 4: Verify VBA Module Contains Ribbon Callbacks

The `TaxonomyExtractorModule.vb` file already includes the required ribbon callback functions:

- ✅ **`RibbonTaxonomyExtractor`** - Called when the IPG Taxonomy Extractor button is clicked
- ✅ **`RibbonOnLoad`** - Optional callback for ribbon initialization  
- ✅ **`myRibbon`** - Global variable to hold ribbon reference

**No additional VBA code needed** - the callbacks are already integrated in the module!

## ✅ **Result**

After completing these steps:

1. **Your XLAM file** now contains embedded CustomUI XML
2. **The ribbon button** will appear automatically when the add-in loads
3. **IPG Tools group** will be visible on Excel's Home tab
4. **IPG Taxonomy Extractor button** will launch your tool
5. **Distribution** via PowerShell installer will preserve the ribbon

## 🎯 **Final Distribution**

Once your XLAM file contains the embedded ribbon, follow the Developer instructions in README.md for details on how to deploy and distribute the file.