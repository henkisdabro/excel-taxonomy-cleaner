# Excel Taxonomy Extractor v1.6.0

![Release](https://img.shields.io/github/v/release/henkisdabro/excel-taxonomy-cleaner)
![Language](https://img.shields.io/badge/language-VBA-blue)
![Platform](https://img.shields.io/badge/platform-Excel-green)
![License](https://img.shields.io/badge/license-MIT-yellow)

**Developer & Maintainer:** [@henkisdabro](https://github.com/henkisdabro)  
**Original Concept, Feature Contribution & User Testing:** [@stueydubs](https://github.com/stueydubs)

A professional VBA utility for extracting specific segments from pipe-delimited taxonomy data in Excel cells, featuring modeless operation, real-time updates, activation ID extraction, and custom undo functionality.

## 🎬 Demonstration

![EXCEL_AkzVTkjKyo](https://github.com/user-attachments/assets/97e7c216-2216-441d-a18e-2c86ca18c41b)

---

## Key Features

### 🎯 Flexible Extraction
- **9 Segments**: Extract any of the first 9 pipe-delimited segments
- **Activation IDs**: Extract unique identifiers after colon characters
- **Batch Processing**: Works with single cells or multiple selected ranges
- **Smart Parsing**: Handles missing segments gracefully
- **Live Preview**: See segment content before extraction
- **Enhanced UI**: Modern professional interface with smart data preview

### 🔄 Custom Undo System
Since Excel's built-in Undo doesn't work with VBA changes, this tool includes:
- **Automatic**: Original values stored before each extraction
- **UserForm**: Click "Undo Last" button 
- **Manual**: Run `UndoTaxonomyCleaning` macro
- **Instant**: Silent operation without confirmation dialogs
- **Smart**: Undo data cleared after each new operation

### 📦 Excel Add-in Ready
- **Professional Distribution**: PowerShell one-liner installation from GitHub
- **Universal Access**: Available in all Excel workbooks once installed
- **Ribbon Integration**: CustomUI XML embedded in XLAM for permanent ribbon buttons
- **Native Folder**: Installs to `%APPDATA%\Microsoft\AddIns` for optimal Excel integration
- **Follow instructions**: See `RIBBON_SOLUTION.md` for ribbon setup and `DEPLOYMENT_CHECKLIST.md` for distribution

---

## 🚀 **Fully Automated Installation**

Install instantly with this PowerShell one-liner:

```powershell
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1?$(Get-Date -UFormat %s)" | iex
```

**What this does:**
- ✅ Downloads latest XLAM version from GitHub Releases
- ✅ Installs to native Excel AddIns folder (optimal compatibility)
- ✅ Configures security settings (trusted location + unblocked file)
- ✅ Registers with Excel for automatic loading
- ✅ Works without admin rights
- ✅ Complete setup in under 30 seconds

**Interactive Installation Experience:**
- 🎨 **Awesome ASCII Logo**: IPG logo!
- 📊 **Real-time Progress Tracker**: Live progress bars showing installation steps
- ⚡ **Animated Interface**: Spinning indicators and step-by-step completion status
- 🎯 **Smart Registry Management**: Automatically removes old version registry entries
- 🔄 **Upgrade Protection**: Prevents duplicate registry keys when re-installing same version
- 🖥️ **REPL-style Interface**: Modern CLI experience with consistent frame layouts
- 🎉 **Clean Success Messages**: Minimal, focused completion information

**After installation:**
- The add-in loads automatically when Excel starts
- The **IPG Taxonomy Extractor** button appears in the **IPG Tools** group on the Home tab
- If ribbon doesn't appear, see `RIBBON_SOLUTION.md` for CustomUI XML setup
- Alternative access: File → Options → Add-ins → Excel Add-ins → Go → Browse

**To uninstall:**
Go to File → Options → Add-ins → Excel Add-ins → Go → Uncheck the add-in

---

## 🔄 Upgrading to a New Version

### Automatic Upgrade (Recommended)
```powershell
# Simply run the installer again - it handles everything
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1?$(Get-Date -UFormat %s)" | iex
```

**What the installer does automatically:**
- ✅ Downloads the latest version from GitHub
- ✅ Removes all old versions from your AddIns folder
- ✅ Cleans orphaned registry entries from previous versions
- ✅ Installs the new version with progress tracking
- ✅ Updates registry entries (prevents duplicates)
- ✅ Preserves your settings
- ✅ Shows real-time progress with animated interface

### Manual Upgrade
If you prefer manual control:

1. **Download new version** from [Releases](https://github.com/henkisdabro/excel-taxonomy-cleaner/releases/latest)
2. **In Excel**: File → Options → Add-ins → Excel Add-ins → Go
3. **Uncheck old version** (e.g., previous version)
4. **Click Browse** → Navigate to new XLAM file → OK
5. **Check the new version** → OK

**After Upgrade:**
- The new version ribbon button will appear in the IPG Tools group
- All your Excel workbooks will use the updated add-in
- Old functionality remains the same with new improvements

### Troubleshooting Upgrades

**If you see multiple versions:**
1. Go to File → Options → Add-ins → Excel Add-ins → Go
2. Uncheck ALL old versions
3. Only keep the latest version checked

**If upgrade fails:**
1. Manually delete old files from: `%APPDATA%\Microsoft\AddIns`
2. Run the PowerShell installer again
3. Restart Excel

---

## ✨ What's New in v1.6.0

### 🔄 Multi-Step Undo System
- **Professional Undo Stack** - Up to 10 sequential operations with LIFO (Last In, First Out) behavior
- **Dynamic Button Captions** - Shows operation count: "Undo Last" → "Undo Last (3)" → "Undo Last (1)"
- **Step-by-Step Reversal** - Undo operations individually in reverse order as expected
- **Capacity Warning** - Visual warning when undo limit reached (10/10)
- **Processing Feedback** - Button shows "Processing..." with disabled appearance during operations

### 🎯 Enhanced User Experience
- **Immediate Focus** - UserForm gets focus immediately when launched from ribbon
- **Smart Focus Management** - Focus returns to clicked buttons after operations, preventing unwanted focus jumps
- **Rapid Click Protection** - Prevents undo operation queuing with visual feedback during processing
- **Consistent UI Behavior** - All extraction buttons maintain focus, Close button no longer steals focus

## Previous Updates - v1.5.0

### 🎯 Smart Targeting Acronym Trimming
- **Intelligent Detection** - Smart overlay button appears automatically when targeting patterns (^ABC^) are detected in cells without pipes
- **Seamless Removal** - One-click removal of targeting text like ^AT^, ^ACX123^, ^FB_Campaign^ 
- **Full Undo Support** - Complete integration with existing undo system
- **Modeless Integration** - Works perfectly with continuous workflow operation

## Previous Updates - v1.4.0

### 🚀 Revolutionary Modeless Operation
- **Keep form open** while working with Excel - no more reopening the form for each extraction
- **Excel remains fully interactive** - click, select, and navigate normally with the form still open
- **Continuous workflow** - select different ranges and extract without interruption

### 🎯 Real-time Interface Updates
- **Automatic refresh** - buttons update instantly when you select new cells with taxonomy data
- **Live preview** - see exactly what segments are available in your currently selected data
- **Cell count display** - shows how many cells will be processed with each extraction

### 🔄 Smart Data Validation
- **Pipe validation** - buttons show "N/A" and are disabled for data without pipe delimiters
- **Post-extraction refresh** - interface immediately updates after extraction to show current state
- **Intelligent button behavior** - only show extractable segments for cleaner interface

---

## Example Usage

For text: `FY24_26|Q1-4|Tourism WA|WA |Always On Remarketing| 4LAOSO | SOC|Facebook_Instagram|Conversions:DJTDOM060725`

- **Segment 1** → `FY24_26`
- **Segment 3** → `Tourism WA` 
- **Segment 5** → `Always On Remarketing`
- **Segment 8** → `Facebook_Instagram`
- **Segment 9** → `Conversions` (text after 8th pipe, before colon)
- **Activation ID** → `DJTDOM060725` (text after colon)

---

## Interface Options

### Modeless UserForm (New in v1.4.0)
- **Continuous operation** - form stays open while Excel remains interactive
- **Real-time updates** - interface automatically refreshes when you select new cells
- **Multi-range workflow** - process different ranges without reopening the form
- **Cell count display** - shows exactly how many cells will be processed
- **Smart validation** - only shows available segments, disables unavailable ones
- Run `TaxonomyExtractorModeless` macro to launch

### Traditional Modal UserForm
- Beautiful interface with 9 segment buttons + Activation ID button
- **Smart Label Display**: Shows complete preview of selected data
- **Dynamic Button Captions**: Buttons show preview of each segment content
- **Context-Aware Interface**: Adapts to your selected data automatically
- **Smart Positioning**: Centers within Excel window while preserving your form size
- Built-in "Undo Last" button for quick reversal
- Run `TaxonomyExtractor` macro to launch

### Basic InputBox (Fallback)
- Simple text input dialog
- Works immediately if UserForm not created
- Type 1-9 to select segment, or 'A' for Activation ID
- Run `UndoTaxonomyCleaning` macro to undo changes

---

## Usage Workflow

### Modeless Interface (Continuous Operation - v1.4.0)
1. **Select cells** with pipe-delimited taxonomy data
2. **Click "IPG Taxonomy Extractor"** button in the IPG Tools group on Home tab
3. **Form stays open** - Excel remains fully interactive
4. **Select different cells** - form automatically updates to show new data
5. **Click segment button** (1-9) or "Activation ID" - processes currently selected cells
6. **Continue selecting** new ranges and extracting without reopening form
7. **Use "Undo Last"** button for instant reversal
8. **Click "Close"** when finished

### Traditional Modal Interface
1. **Select cells** with pipe-delimited data
2. **Run `TaxonomyExtractor`** macro or click ribbon button
3. **See your data preview** - label shows complete content, buttons show segment previews
4. **Click segment button** (1-9) or "Activation ID" - extraction happens instantly
5. **Review results** - use "Undo Last" button for instant reversal if needed
6. **Click "Close"** when finished

---

## Data Format Support

### IPG Interact Taxonomy Format
This tool is specifically designed to work with the taxonomy format outputted from the **IPG Interact Taxonomy tool**. This format is used consistently across:
- **Campaign names**
- **Insertion Order names** 
- **Ad group names**
- **Line item names**
- **Ad names**

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

---

## Installation Methods

### Method 1: One-Click PowerShell (Recommended)
Use the PowerShell command at the top of this README - it handles everything automatically.

### Method 2: Manual VBA Setup (Development)
1. **Copy the main code**: Use `TaxonomyExtractorModule.vb` - copy this code into an Excel VBA module
2. **Create the UserForm**: Follow instructions in `TaxonomyExtractorForm.vb` to create the interface
3. **Run macro**: Execute `TaxonomyExtractor` or `TaxonomyExtractorModeless` macro
4. **Choose option**: Click segment button (1-9) or Activation ID button

### Method 3: Manual Add-in Installation (Legacy)
1. Complete Method 2 setup above
2. Save as Excel Add-in (.xlam) format
3. Install via File → Options → Add-ins
4. Available in ALL Excel workbooks automatically
5. Perfect for team distribution

---

## Technical Highlights

- **Modeless Operation**: Revolutionary continuous workflow without interrupting Excel
- **Real-time Updates**: Interface adapts instantly to new data selections
- **Application Event Management**: Proper Excel integration with automatic cleanup
- **Robust Error Handling**: Prevents crashes during batch processing
- **Performance Optimized**: Screen updating control for smooth operation
- **Memory Efficient**: Automatic cleanup of undo data and event handlers
- **Silent Operation**: No interruptions - only error messages when needed
- **Professional UI**: Looks and feels like built-in Excel tools
- **Context-Aware Display**: Interface adapts to show your actual data content

---

## Files

- **`TaxonomyExtractorModule.vb`** - Main VBA code with all functionality and modeless operation
- **`TaxonomyExtractorForm.vb`** - UserForm setup instructions and button code  
- **`clsAppEvents.vb`** - Application event handler for modeless operation
- **`install.ps1`** - PowerShell installation script for GitHub one-liner deployment
- **`RIBBON_SOLUTION.md`** - Complete guide for embedding CustomUI ribbon buttons in XLAM
- **`DEPLOYMENT_CHECKLIST.md`** - Production deployment guide and testing procedures
- **`ADDON_INSTRUCTIONS.md`** - Manual Excel Add-in creation guide
- **`CLAUDE.md`** - Development documentation and architecture notes

---

## Version History

### v1.6.0 (Latest)
- **Multi-Step Undo System**: Professional undo stack supporting up to 10 sequential operations with LIFO behavior
- **Dynamic Button Captions**: Real-time operation count display ("Undo Last (3)") with automatic updates
- **Enhanced User Experience**: Immediate focus management, smart focus restoration, and rapid click protection
- **Processing Feedback**: Visual "Processing..." state with disabled appearance during undo operations
- **Capacity Management**: Warning label when undo limit reached with automatic oldest operation removal

### v1.5.0
- **Smart Targeting Acronym Trimming**: Intelligent overlay button for removing targeting patterns (^ABC^) from cells without pipes
- **Enhanced Detection Logic**: Automatically detects targeting patterns and enables seamless removal
- **Full Undo Integration**: Complete integration with existing undo system for acronym trimming
- **Modeless Operation Enhancement**: Perfect integration with continuous workflow operation

### v1.4.0
- **Revolutionary Modeless Operation**: Keep form open while Excel remains fully interactive
- **Real-time Interface Updates**: Buttons automatically refresh when selecting new cells
- **Smart Pipe Validation**: Buttons show "N/A" for single values without pipe delimiters
- **Post-extraction Refresh**: Interface immediately reflects extraction results
- **Selected Cell Count Display**: Shows exactly how many cells will be processed
- **Enhanced UX Workflow**: Seamless batch processing for multiple data ranges 
- **Application Event Management**: Proper Excel integration with automatic cleanup
- **Ribbon Button Enhancement**: Launches superior modeless version by default

### v1.3.0
- **Smart Positioning System**: UserForm now centers perfectly within Excel window
- **Respects Design Dimensions**: Preserves UserForm's design-time Width and Height properties
- **Enhanced Install Script**: Automatically removes old versions during upgrades
- **Improved Developer Workflow**: Comprehensive release process documentation
- **User Upgrade Instructions**: Clear upgrade path for existing users
- **Version Management**: Systematic approach to version increments and releases

### v1.2.0
- **Enhanced UserForm Interface**: Modern professional UI with smart data preview
- **Dynamic Button Captions**: Buttons show actual segment content from your data
- **Smart Label Display**: Complete preview of selected data
- **Context-Aware Parsing**: Automatically analyzes first selected cell
- **Smart Positioning**: Centers UserForm within Excel window while respecting design dimensions
- **PowerShell One-Liner Installation**: GitHub-hosted automated deployment
- **Native AddIns Folder**: Optimal Excel integration and compatibility
- **CustomUI Ribbon Support**: Embedded ribbon buttons for professional distribution
- **Improved Error Handling**: More robust validation and processing
- **Silent Operation**: No unnecessary confirmation dialogs
- **Performance Optimizations**: Better memory management and screen updating

### v1.1.0
- Added professional UserForm with 9 segment buttons
- Custom undo functionality
- Excel Add-in support

### v1.0.0
- Initial release with basic InputBox interface
- Core segment extraction functionality

---

## 🛠️ Developer Workflow - Building New Releases

This section is for developers maintaining and improving the Excel Taxonomy Extractor codebase.

### When to Increment Version

**Follow Semantic Versioning (SemVer):**
- ✅ **Major (X.0.0)**: Breaking changes that affect existing functionality
- ✅ **Minor (X.Y.0)**: New features, UI improvements, or significant enhancements
- ✅ **Patch (X.Y.Z)**: Bug fixes, small improvements, or performance optimizations

**Version Locations to Update:**
1. `TaxonomyExtractorForm.vb` - UserForm caption with new version number
2. `TaxonomyExtractorModule.vb` - Error messages with new version number
3. `install.ps1` - Script header, AddInName, DisplayName with new version
4. `README.md` - Main heading and version history section
5. `CLAUDE.md` - Project overview version references

### Step-by-Step Release Process

#### 1. **Code Development & Testing**
```bash
# Create feature branch
git checkout -b feature-name

# Make your VB code changes in:
# - TaxonomyExtractorModule.vb
# - TaxonomyExtractorForm.vb

# Update version numbers in all files listed above
# Test thoroughly in Excel VBA environment
```

#### 2. **Update Documentation**
- Update `README.md` version history with new features
- Update `CLAUDE.md` with technical changes
- Commit all code and documentation changes

#### 3. **Build XLAM Binary**
**Critical: Build the XLAM from the PREVIOUS release, not from scratch**

```bash
# Download the current release XLAM file
# Go to: https://github.com/henkisdabro/excel-taxonomy-cleaner/releases/latest
# Download the latest XLAM file (e.g., ipg_taxonomy_extractor_addonvX.Y.Z.xlam)
```

**In Excel:**
1. **Open the downloaded XLAM** from the previous release
2. **Press Alt+F11** to open VBA Editor
3. **Replace the VB code** with your updated code:
   - Copy new `TaxonomyExtractorModule.vb` content into the existing module
   - Update `TaxonomyExtractorForm` with new form code
4. **Verify the ribbon CustomUI XML** is still embedded (should be preserved)
5. **Test the functionality** thoroughly
6. **Save as new version**: `File → Save As` → `ipg_taxonomy_extractor_addonvX.Y.Z.xlam` (using semantic versioning)
7. **Close Excel**

#### 4. **Create GitHub Release**
```bash
# Push your branch and create PR
git push origin feature-name

# After merging to main:
git checkout main
git pull origin main

# Create and push tag (using semantic versioning)
git tag vX.Y.Z
git push origin vX.Y.Z
```

**On GitHub:**
1. Go to **Releases** → **Create a new release**
2. **Tag**: `vX.Y.Z` (semantic version)
3. **Title**: `Excel Taxonomy Extractor vX.Y.Z`
4. **Description**: List new features, improvements, and bug fixes
5. **Upload the XLAM file**: `ipg_taxonomy_extractor_addonvX.Y.Z.xlam`
6. **Publish release**

#### 5. **Verify Installation**
Test the PowerShell installer picks up the new version:
```powershell
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1?$(Get-Date -UFormat %s)" | iex
```

### 🎯 **Developer Checklist**
- [ ] Version numbers updated in all 5 locations (following semantic versioning)
- [ ] VB code tested in Excel environment
- [ ] Documentation updated (README.md, CLAUDE.md)
- [ ] XLAM built from previous release (preserves CustomUI)
- [ ] GitHub release created with proper semantic version tag
- [ ] XLAM binary uploaded to release with correct filename
- [ ] PowerShell installer tested with new version
- [ ] Old version cleanup verified in install script

---

## Support

See `CLAUDE.md` for detailed development documentation and `ADDON_INSTRUCTIONS.md` for complete add-in creation guide.

Ready to streamline your taxonomy data extraction! 🚀
