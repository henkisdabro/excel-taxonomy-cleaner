# Deployment Checklist - Excel Taxonomy Extractor v1.6.0

This checklist ensures proper deployment of the one-liner PowerShell installation system.

## 📋 Pre-Deployment Checklist

### 1. XLAM File Preparation
- [ ] XLAM file contains all VBA code from `TaxonomyExtractorModule.vb`
- [ ] UserForm created following `TaxonomyExtractorForm.vb` instructions  
- [ ] UserForm named exactly `TaxonomyExtractorForm`
- [ ] CustomUI ribbon XML embedded with IPG branding (see `RIBBON_SOLUTION.md`)
- [ ] Ribbon shows "IPG Taxonomy Extractor" button in "IPG Tools" group on Home tab
- [ ] Ribbon callback functions already included in VBA module (no manual addition needed)
- [ ] File tested on clean Excel installation
- [ ] File saved as `.xlam` format (not .xlsm)

### 2. Repository Structure
```
excel-taxonomy-cleaner/
├── install.ps1                     # ✅ Created - Main installation script
├── README.md                       # ✅ Updated - With one-liner installation
├── TaxonomyExtractorModule.vb      # ✅ Existing - Main VBA code
├── TaxonomyExtractorForm.vb        # ✅ Existing - UserForm instructions
├── RIBBON_SOLUTION.md              # ✅ Created - Complete ribbon guide
├── DEPLOYMENT_CHECKLIST.md         # ✅ This file
├── ADDON_INSTRUCTIONS.md           # ✅ Existing - Manual add-in guide
└── CLAUDE.md                       # ✅ Existing - Development docs
```

### 3. GitHub Repository Setup
- [ ] Repository exists: `henkisdabro/excel-taxonomy-cleaner`
- [ ] `install.ps1` committed to main branch
- [ ] Updated README.md committed
- [ ] All documentation files committed
- [ ] Repository is public (required for raw.githubusercontent.com access)

### 4. GitHub Release Creation
- [ ] Create new release (e.g., v1.6.0)
- [ ] Upload `TaxonomyExtractor.xlam` as release asset
- [ ] Verify asset name matches exactly: `TaxonomyExtractor.xlam`
- [ ] Release marked as "Latest release"
- [ ] Release description includes changelog

## 🚀 Deployment Steps

### Step 1: Upload XLAM to GitHub Release
```bash
# Create release and upload XLAM file
gh release create v1.6.0 TaxonomyExtractor.xlam --title "Excel Taxonomy Extractor v1.6.0" --notes "Professional VBA utility for extracting taxonomy segments with multi-step undo and enhanced UX"
```

### Step 2: Test Installation Script Locally
```powershell
# Test the installation process
irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex
```

### Step 3: Verify Download URL
Test that the GitHub Releases API returns the correct download URL:
```powershell
$release = Invoke-RestMethod "https://api.github.com/repos/henkisdabro/excel-taxonomy-cleaner/releases/latest"
$release.assets | Where-Object { $_.name -eq "TaxonomyExtractor.xlam" } | Select-Object browser_download_url
```

### Step 4: Test Complete Installation Flow
- [ ] Fresh Excel installation (or VM)
- [ ] Run PowerShell one-liner
- [ ] Verify XLAM downloads to Templates folder
- [ ] Verify file is unblocked
- [ ] Verify registry entry created
- [ ] Open Excel and check add-in loads
- [ ] Verify ribbon button appears
- [ ] Test functionality with sample data
- [ ] Test uninstall command

## 🔧 Configuration Verification

### PowerShell Script Configuration
Verify these values in `install.ps1`:
- [ ] `$RepoOwner = "henkisdabro"`
- [ ] `$RepoName = "excel-taxonomy-cleaner"`
- [ ] `$AddInName = "TaxonomyExtractor.xlam"`
- [ ] GitHub API URL correct
- [ ] Templates folder path correct: `$env:APPDATA\Microsoft\Templates`

### Registry Paths (Windows 11/Excel 2024)
- [ ] Add-in registration: `HKCU:\Software\Microsoft\Office\16.0\Excel\Options`
- [ ] Trusted locations: `HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations`

## 🧪 Testing Matrix

### Test Environments
- [ ] Windows 11 + Excel 2024 (Office 365)
- [ ] Windows 11 + Excel 2021
- [ ] Windows 10 + Excel 365
- [ ] Clean Excel installation (no existing add-ins)
- [ ] Excel with existing add-ins (conflict testing)

### Test Scenarios
- [ ] **Fresh installation**: No prior version installed
- [ ] **Upgrade scenario**: Previous version exists
- [ ] **Reinstallation**: Same version already installed
- [ ] **Network restrictions**: Corporate firewall/proxy
- [ ] **Antivirus software**: Windows Defender + third-party AV
- [ ] **PowerShell execution policy**: Default restrictive settings

### Functional Testing
- [ ] Add-in loads automatically on Excel startup
- [ ] Ribbon button appears in correct location
- [ ] Button click launches TaxonomyExtractor function
- [ ] UserForm displays with correct data preview
- [ ] All 9 segment buttons work correctly
- [ ] Activation ID button works
- [ ] Undo functionality works
- [ ] Uninstall removes all components

## 📈 Success Metrics

### Installation Success Indicators
- [ ] PowerShell script completes without errors
- [ ] XLAM file downloaded to Templates folder
- [ ] File successfully unblocked
- [ ] Registry entries created
- [ ] Desktop instructions file created
- [ ] No Windows security warnings

### Runtime Success Indicators
- [ ] Excel loads add-in automatically
- [ ] "IPG Tools" group appears on Home tab with "IPG Taxonomy Extractor" button
- [ ] Button click responds immediately
- [ ] UserForm displays actual data preview
- [ ] Segment extraction works silently
- [ ] Undo system functions properly

## 🐛 Common Issues & Solutions

### PowerShell Execution Policy
**Issue**: Script won't run due to execution policy
**Solution**: Include in documentation:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### GitHub API Rate Limiting
**Issue**: Too many requests to GitHub API
**Solution**: Script includes proper User-Agent header and error handling

### Antivirus False Positives
**Issue**: XLAM file flagged as suspicious
**Solution**: Include file hash verification and code signing information

### Excel Trust Issues
**Issue**: Add-in loads but ribbon doesn't appear
**Solution**: Script adds Templates folder to trusted locations as backup

## 📊 Monitoring & Analytics

### Track These Metrics
- [ ] PowerShell script download frequency (GitHub raw file requests)
- [ ] GitHub Release download counts
- [ ] Repository stars/forks as adoption indicator
- [ ] Issues opened related to installation problems

### Success Rate Monitoring
- [ ] Monitor GitHub Issues for installation problems
- [ ] Track success rate based on user feedback
- [ ] Document common installation failures
- [ ] Update script based on user issues

## 🔄 Update Process

### For Future Versions
1. [ ] Update version numbers in all files
2. [ ] Create new GitHub Release with updated XLAM
3. [ ] Update `install.ps1` if needed (usually not required)
4. [ ] Test installation of new version
5. [ ] Update documentation

### Hotfix Process
1. [ ] Fix issue in XLAM file
2. [ ] Create patch release (e.g., v1.6.1)
3. [ ] Upload fixed XLAM file
4. [ ] No script changes needed (auto-downloads latest)

## ✅ Final Deployment Verification

Before going live, verify:
- [ ] All checkboxes in this document completed
- [ ] Test installation successful on 3+ different systems
- [ ] Documentation is clear and complete
- [ ] One-liner command works from any PowerShell session
- [ ] Uninstall process works correctly
- [ ] No errors in PowerShell script execution
- [ ] All GitHub repository settings correct
- [ ] Release assets properly uploaded

## 🎯 Post-Deployment Tasks

### Immediate (0-24 hours)
- [ ] Monitor for initial user feedback
- [ ] Watch for GitHub Issues related to installation
- [ ] Test one-liner from multiple networks
- [ ] Document any issues that arise

### Short-term (1-7 days)
- [ ] Gather user feedback on installation experience
- [ ] Monitor download statistics
- [ ] Update documentation based on user questions
- [ ] Address any critical issues

### Long-term (1+ weeks)
- [ ] Plan next version improvements
- [ ] Consider additional distribution methods
- [ ] Evaluate need for code signing
- [ ] Document lessons learned

---

**Ready for deployment when all checkboxes are completed!** ✅