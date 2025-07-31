# Excel Taxonomy Cleaner v1.2.0 - One-Click Installation Script
# Repository: https://github.com/henkisdabro/excel-taxonomy-cleaner
# Usage: 
#   Install: irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex
#   Uninstall: $env:TAXONOMY_UNINSTALL="true"; irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex

[CmdletBinding()]
param(
    [switch]$Uninstall,
    [string]$Version = "latest"
)

# Configuration
$RepoOwner = "henkisdabro"
$RepoName = "excel-taxonomy-cleaner"
$AddInName = "ipg_taxonomy_extractor_addonv1.2.0.xlam"
$DisplayName = "Excel Taxonomy Cleaner v1.2.0"

# Paths
$AddInsPath = "$env:APPDATA\Microsoft\AddIns"
$AddInPath = Join-Path $AddInsPath $AddInName
$TempPath = Join-Path $env:TEMP "taxonomy-extractor-install"

function Write-Status {
    param([string]$Message, [string]$Color = "Green")
    Write-Host "→ $Message" -ForegroundColor $Color
}

function Write-Error {
    param([string]$Message)
    Write-Host "✗ ERROR: $Message" -ForegroundColor Red
}

function Write-Success {
    param([string]$Message)
    Write-Host "✓ $Message" -ForegroundColor Green
}

function Test-ExcelInstalled {
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Get-LatestReleaseUrl {
    try {
        Write-Status "Fetching latest release information..."
        $apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
        $release = Invoke-RestMethod -Uri $apiUrl -Headers @{"User-Agent" = "PowerShell-ExcelTaxonomyInstaller"}
        
        $asset = $release.assets | Where-Object { $_.name -eq $AddInName }
        if (-not $asset) {
            throw "Add-in file '$AddInName' not found in latest release"
        }
        
        return @{
            DownloadUrl = $asset.browser_download_url
            Version = $release.tag_name
            ReleaseDate = $release.published_at
        }
    }
    catch {
        Write-Error "Failed to fetch release information: $($_.Exception.Message)"
        throw
    }
}

function Install-AddIn {
    try {
        # Verify Excel is installed
        if (-not (Test-ExcelInstalled)) {
            throw "Microsoft Excel is not installed or cannot be accessed"
        }

        # Get latest release
        $releaseInfo = Get-LatestReleaseUrl
        Write-Status "Found version: $($releaseInfo.Version)"

        # Create temp directory
        if (Test-Path $TempPath) {
            Remove-Item $TempPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null

        # Download add-in
        $tempAddInPath = Join-Path $TempPath $AddInName
        Write-Status "Downloading $DisplayName..."
        Invoke-WebRequest -Uri $releaseInfo.DownloadUrl -OutFile $tempAddInPath -UseBasicParsing

        # Verify download
        if (-not (Test-Path $tempAddInPath)) {
            throw "Failed to download add-in file"
        }

        # Ensure AddIns directory exists
        if (-not (Test-Path $AddInsPath)) {
            New-Item -ItemType Directory -Path $AddInsPath -Force | Out-Null
        }

        # Install to AddIns folder (native Excel add-in location)
        Write-Status "Installing to native Excel AddIns folder..."
        Copy-Item $tempAddInPath -Destination $AddInPath -Force

        # Unblock file to prevent security warnings
        Write-Status "Configuring security permissions..."
        Unblock-File -Path $AddInPath -ErrorAction SilentlyContinue

        # Register add-in in Excel registry
        Write-Status "Registering add-in with Excel..."
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        
        # Find next available OPEN slot
        $openNumber = ""
        $counter = 0
        do {
            $keyName = if ($counter -eq 0) { "OPEN" } else { "OPEN$counter" }
            $existingValue = Get-ItemProperty -Path $regPath -Name $keyName -ErrorAction SilentlyContinue
            if (-not $existingValue) {
                $openNumber = $keyName
                break
            }
            $counter++
        } while ($counter -lt 50)

        if ($openNumber) {
            New-ItemProperty -Path $regPath -Name $openNumber -Value $AddInPath -PropertyType String -Force | Out-Null
        }

        # Add AddIns folder to trusted locations (defense in depth)
        Write-Status "Ensuring AddIns folder is trusted..."
        $trustRegPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
        $locationPath = "$trustRegPath\Location99"
        
        if (-not (Test-Path $locationPath)) {
            New-Item -Path $locationPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $locationPath -Name "Path" -Value "$AddInsPath\"
        Set-ItemProperty -Path $locationPath -Name "AllowSubFolders" -Value 1
        Set-ItemProperty -Path $locationPath -Name "Description" -Value "Excel AddIns (Auto-trusted by Taxonomy Extractor installer)"

        # Create desktop shortcut with instructions
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Excel Taxonomy Cleaner - Instructions.txt"
        
        $instructions = @"
Excel Taxonomy Cleaner v1.2.0 - Installation Complete!

✓ Add-in installed successfully to: $AddInPath
✓ Registered with Excel for automatic loading
✓ Security configured (trusted location + unblocked)

HOW TO USE:
1. Open Microsoft Excel
2. Go to File → Options → Add-ins
3. At the bottom, select "Excel Add-ins" and click "Go..."
4. Browse and select: $AddInPath
5. Click OK - the add-in will load and ribbon button will appear

Or simply restart Excel - the add-in should load automatically from the native AddIns folder!

The add-in provides a professional interface for extracting segments from pipe-delimited taxonomy data.
The IPG Taxonomy Extractor button will appear in the IPG Tools group on Excel's Home tab.

Support: https://github.com/$RepoOwner/$RepoName
Version: $($releaseInfo.Version)
Installed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

To uninstall: Run the same PowerShell command with -Uninstall flag
"@
        
        $instructions | Out-File -FilePath $shortcutPath -Encoding UTF8

        # Cleanup
        Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue

        Write-Success "Installation completed successfully!"
        Write-Host ""
        Write-Host "🎉 Excel Taxonomy Cleaner v1.2.0 is now installed!" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor Yellow
        Write-Host "1. Open Microsoft Excel" -ForegroundColor White
        Write-Host "2. Go to File → Options → Add-ins → Excel Add-ins → Go → Browse" -ForegroundColor White
        Write-Host "3. Select: $AddInPath" -ForegroundColor Gray
        Write-Host "4. The add-in will load with its ribbon button" -ForegroundColor White
        Write-Host ""
        Write-Host "📄 Full instructions saved to desktop: Excel Taxonomy Cleaner - Instructions.txt" -ForegroundColor Green
        Write-Host "📂 Add-in location: $AddInPath" -ForegroundColor Gray
        Write-Host "🏠 Native Excel AddIns folder used for optimal compatibility" -ForegroundColor Gray
        Write-Host "🎯 IPG Taxonomy Extractor button will appear in IPG Tools group on Home tab" -ForegroundColor Gray
        Write-Host ""
        Write-Host "To uninstall later:" -ForegroundColor Yellow
        Write-Host "`$env:TAXONOMY_UNINSTALL=`"true`"; irm `"https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1`" | iex" -ForegroundColor Gray
        Write-Host ""

    }
    catch {
        Write-Error $_.Exception.Message
        Write-Host ""
        Write-Host "Installation failed. Please try again or check:" -ForegroundColor Yellow
        Write-Host "• Internet connection" -ForegroundColor White
        Write-Host "• Excel installation" -ForegroundColor White
        Write-Host "• PowerShell execution policy (try: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser)" -ForegroundColor White
        exit 1
    }
}

function Uninstall-AddIn {
    try {
        Write-Status "Uninstalling Excel Taxonomy Cleaner..."

        # Remove add-in file
        if (Test-Path $AddInPath) {
            Remove-Item $AddInPath -Force
            Write-Success "Removed add-in file"
        }

        # Remove registry entries
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        $counter = 0
        do {
            $keyName = if ($counter -eq 0) { "OPEN" } else { "OPEN$counter" }
            $value = Get-ItemProperty -Path $regPath -Name $keyName -ErrorAction SilentlyContinue
            if ($value -and $value.$keyName -eq $AddInPath) {
                Remove-ItemProperty -Path $regPath -Name $keyName -Force
                Write-Success "Removed registry entry: $keyName"
                break
            }
            $counter++
        } while ($counter -lt 50)

        # Clean up desktop instructions
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Excel Taxonomy Cleaner - Instructions.txt"
        if (Test-Path $shortcutPath) {
            Remove-Item $shortcutPath -Force
        }

        Write-Success "Uninstallation completed successfully!"
        Write-Host ""
        Write-Host "Excel Taxonomy Cleaner has been removed from your system." -ForegroundColor Green
        Write-Host "You may need to restart Excel to complete the removal." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Uninstallation failed: $($_.Exception.Message)"
        exit 1
    }
}

# Check for uninstall environment variable (for one-liner compatibility)
if ($env:TAXONOMY_UNINSTALL -eq "true") {
    $Uninstall = $true
    Write-Host "DEBUG: Environment variable detected, setting Uninstall=true" -ForegroundColor Yellow
}

# Main execution
try {
    Write-Host ""
    if ($Uninstall) {
        Write-Host "Excel Taxonomy Cleaner v1.2.0 - Uninstaller" -ForegroundColor Red
    } else {
        Write-Host "Excel Taxonomy Cleaner v1.2.0 - Installer" -ForegroundColor Cyan
    }
    Write-Host "Repository: https://github.com/$RepoOwner/$RepoName" -ForegroundColor Gray
    Write-Host ""

    if ($Uninstall) {
        Uninstall-AddIn
    } else {
        Install-AddIn
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
finally {
    # Clean up environment variable
    if ($env:TAXONOMY_UNINSTALL) {
        Remove-Item Env:TAXONOMY_UNINSTALL -ErrorAction SilentlyContinue
    }
}