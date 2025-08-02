# Excel Taxonomy Extractor v1.6.0 - One-Click Installation Script
# Repository: https://github.com/henkisdabro/excel-taxonomy-cleaner
# Usage: 
#   Install: irm "https://raw.githubusercontent.com/henkisdabro/excel-taxonomy-cleaner/main/install.ps1" | iex

[CmdletBinding()]
param(
    [string]$Version = "latest"
)

# Configuration
$RepoOwner = "henkisdabro"
$RepoName = "excel-taxonomy-cleaner"
$AddInName = "ipg_taxonomy_extractor_addonv1.6.0.xlam"
$DisplayName = "Excel Taxonomy Extractor AddIn v1.6.0"

# Paths
$AddInsPath = "$env:APPDATA\Microsoft\AddIns"
$AddInPath = Join-Path $AddInsPath $AddInName
$TempPath = Join-Path $env:TEMP "taxonomy-extractor-install"

# Global variables for progress tracking
$Global:InstallSteps = @()
$Global:CurrentStep = 0
$Global:StartTime = Get-Date

function Initialize-ProgressTracker {
    $Global:InstallSteps = @(
        @{ Name = "Validating Environment"; Status = "pending"; Icon = "🔍"; Time = $null },
        @{ Name = "Fetching Latest Release"; Status = "pending"; Icon = "📡"; Time = $null },
        @{ Name = "Downloading AddIn"; Status = "pending"; Icon = "⬇️"; Time = $null },
        @{ Name = "Cleaning Old Versions"; Status = "pending"; Icon = "🧹"; Time = $null },
        @{ Name = "Registry Cleanup"; Status = "pending"; Icon = "🧼"; Time = $null },
        @{ Name = "Installing AddIn"; Status = "pending"; Icon = "📦"; Time = $null },
        @{ Name = "Configuring Security"; Status = "pending"; Icon = "🔐"; Time = $null },
        @{ Name = "Registry Registration"; Status = "pending"; Icon = "📝"; Time = $null },
        @{ Name = "Finalizing Setup"; Status = "pending"; Icon = "🎯"; Time = $null }
    )
}

function Write-ProgressBar {
    param([int]$Percent, [string]$Label = "", [string]$Color = "Green")
    
    $barWidth = 50
    $filled = [Math]::Floor($barWidth * $Percent / 100)
    $empty = $barWidth - $filled
    
    $bar = "█" * $filled + "░" * $empty
    $percentStr = "$Percent%".PadLeft(4)
    $barText = "[$bar] $percentStr $Label"
    
    Write-Host $barText.PadRight(77) -ForegroundColor $Color -NoNewline
}

function Update-StepStatus {
    param([int]$StepIndex, [string]$Status, [string]$Message = "")
    
    if ($StepIndex -lt $Global:InstallSteps.Count) {
        $Global:InstallSteps[$StepIndex].Status = $Status
        $Global:InstallSteps[$StepIndex].Time = (Get-Date) - $Global:StartTime
        
        if ($Status -eq "running") {
            $Global:CurrentStep = $StepIndex
        }
        
        Update-ProgressDisplay $Message
    }
}

function Show-Spinner {
    param([string]$Message = "Processing...")
    
    $spinChars = @("⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏")
    $counter = 0
    
    for ($i = 0; $i -lt 10; $i++) {
        Write-Host "`r  $($spinChars[$counter % $spinChars.Length]) $Message" -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 100
        $counter++
    }
    Write-Host ""
}

function Update-ProgressDisplay {
    param([string]$CurrentMessage = "")
    
    # Simple and reliable approach: just clear screen and redraw
    # This avoids cursor positioning issues that cause duplication
    Clear-Host
    
    # Header
    Write-Host "┌───────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor White
    Write-Host "│" -ForegroundColor White -NoNewline
    Write-Host "  📊 INSTALLATION PROGRESS".PadRight(79) -ForegroundColor White -NoNewline
    Write-Host "│" -ForegroundColor White
    Write-Host "├───────────────────────────────────────────────────────────────────────────────┤" -ForegroundColor White
    
    # Progress steps
    for ($i = 0; $i -lt $Global:InstallSteps.Count; $i++) {
        $step = $Global:InstallSteps[$i]
        $statusIcon = switch ($step.Status) {
            "completed" { "✅" }
            "running" { "⚡" }
            "failed" { "❌" }
            default { "⏳" }
        }
        
        $color = switch ($step.Status) {
            "completed" { "Green" }
            "running" { "Yellow" }
            "failed" { "Red" }
            default { "Gray" }
        }
        
        $timeStr = if ($step.Time) { " ({0:F1}s)" -f $step.Time.TotalSeconds } else { "" }
        $mainText = "  $statusIcon $($step.Icon) $($step.Name)"
        
        # Calculate total length and ensure proper alignment
        $totalTextLength = $mainText.Length + $timeStr.Length
        $paddingNeeded = [Math]::Max(0, 78 - $totalTextLength)
        
        Write-Host "│" -ForegroundColor White -NoNewline
        Write-Host $mainText -ForegroundColor $color -NoNewline
        if ($step.Time) {
            Write-Host $timeStr -ForegroundColor Gray -NoNewline
        }
        Write-Host (" " * $paddingNeeded) -NoNewline
        Write-Host "│" -ForegroundColor White
    }
    
    # Overall progress bar - calculate based on current step being worked on
    $completedSteps = ($Global:InstallSteps | Where-Object { $_.Status -eq "completed" }).Count
    $currentRunningStep = -1
    
    # Find the currently running step
    for ($i = 0; $i -lt $Global:InstallSteps.Count; $i++) {
        if ($Global:InstallSteps[$i].Status -eq "running") {
            $currentRunningStep = $i
            break
        }
    }
    
    # Progress calculation: 0%, 11%, 22%, 33%, 44%, 55%, 66%, 77%, 88%, 100%
    if ($completedSteps -eq $Global:InstallSteps.Count) {
        # All steps completed
        $overallPercent = 100
    } elseif ($currentRunningStep -ge 0) {
        # A step is currently running - show progress for that step
        $overallPercent = $currentRunningStep * 11
    } else {
        # No steps running, show progress based on completed steps
        $overallPercent = $completedSteps * 11
    }
    
    Write-Host "├───────────────────────────────────────────────────────────────────────────────┤" -ForegroundColor White
    Write-Host "│" -ForegroundColor White -NoNewline
    Write-Host "  " -NoNewline
    Write-ProgressBar $overallPercent "Overall Progress" "Cyan"
    Write-Host "│" -ForegroundColor White
    
    # Current action
    if ($CurrentMessage) {
        Write-Host "│" -ForegroundColor White -NoNewline
        Write-Host "  🔄 $CurrentMessage".PadRight(79) -ForegroundColor Yellow -NoNewline
        Write-Host "│" -ForegroundColor White
    }
    
    Write-Host "└───────────────────────────────────────────────────────────────────────────────┘" -ForegroundColor White
}

function Write-Status {
    param([string]$Message, [string]$Color = "Green")
    Write-Host "  🔄 $Message" -ForegroundColor $Color
}

function Write-Error {
    param([string]$Message)
    Write-Host "  ❌ ERROR: $Message" -ForegroundColor Red
}

function Write-Success {
    param([string]$Message)
    Write-Host "  ✅ $Message" -ForegroundColor Green
}

function Write-Logo {
    $logo = @"
::::::::::     :::::::::::::::::::::::        xxxxxxxxxxxxxxx          
::::::::::     ::::::::::::::::::::::::::  xxxxxxxxxxxxxxxxxxxxx       
::::::::::     ::::::::::::::::::::::::::x&xxxxxxxxxxxxxxxxxxxxxxx     
::::::::::     ::::::::::::::::::::::::;&&&&xxxxxxxxxxxxxxxxxxxxx      
::::::::::     :::::::::::::::::::::::X&&&&&&xxxxxxxxxxxxxxxxxx        
::::::::::     ::::::::::         :::X&&&&&&&xxxxx      xxxx           
::::::::::     ::::::::::          :x&&&&&&&&Xxx                       
::::::::::     ::::::::::          $&&&&&&&&&&X                        
::::::::::     ::::::::::          &&&&&&&&&&&       xxxxxxxxxxxxxxxxxx
::::::::::     ::::::::::         :;&&&&&&&&&&       xxxxxxxxxxxxxxxxxx
::::::::::     ::::::::::::::::::::;&&&&&&&&&x       xxxxxxxxxxxxxxxxxx
::::::::::     ::::::::::::::::::::;$&&&&&&Xxxx      xxxxxxxxxxxxxxxxxx
::::::::::     :::::::::::::::::::::;&&&&&xxxxxx      xxxxxxxxxxxxxxxxX
::::::::::     ::::::::::::::::::::::x&&xxxxxxxxxxx     xxxxxxxxxxxxxx 
::::::::::     :::::::::::::::::::::  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx  
::::::::::     ::::::::::              xxxxxxxxxxxxxxxxxxxxxxxxxxxxx   
::::::::::     ::::::::::                xxxxxxxxxxxxxxxxxxxxxxxxx     
::::::::::     ::::::::::                  xxxxxxxxxxxxxxxxxxxxx       
::::::::::     ::::::::::                     xxxxxxxxxxxxxxX  

🏢 IPG MEDIABRANDS TAXONOMY EXTRACTOR ADDIN FOR EXCEL v1.6.0 🏢

"@
    Write-Host $logo -ForegroundColor Cyan
}

function Write-Header {
    param([string]$Title, [string]$Subtitle = "")
    
    Write-Host ""
    Write-Host "┌───────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor White
    Write-Host "│" -ForegroundColor White -NoNewline
    Write-Host ("  " + $Title).PadRight(79) -ForegroundColor White -NoNewline
    Write-Host "│" -ForegroundColor White
    
    if ($Subtitle) {
        Write-Host "│" -ForegroundColor White -NoNewline
        Write-Host ("  " + $Subtitle).PadRight(79) -ForegroundColor Gray -NoNewline
        Write-Host "│" -ForegroundColor White
    }
    
    Write-Host "└───────────────────────────────────────────────────────────────────────────────┘" -ForegroundColor White
    Write-Host ""
}

function Remove-OrphanedAddinRegistryKeys {
    param(
        [string]$CurrentAddInName
    )
    
    try {
        Write-Status "🧼 Cleaning orphaned AddIn registry keys..." "Yellow"
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        
        # Ensure registry path exists
        if (-not (Test-Path $regPath)) {
            Write-Status "Excel registry path not found - skipping registry cleanup" "Yellow"
            return
        }
        
        $removedCount = 0
        $counter = 0
        
        # Check OPEN and OPEN1, OPEN2, etc. up to OPEN50
        do {
            $keyName = if ($counter -eq 0) { "OPEN" } else { "OPEN$counter" }
            
            try {
                $regValue = Get-ItemProperty -Path $regPath -Name $keyName -ErrorAction SilentlyContinue
                
                if ($regValue -and $regValue.$keyName) {
                    $registryValue = $regValue.$keyName
                    $shouldDelete = $false
                    
                    # Remove quotes if present and extract filename from registry value
                    $cleanRegistryValue = $registryValue.Trim('"')
                    $registryFileName = [System.IO.Path]::GetFileName($cleanRegistryValue)
                    
                    Write-Status "Checking registry key '$keyName' with value: $registryValue" "Gray"
                    Write-Status "  Extracted filename: $registryFileName" "Gray"
                    
                    # Check if registry value contains our AddIn patterns and is not the current version
                    if (($registryFileName -like "ipg_taxonomy_extractor_addon*") -and 
                        ($registryFileName -ne $CurrentAddInName)) {
                        $shouldDelete = $true
                        Write-Status "🔍 Found old AddIn registry key '$keyName' referencing: $registryFileName" "Yellow"
                    }
                    elseif (($registryFileName -like "TaxonomyExtractor*") -and 
                            ($registryFileName -ne $CurrentAddInName)) {
                        $shouldDelete = $true
                        Write-Status "🔍 Found old AddIn registry key '$keyName' referencing: $registryFileName" "Yellow"
                    }
                    elseif (($registryFileName -like "taxonomy_extractor*") -and 
                            ($registryFileName -ne $CurrentAddInName)) {
                        $shouldDelete = $true
                        Write-Status "🔍 Found old AddIn registry key '$keyName' referencing: $registryFileName" "Yellow"
                    }
                    
                    # Remove the old registry key
                    if ($shouldDelete) {
                        Remove-ItemProperty -Path $regPath -Name $keyName -ErrorAction Stop
                        Write-Success "🗑️ Removed old registry key: $keyName (was: $registryFileName)"
                        $removedCount++
                    }
                }
            }
            catch {
                Write-Status "Failed to process registry key '$keyName': $($_.Exception.Message)" "Red"
            }
            
            $counter++
        } while ($counter -lt 50)
        
        if ($removedCount -gt 0) {
            Write-Success "Cleaned $removedCount old registry entries"
        } else {
            Write-Status "No old registry entries found to clean"
        }
    }
    catch {
        Write-Status "Registry cleanup failed: $($_.Exception.Message)" "Red"
        # Don't throw - registry cleanup failure shouldn't stop installation
    }
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
            throw "AddIn file '$AddInName' not found in latest release"
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
        # Initialize the progress tracker
        Initialize-ProgressTracker
        Update-ProgressDisplay
        Start-Sleep -Milliseconds 500
        
        # Step 1: Verify Excel is installed
        Update-StepStatus 0 "running" "Checking Excel installation..."
        Show-Spinner "Validating Excel environment"
        
        if (-not (Test-ExcelInstalled)) {
            Update-StepStatus 0 "failed"
            throw "Microsoft Excel is not installed or cannot be accessed"
        }
        Update-StepStatus 0 "completed"

        # Step 2: Get latest release
        Update-StepStatus 1 "running" "Fetching release information from GitHub..."
        Show-Spinner "Contacting GitHub API"
        $releaseInfo = Get-LatestReleaseUrl
        Update-StepStatus 1 "completed"
        Write-Status "Found version: $($releaseInfo.Version)" "Cyan"

        # Step 3: Download AddIn
        Update-StepStatus 2 "running" "Downloading latest AddIn file..."
        
        # Create temp directory
        if (Test-Path $TempPath) {
            Remove-Item $TempPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null

        $tempAddInPath = Join-Path $TempPath $AddInName
        Show-Spinner "Downloading from GitHub releases"
        Invoke-WebRequest -Uri $releaseInfo.DownloadUrl -OutFile $tempAddInPath -UseBasicParsing

        # Verify download
        if (-not (Test-Path $tempAddInPath)) {
            Update-StepStatus 2 "failed"
            throw "Failed to download AddIn file"
        }
        Update-StepStatus 2 "completed"

        # Ensure AddIns directory exists
        if (-not (Test-Path $AddInsPath)) {
            New-Item -ItemType Directory -Path $AddInsPath -Force | Out-Null
        }

        # Step 4: Remove old versions before installing new one
        Update-StepStatus 3 "running" "Scanning for old versions to remove..."
        Show-Spinner "Analyzing installed AddIns"
        
        # Get all XLAM files in the AddIns directory
        $allXlamFiles = Get-ChildItem -Path $AddInsPath -Filter "*.xlam" -ErrorAction SilentlyContinue
        
        foreach ($file in $allXlamFiles) {
            $fileName = $file.Name
            $shouldDelete = $false
            
            # Check if it matches our taxonomy extractor patterns
            if ($fileName -like "ipg_taxonomy_extractor_addon*" -or 
                $fileName -like "TaxonomyExtractor*" -or 
                $fileName -like "taxonomy_extractor*") {
                $shouldDelete = $true
            }
            
            # Skip the current version we're about to install
            if ($fileName -eq $AddInName) {
                $shouldDelete = $false
            }
            
            # Delete old versions
            if ($shouldDelete) {
                try {
                    Remove-Item $file.FullName -Force -ErrorAction Stop
                    Write-Status "🗑️ Removed old version: $fileName" "Yellow"
                } catch {
                    Write-Status "Failed to remove $fileName`: $($_.Exception.Message)" "Red"
                }
            }
        }
        
        Update-StepStatus 3 "completed"
        
        # Step 5: Clean orphaned registry keys
        Update-StepStatus 4 "running" "Cleaning registry entries..."
        Show-Spinner "Scanning Windows registry"
        Remove-OrphanedAddinRegistryKeys -CurrentAddInName $AddInName
        Update-StepStatus 4 "completed"

        # Step 6: Install to AddIns folder
        Update-StepStatus 5 "running" "Installing AddIn to Excel folder..."
        Show-Spinner "Copying files to AddIns directory"
        Copy-Item $tempAddInPath -Destination $AddInPath -Force

        # Verify installation was successful
        if (-not (Test-Path $AddInPath)) {
            Update-StepStatus 5 "failed"
            throw "Failed to copy AddIn file to destination"
        }
        Update-StepStatus 5 "completed"

        # Step 7: Configure security
        Update-StepStatus 6 "running" "Configuring security permissions..."
        Show-Spinner "Unblocking files and setting permissions"
        Unblock-File -Path $AddInPath -ErrorAction SilentlyContinue
        Update-StepStatus 6 "completed"

        # Step 8: Register AddIn in Excel registry
        Update-StepStatus 7 "running" "Registering with Excel..."
        Show-Spinner "Creating registry entries"
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        
        # Ensure registry path exists
        if (-not (Test-Path $regPath)) {
            Write-Status "Creating Excel registry path..."
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Check if current version is already registered
        $currentVersionRegistered = $false
        $existingKeyName = ""
        $counter = 0
        
        # First pass: check if current version already exists
        do {
            $keyName = if ($counter -eq 0) { "OPEN" } else { "OPEN$counter" }
            $existingValue = Get-ItemProperty -Path $regPath -Name $keyName -ErrorAction SilentlyContinue
            
            if ($existingValue -and $existingValue.$keyName) {
                $existingRegistryValue = $existingValue.$keyName
                $cleanExistingValue = $existingRegistryValue.Trim('"')
                $existingFileName = [System.IO.Path]::GetFileName($cleanExistingValue)
                
                if ($existingFileName -eq $AddInName) {
                    $currentVersionRegistered = $true
                    $existingKeyName = $keyName
                    Write-Status "Current version already registered as '$keyName'" "Green"
                    break
                }
            }
            $counter++
        } while ($counter -lt 50)
        
        # If not already registered, find next available slot
        if (-not $currentVersionRegistered) {
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
                # Use filename-only format with quotes as required by Excel
                $registryValue = "`"$AddInName`""
                New-ItemProperty -Path $regPath -Name $openNumber -Value $registryValue -PropertyType String -Force | Out-Null
                Write-Success "Registered AddIn as '$openNumber' with value: $registryValue"
            } else {
                Write-Status "Warning: Could not find available OPEN slot in registry (checked OPEN0-OPEN49)" "Yellow"
            }
        }
        Update-StepStatus 7 "completed"

        # Step 9: Finalize setup
        Update-StepStatus 8 "running" "Finalizing installation..."
        Show-Spinner "Setting up trusted locations and creating shortcuts"
        $trustRegPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
        $locationPath = "$trustRegPath\Location99"
        
        if (-not (Test-Path $locationPath)) {
            New-Item -Path $locationPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $locationPath -Name "Path" -Value "$AddInsPath\"
        Set-ItemProperty -Path $locationPath -Name "AllowSubFolders" -Value 1
        Set-ItemProperty -Path $locationPath -Name "Description" -Value "Excel AddIns (Auto-trusted by Taxonomy Extractor installer)"


        # Cleanup
        Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue
        Update-StepStatus 8 "completed"

        # Brief pause to show 100% completion
        Start-Sleep -Milliseconds 1000
        
        Write-Header "🎉 INSTALLATION COMPLETE!" "Open Excel and find the IPG Taxonomy Extractor button in the Home tab"
        
        Write-Host "📂 AddIn location: " -ForegroundColor Gray -NoNewline
        Write-Host "$AddInPath" -ForegroundColor Cyan
        Write-Host "🎯 IPG Taxonomy Extractor button will appear in IPG Tools group on Home tab" -ForegroundColor Gray
        Write-Host ""
        Write-Host "🗑️  " -ForegroundColor Red -NoNewline
        Write-Host "TO UNINSTALL:" -ForegroundColor Yellow
        Write-Host "   Go to File → Options → Add-ins → Excel Add-ins → Go → Uncheck the AddIn" -ForegroundColor Gray
        
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


# Main execution
try {
    Clear-Host
    
    # Startup sequence with animation
    Write-Host "⚡ " -ForegroundColor Yellow -NoNewline
    Write-Host "Initializing IPG Taxonomy Extractor AddIn Installer..." -ForegroundColor White
    Show-Spinner "Loading installer components"
    
    Clear-Host
    Write-Logo
    Write-Header "🚀 AUTOMATED INSTALLER" "One-click installation with smart upgrade handling"
    
    Write-Host "📍 Repository: " -ForegroundColor Gray -NoNewline
    Write-Host "https://github.com/$RepoOwner/$RepoName" -ForegroundColor Cyan
    Write-Host ""
    
    # Interactive prompt
    Write-Host "┌───────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor White
    Write-Host "│" -ForegroundColor White -NoNewline
    Write-Host "  🎯 Ready to install Excel Taxonomy Extractor AddIn v1.6.0?".PadRight(79) -ForegroundColor White -NoNewline
    Write-Host "│" -ForegroundColor White
    Write-Host "│" -ForegroundColor White -NoNewline
    Write-Host "  📦 This will automatically download, install, and configure the AddIn".PadRight(79) -ForegroundColor Gray -NoNewline
    Write-Host "│" -ForegroundColor White
    Write-Host "└───────────────────────────────────────────────────────────────────────────────┘" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Press " -ForegroundColor Gray -NoNewline
    Write-Host "ENTER" -ForegroundColor Yellow -NoNewline
    Write-Host " to continue or " -ForegroundColor Gray -NoNewline
    Write-Host "CTRL+C" -ForegroundColor Red -NoNewline
    Write-Host " to cancel..." -ForegroundColor Gray
    Read-Host
    
    Clear-Host
    Write-Logo

    Install-AddIn
}
catch {
    Write-Host ""
    Write-Host "💥 INSTALLATION FAILED" -ForegroundColor Red -BackgroundColor Black
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "📞 Need help? Create an issue at: https://github.com/$RepoOwner/$RepoName/issues" -ForegroundColor Yellow
    exit 1
}