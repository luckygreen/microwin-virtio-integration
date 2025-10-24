# Filename: Integrate-VirtIO-MicroWin.ps1
# Version: 3.8.0
# Date: 2025-10-24T18:30:00Z
# Authors: Claude & Lucky Green <shamrock@cypherpunks.to>
#
# Purpose:
#   Post-processes MicroWin Windows 11 ISOs to achieve complete "first-boot-ready" 
#   VirtIO integration for Proxmox VE virtual machines. This script eliminates all 
#   manual driver installation steps by:
#
#   - Injecting all 10 VirtIO drivers into install.wim (for the installed system)
#   - Injecting all 10 VirtIO drivers into boot.wim (for the Windows installer/WinPE)
#   - Bundling virtio-win-guest-tools.exe and configuring automatic silent installation
#     post-OOBE via SetupComplete.cmd
#   - Generating properly bootable ISOs with UEFI/BIOS dual-boot support
#   - Auto-detecting source files in the working directory
#   - Smart ISO naming with descriptive volume labels (under 32 characters)
#
#   The script solves the common "no drives found" error during Windows installation
#   by ensuring VirtIO storage drivers are available in the installer environment,
#   eliminating the need for IDE workarounds or manual driver loading.
#
#   The authors strongly believe this functionality would best be implemented directly
#   in MicroWin itself, and we encourage and hope the MicroWin development team will
#   consider integrating this capability into their excellent tool. Until then, this
#   script bridges the gap for Proxmox VE users who want fully automated VirtIO
#   deployments.
#
# Usage:
#   .\Integrate-VirtIO-MicroWin.ps1
#   
#   Optional parameters:
#   -MicroWinISO <path>     Path to MicroWin ISO (auto-detected if not specified)
#   -VirtIOISO <path>       Path to VirtIO drivers ISO (auto-detected if not specified)
#   -GuestToolsExe <path>   Path to virtio-win-guest-tools.exe (auto-detected if not specified)
#
#   Example with explicit paths:
#   .\Integrate-VirtIO-MicroWin.ps1 -MicroWinISO ".\MicroWin11_25H2_English_x64.iso" `
#                                    -VirtIOISO ".\virtio-win-0.1.285.iso" `
#                                    -GuestToolsExe ".\virtio-win-guest-tools.exe"
#
# Copyright & License:
#   This is free and unencumbered software released into the public domain.
#
#   Anyone is free to copy, modify, publish, use, compile, sell, or distribute this
#   software, either in source code form or as a compiled binary, for any purpose,
#   commercial or non-commercial, and by any means.
#
#   In jurisdictions that recognize copyright laws, the author or authors of this
#   software dedicate any and all copyright interest in the software to the public
#   domain. We make this dedication for the benefit of the public at large and to
#   the detriment of our heirs and successors. We intend this dedication to be an
#   overt act of relinquishment in perpetuity of all present and future rights to
#   this software under copyright law.
#
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#   AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
#   ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
#   WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#
#   For more information, please refer to <https://unlicense.org/>
#
# Change Log:
#   v3.8.0 (2025-10-24) - First public release
#     - Enhanced header with expanded purpose and public domain dedication
#     - Improved comments throughout for clarity and maintainability
#     - Added comprehensive usage examples
#     - No functional changes from v3.7.0

#Requires -RunAsAdministrator

param(
    [Parameter(Mandatory=$false)]
    [string]$MicroWinISO,      # Path to MicroWin-processed Windows 11 ISO
    
    [Parameter(Mandatory=$false)]
    [string]$VirtIOISO,        # Path to VirtIO drivers ISO (virtio-win-X.X.XXX.iso)
    
    [Parameter(Mandatory=$false)]
    [string]$GuestToolsExe     # Path to virtio-win-guest-tools.exe installer
)

#
# Function: Write-Progress-Status
# Purpose: Display progress information both in the progress bar and console
# Parameters:
#   - Activity: Description of the overall operation
#   - Status: Current step being performed
#   - PercentComplete: Progress percentage (0-100)
#
function Write-Progress-Status {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    Write-Host "[$PercentComplete%] $Status" -ForegroundColor Cyan
}

# Get the directory where this script is located
$ScriptDir = $PSScriptRoot

#
# Function: Resolve-FullPath
# Purpose: Convert relative paths to absolute paths, resolving them relative to script location
# Parameters:
#   - Path: Path to resolve (can be relative or already absolute)
# Returns: Fully qualified absolute path
#
function Resolve-FullPath {
    param([string]$Path)
    if ($Path -and -not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path $ScriptDir $Path
        $Path = [System.IO.Path]::GetFullPath($Path)
    }
    return $Path
}

#
# Function: Get-VirtIOISOVersion
# Purpose: Extract version number from VirtIO ISO filename for smart naming
# Parameters:
#   - FileName: Name of the VirtIO ISO file
# Returns: Version object (e.g., 0.1.285) or 0.0.0 if version cannot be determined
#
function Get-VirtIOISOVersion {
    param([string]$FileName)
    
    if ($FileName -match 'virtio-win-(\d+\.\d+\.\d+)\.iso') {
        return [version]$matches[1]
    }
    return [version]"0.0.0"
}

#
# Function: Get-ExeVersion
# Purpose: Read version information from executable file metadata
# Parameters:
#   - ExePath: Full path to executable file
# Returns: Version object extracted from file properties, or 0.0.0 if unavailable
#
function Get-ExeVersion {
    param([string]$ExePath)
    
    try {
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
        if ($versionInfo.FileVersion) {
            # Remove trailing .0 for cleaner version strings
            $version = $versionInfo.FileVersion -replace '\.0$', ''
            return [version]$version
        }
    } catch {
        Write-Warning "Could not read version from $ExePath"
    }
    return [version]"0.0.0"
}

#
# Function: Get-ISOType
# Purpose: Identify the type of ISO by mounting and inspecting its contents
# Parameters:
#   - ISOPath: Full path to ISO file to examine
# Returns: String indicating ISO type: "VirtIO", "MicroWin", "WindowsOriginal", or "Unknown"
# Note: Temporarily mounts the ISO and dismounts it after inspection
#
function Get-ISOType {
    param([string]$ISOPath)
    
    try {
        $mount = Mount-DiskImage -ImagePath $ISOPath -PassThru -ErrorAction Stop
        Start-Sleep -Milliseconds 500  # Brief pause to ensure mount completes
        $drive = ($mount | Get-Volume).DriveLetter + ":"
        
        $type = "Unknown"
        
        # VirtIO ISO has driver folders like NetKVM, viostor at root
        if ((Test-Path "$drive\NetKVM") -and (Test-Path "$drive\viostor")) {
            $type = "VirtIO"
        }
        # MicroWin ISO has a VirtIO folder (but not VirtIO drivers at root)
        elseif (Test-Path "$drive\VirtIO") {
            $type = "MicroWin"
        }
        # Original Windows ISO has install.wim in sources folder
        elseif (Test-Path "$drive\sources\install.wim") {
            $type = "WindowsOriginal"
        }
        
        Dismount-DiskImage -ImagePath $ISOPath -ErrorAction SilentlyContinue | Out-Null
        return $type
    } catch {
        Write-Warning "Could not mount/check ISO: $ISOPath"
        return "Unknown"
    }
}

#
# Function: Get-OutputISOName
# Purpose: Construct intelligent output ISO filename from input ISO names and versions
# Parameters:
#   - MicroWinISOPath: Path to the MicroWin ISO
#   - VirtIOISOPath: Path to the VirtIO drivers ISO
# Returns: Filename string in format: MicroWin11_25H2_Eng_x64_VIO285.iso
# Note: Automatically truncates to 32 characters if needed for ISO9660 compatibility
#
function Get-OutputISOName {
    param(
        [string]$MicroWinISOPath,
        [string]$VirtIOISOPath
    )
    
    $microWinName = [System.IO.Path]::GetFileNameWithoutExtension($MicroWinISOPath)
    
    # Extract VirtIO version for output filename suffix
    $virtioName = [System.IO.Path]::GetFileNameWithoutExtension($VirtIOISOPath)
    $virtioShortVersion = "000"
    if ($virtioName -match '(\d+)\.iso$') {
        $virtioShortVersion = $matches[1]
    }
    elseif ($virtioName -match '\.(\d{3})\.iso$') {
        $virtioShortVersion = $matches[1]
    }
    elseif ($virtioName -match '-(\d+\.\d+\.)(\d+)') {
        # Extract just the patch version (e.g., 285 from 0.1.285)
        $virtioShortVersion = $matches[2]
    }
    
    # Parse MicroWin filename if it follows standard naming convention
    if ($microWinName -match '^(MicroWin\d+)_(\d+H\d+)_([^_]+)_(x64|x86)') {
        $winVersion = $matches[1]   # e.g., MicroWin11
        $release = $matches[2]       # e.g., 25H2
        $language = $matches[3]      # e.g., English
        $arch = $matches[4]          # e.g., x64
        
        # Abbreviate language to 3 characters for compact naming
        $langAbbr = switch ($language) {
            "English" { "Eng" }
            "EnglishInternational" { "Eng" }
            "Spanish" { "Esp" }
            "French" { "Fra" }
            "German" { "Deu" }
            "Italian" { "Ita" }
            "Portuguese" { "Por" }
            "Chinese" { "Chn" }
            "Japanese" { "Jpn" }
            default { $language.Substring(0, [Math]::Min(3, $language.Length)) }
        }
        
        $outputName = "${winVersion}_${release}_${langAbbr}_${arch}_VIO${virtioShortVersion}"
    }
    else {
        # Fallback for non-standard MicroWin naming
        $outputName = "${microWinName}_VIO${virtioShortVersion}"
    }
    
    # ISO9660 volume labels have a 32-character limit
    if ($outputName.Length -gt 32) {
        Write-Warning "Output name exceeds 32 chars ($($outputName.Length)): $outputName"
        $outputName = $outputName.Substring(0, 32)
        Write-Host "Truncated to: $outputName" -ForegroundColor Yellow
    }
    
    return $outputName + ".iso"
}

# ============================================================================
# MAIN SCRIPT EXECUTION BEGINS HERE
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "VirtIO Integration for MicroWin ISO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Progress-Status -Activity "VirtIO Integration" -Status "Scanning directory for files..." -PercentComplete 0

#
# Auto-Detection Phase
# If ISOs or guest tools exe not explicitly specified, scan the script directory
# to automatically locate the required files
#
if (-not $MicroWinISO -or -not $VirtIOISO -or -not $GuestToolsExe) {
    Write-Host "Scanning directory for required files..." -ForegroundColor Yellow
    
    $allISOs = Get-ChildItem -Path $ScriptDir -Filter "*.iso"
    $virtioISOs = @()
    
    foreach ($iso in $allISOs) {
        $type = Get-ISOType -ISOPath $iso.FullName
        Write-Host "  $($iso.Name): $type" -ForegroundColor Gray
        
        if ($type -eq "MicroWin" -and -not $MicroWinISO) {
            $MicroWinISO = $iso.FullName
        }
        elseif ($type -eq "VirtIO") {
            # Collect all VirtIO ISOs and select the newest version later
            $version = Get-VirtIOISOVersion -FileName $iso.Name
            $virtioISOs += @{
                Path = $iso.FullName
                Name = $iso.Name
                Version = $version
            }
        }
    }
    
    # Select the newest VirtIO ISO if multiple versions found
    if ($virtioISOs.Count -gt 0 -and -not $VirtIOISO) {
        $selectedVirtIO = $virtioISOs | Sort-Object { $_.Version } -Descending | Select-Object -First 1
        $VirtIOISO = $selectedVirtIO.Path
        $selectedISOVersion = $selectedVirtIO.Version
        
        # Look for matching guest tools executable
        $guestToolsName = "virtio-win-guest-tools.exe"
        $guestToolsSearch = Get-ChildItem -Path $ScriptDir -Filter $guestToolsName -ErrorAction SilentlyContinue
        
        if ($guestToolsSearch -and -not $GuestToolsExe) {
            $guestToolsVersion = Get-ExeVersion -ExePath $guestToolsSearch.FullName
            
            # Verify that guest tools version matches VirtIO ISO version (at least major.minor)
            $versionMatch = $false
            if ($guestToolsVersion -and $selectedISOVersion) {
                $guestToolsMajorMinor = "$($guestToolsVersion.Major).$($guestToolsVersion.Minor)"
                $isoMajorMinor = "$($selectedISOVersion.Major).$($selectedISOVersion.Minor)"
                $versionMatch = ($guestToolsMajorMinor -eq $isoMajorMinor)
            }
            
            if ($versionMatch) {
                Write-Host "  ${guestToolsName}: Found (v$guestToolsVersion)" -ForegroundColor Gray
                $GuestToolsExe = $guestToolsSearch.FullName
            }
            else {
                Write-Warning "  ${guestToolsName}: Version mismatch (exe: v$guestToolsVersion, ISO: v$selectedISOVersion)"
                Write-Host "  Download matching version from: https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/latest-virtio/" -ForegroundColor Yellow
            }
        }
    }
}

# Convert any relative paths to absolute paths
$MicroWinISO = Resolve-FullPath $MicroWinISO
$VirtIOISO = Resolve-FullPath $VirtIOISO
$GuestToolsExe = Resolve-FullPath $GuestToolsExe

# Validate that required files exist
if (-not $MicroWinISO -or -not (Test-Path $MicroWinISO)) {
    Write-Error "MicroWin ISO not found"
    exit 1
}

if (-not $VirtIOISO -or -not (Test-Path $VirtIOISO)) {
    Write-Error "VirtIO ISO not found"
    exit 1
}

# Display the files that will be used
Write-Host "`nUsing:" -ForegroundColor Green
Write-Host "  MicroWin ISO: $MicroWinISO" -ForegroundColor White
Write-Host "  VirtIO ISO: $VirtIOISO" -ForegroundColor White
if ($GuestToolsExe) {
    Write-Host "  Guest Tools: $GuestToolsExe" -ForegroundColor White
}
Write-Host ""

#
# Download oscdimg.exe if not already present
# oscdimg is the Microsoft tool for creating bootable ISO images
# We download it from ChrisTitusTech's WinUtil repository for reliability
#
$oscdimgPath = Join-Path $ScriptDir "oscdimg.exe"
if (-not (Test-Path $oscdimgPath)) {
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Downloading oscdimg.exe from CTT GitHub..." -PercentComplete 7
    $oscdimgUrl = "https://github.com/ChrisTitusTech/winutil/raw/main/releases/oscdimg.exe"
    $expectedHash = "AB9E161049D293B544961BFDF2D61244ADE79376D6423DF4F60BF9B147D3C78D"
    
    Invoke-WebRequest -Uri $oscdimgUrl -OutFile $oscdimgPath
    
    # Verify file integrity with SHA256 hash
    $actualHash = (Get-FileHash -Path $oscdimgPath -Algorithm SHA256).Hash
    if ($actualHash -ne $expectedHash) {
        Write-Error "oscdimg.exe hash mismatch!"
        Remove-Item $oscdimgPath
        exit 1
    }
    Write-Host "oscdimg.exe verified" -ForegroundColor Green
} else {
    Write-Host "oscdimg.exe already present" -ForegroundColor Green
}

Write-Progress-Status -Activity "VirtIO Integration" -Status "Creating working directories..." -PercentComplete 10

# Set up working directory structure
$WorkDir = Join-Path $ScriptDir "VirtIO_Integration_TMP"
$MountWIM = Join-Path $WorkDir "Mount_WIM"          # Mount point for install.wim
$MountBootWIM = Join-Path $WorkDir "Mount_BootWIM"  # Mount point for boot.wim
$ExtractISO = Join-Path $WorkDir "Extract_ISO"      # Extracted ISO contents

# Clean up any previous working directory and create fresh structure
Remove-Item -Path $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $WorkDir, $MountWIM, $MountBootWIM, $ExtractISO | Out-Null

# Flag to control whether cleanup happens on exit (preserved on failure for debugging)
$script:CleanupOnExit = $true

try {
    # ========================================================================
    # Phase 1: Mount source ISOs
    # ========================================================================
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Mounting VirtIO ISO..." -PercentComplete 15
    $VirtIOMount = Mount-DiskImage -ImagePath $VirtIOISO -PassThru -ErrorAction Stop
    $VirtIODrive = ($VirtIOMount | Get-Volume).DriveLetter + ":"
    Write-Host "VirtIO mounted at: $VirtIODrive" -ForegroundColor Green

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Mounting MicroWin ISO..." -PercentComplete 20
    $MicroWinMount = Mount-DiskImage -ImagePath $MicroWinISO -PassThru -ErrorAction Stop
    $MicroWinDrive = ($MicroWinMount | Get-Volume).DriveLetter + ":"
    Write-Host "MicroWin mounted at: $MicroWinDrive" -ForegroundColor Green

    # ========================================================================
    # Phase 2: Copy ISO contents to working directory
    # ========================================================================
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Copying ISO contents (this will take a while)..." -PercentComplete 25
    $startTime = Get-Date
    
    # Copy with explicit error handling to catch failures immediately
    try {
        Copy-Item -Path "$MicroWinDrive\*" -Destination $ExtractISO -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to copy ISO contents: $_"
        $script:CleanupOnExit = $false
        throw
    }
    
    $elapsed = ((Get-Date) - $startTime).TotalSeconds
    Write-Host "Copy complete! Time taken: $([int]$elapsed) seconds" -ForegroundColor Green
    
    # Verify that files were actually copied
    $copiedFiles = Get-ChildItem -Path $ExtractISO -Recurse -File
    Write-Host "Files copied: $($copiedFiles.Count)" -ForegroundColor Gray
    
    if ($copiedFiles.Count -eq 0) {
        Write-Error "No files were copied from ISO"
        $script:CleanupOnExit = $false
        throw "Copy operation failed - no files in destination"
    }
    
    # ========================================================================
    # Phase 3: Verify boot files exist
    # Critical for creating bootable ISOs
    # ========================================================================
    
    $bootFile = Join-Path $ExtractISO "boot\etfsboot.com"  # BIOS boot file
    $efiFile = Join-Path $ExtractISO "efi\microsoft\boot\efisys.bin"  # UEFI boot file
    
    if (-not (Test-Path $bootFile)) {
        Write-Error "Boot file missing: $bootFile"
        Write-Host "Directory contents:" -ForegroundColor Yellow
        Get-ChildItem -Path $ExtractISO -Directory | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        $script:CleanupOnExit = $false
        throw "Required boot file not found"
    }
    
    if (-not (Test-Path $efiFile)) {
        Write-Error "EFI boot file missing: $efiFile"
        Write-Host "EFI directory structure:" -ForegroundColor Yellow
        if (Test-Path (Join-Path $ExtractISO "efi")) {
            Get-ChildItem -Path (Join-Path $ExtractISO "efi") -Recurse | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        }
        $script:CleanupOnExit = $false
        throw "Required EFI boot file not found"
    }
    
    Write-Host "Boot files verified:" -ForegroundColor Green
    Write-Host "  BIOS: $bootFile" -ForegroundColor Gray
    Write-Host "  UEFI: $efiFile" -ForegroundColor Gray
    
    # Remove read-only attribute from all copied files to allow modifications
    Get-ChildItem -Path $ExtractISO -Recurse | ForEach-Object { 
        $_.Attributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly) 
    }

    # Locate the WIM files we need to modify
    $WIMPath = Join-Path $ExtractISO "sources\install.wim"
    $BootWIMPath = Join-Path $ExtractISO "sources\boot.wim"
    
    if (-not (Test-Path $WIMPath)) {
        Write-Error "install.wim not found at: $WIMPath"
        $script:CleanupOnExit = $false
        throw "install.wim missing"
    }
    
    if (-not (Test-Path $BootWIMPath)) {
        Write-Error "boot.wim not found at: $BootWIMPath"
        $script:CleanupOnExit = $false
        throw "boot.wim missing"
    }

    # ========================================================================
    # Define VirtIO driver paths
    # These are the essential drivers needed for VirtIO hardware in Proxmox
    # ========================================================================
    
    $DriverPaths = @(
        @{Path="$VirtIODrive\NetKVM\w11\amd64"; Name="Network adapter"},           # Virtual network card
        @{Path="$VirtIODrive\viostor\w11\amd64"; Name="SCSI block driver"},        # VirtIO block storage
        @{Path="$VirtIODrive\vioscsi\w11\amd64"; Name="SCSI controller"},          # VirtIO SCSI controller
        @{Path="$VirtIODrive\Balloon\w11\amd64"; Name="Memory ballooning"},        # Dynamic memory management
        @{Path="$VirtIODrive\viorng\w11\amd64"; Name="RNG device"},                # Random number generator
        @{Path="$VirtIODrive\vioserial\w11\amd64"; Name="Serial driver"},          # Virtual serial port
        @{Path="$VirtIODrive\qemupciserial\w11\amd64"; Name="PCI serial"},         # PCI-based serial
        @{Path="$VirtIODrive\vioinput\w11\amd64"; Name="Input devices"},           # Keyboard/mouse
        @{Path="$VirtIODrive\pvpanic\w11\amd64"; Name="Panic notifier"},           # Guest crash notification
        @{Path="$VirtIODrive\viofs\w11\amd64"; Name="Shared filesystem"}           # VirtIO-FS for folder sharing
    )

    # ========================================================================
    # Phase 4: Inject drivers into INSTALL.WIM
    # This is the Windows image that gets installed to disk
    # ========================================================================
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Getting install.wim information..." -PercentComplete 35
    $WIMInfo = Get-WindowsImage -ImagePath $WIMPath
    $ImageIndex = $WIMInfo[0].ImageIndex  # Use first image index (typically there's only one in MicroWin ISOs)
    Write-Host "Using Image Index: $ImageIndex" -ForegroundColor Green

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Mounting install.wim (this may take a while)..." -PercentComplete 38
    Mount-WindowsImage -ImagePath $WIMPath -Index $ImageIndex -Path $MountWIM -ErrorAction Stop

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Injecting drivers into install.wim..." -PercentComplete 40
    
    $driverCount = $DriverPaths.Count
    $currentDriver = 0
    
    # Inject each VirtIO driver into the Windows installation image
    foreach ($driver in $DriverPaths) {
        $currentDriver++
        $progressPercent = 40 + ([int](($currentDriver / $driverCount) * 15))
        
        if (Test-Path $driver.Path) {
            Write-Progress-Status -Activity "VirtIO Integration" -Status "Adding $($driver.Name) to install.wim..." -PercentComplete $progressPercent
            # -Recurse ensures all .inf files in subdirectories are included
            # -ForceUnsigned allows unsigned drivers (VirtIO drivers are signed by Red Hat)
            Add-WindowsDriver -Path $MountWIM -Driver $driver.Path -Recurse -ForceUnsigned | Out-Null
            Write-Host "  Added to install.wim: $($driver.Name)" -ForegroundColor Green
        } else {
            Write-Warning "  Skipping missing: $($driver.Name)"
        }
    }

    # ========================================================================
    # Phase 5: Add guest tools and SetupComplete.cmd
    # This ensures guest tools install automatically after Windows setup
    # ========================================================================
    
    if ($GuestToolsExe -and (Test-Path $GuestToolsExe)) {
        Write-Progress-Status -Activity "VirtIO Integration" -Status "Copying guest tools and creating SetupComplete.cmd..." -PercentComplete 55
        
        # Copy guest tools executable to Windows directory
        Copy-Item -Path $GuestToolsExe -Destination "$MountWIM\Windows\virtio-win-guest-tools.exe" -Force
        Write-Host "Guest tools copied to Windows directory" -ForegroundColor Green
        
        # Create Scripts directory for SetupComplete.cmd
        $setupScriptsDir = Join-Path $MountWIM "Windows\Setup\Scripts"
        New-Item -ItemType Directory -Force -Path $setupScriptsDir | Out-Null
        
        # SetupComplete.cmd runs automatically after Windows Setup completes but before first login
        $setupCompleteContent = @"
@echo off
REM Auto-install VirtIO Guest Tools after Windows installation
REM This script runs automatically after OOBE completes
echo Installing VirtIO Guest Tools...
start /wait C:\Windows\virtio-win-guest-tools.exe /install /passive /norestart
echo VirtIO Guest Tools installation complete.
"@
        
        $setupCompletePath = Join-Path $setupScriptsDir "SetupComplete.cmd"
        Set-Content -Path $setupCompletePath -Value $setupCompleteContent -Encoding ASCII
        Write-Host "SetupComplete.cmd created for auto-installation" -ForegroundColor Green
        
        $includeGuestTools = $true
    } else {
        $includeGuestTools = $false
    }

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Saving and unmounting install.wim..." -PercentComplete 58
    # -Save commits all changes back to the WIM file (this takes several minutes)
    Dismount-WindowsImage -Path $MountWIM -Save -ErrorAction Stop

    # ========================================================================
    # Phase 6: Inject drivers into BOOT.WIM
    # This is the Windows PE environment used during installation
    # Without these drivers, the installer cannot see VirtIO storage devices
    # ========================================================================
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Mounting boot.wim for installer drivers..." -PercentComplete 60
    # Boot.wim index 2 is typically the Windows Setup environment (index 1 is usually the recovery environment)
    Mount-WindowsImage -ImagePath $BootWIMPath -Index 2 -Path $MountBootWIM -ErrorAction Stop
    Write-Host "boot.wim mounted for driver injection" -ForegroundColor Green

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Injecting drivers into boot.wim..." -PercentComplete 62
    
    $currentDriver = 0
    # Inject the same VirtIO drivers into the boot environment
    foreach ($driver in $DriverPaths) {
        $currentDriver++
        $progressPercent = 62 + ([int](($currentDriver / $driverCount) * 13))
        
        if (Test-Path $driver.Path) {
            Write-Progress-Status -Activity "VirtIO Integration" -Status "Adding $($driver.Name) to boot.wim..." -PercentComplete $progressPercent
            Add-WindowsDriver -Path $MountBootWIM -Driver $driver.Path -Recurse -ForceUnsigned | Out-Null
            Write-Host "  Added to boot.wim: $($driver.Name)" -ForegroundColor Green
        } else {
            Write-Warning "  Skipping missing: $($driver.Name)"
        }
    }

    Write-Progress-Status -Activity "VirtIO Integration" -Status "Saving and unmounting boot.wim..." -PercentComplete 75
    Dismount-WindowsImage -Path $MountBootWIM -Save -ErrorAction Stop
    Write-Host "boot.wim updated - installer can now see VirtIO storage!" -ForegroundColor Green

    # ========================================================================
    # Phase 7: Create new bootable ISO
    # ========================================================================
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Creating new ISO..." -PercentComplete 80
    $NewISOName = Get-OutputISOName -MicroWinISOPath $MicroWinISO -VirtIOISOPath $VirtIOISO
    $NewISOPath = Join-Path $ScriptDir $NewISOName

    # Construct bootdata parameter for oscdimg
    # Format: 2#p0,e,b<bios_boot>#pEF,e,b<efi_boot>
    # NO quotes around the paths - oscdimg will add them as needed
    $bootFile = Join-Path $ExtractISO "boot\etfsboot.com"
    $efiFile = Join-Path $ExtractISO "efi\microsoft\boot\efisys.bin"
    $BootData = "2#p0,e,b$bootFile#pEF,e,b$efiFile"
    
    # Volume label is derived from output filename (without .iso extension)
    $volumeLabel = [System.IO.Path]::GetFileNameWithoutExtension($NewISOName)
    
    # Build oscdimg command arguments
    $oscdimgArgs = @(
        "-m"                    # Ignore maximum image size limit
        "-o"                    # Optimize storage by encoding duplicate files only once
        "-u2"                   # Produce UDF file system in addition to ISO9660
        "-udfver102"            # UDF version 1.02
        "-l$volumeLabel"        # Volume label
        "-bootdata:$BootData"   # Boot configuration for BIOS and UEFI
        $ExtractISO             # Source directory
        $NewISOPath             # Output ISO file
    )
    
    Write-Progress-Status -Activity "VirtIO Integration" -Status "Running oscdimg..." -PercentComplete 85
    Write-Host "oscdimg command:" -ForegroundColor Gray
    Write-Host "  $oscdimgPath $($oscdimgArgs -join ' ')" -ForegroundColor DarkGray
    
    # Execute oscdimg and capture output
    $oscdimgOutput = & $oscdimgPath $oscdimgArgs 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        # ====================================================================
        # SUCCESS!
        # ====================================================================
        
        Write-Progress-Status -Activity "VirtIO Integration" -Status "Complete!" -PercentComplete 100
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "Success!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "New ISO: $NewISOPath" -ForegroundColor Cyan
        Write-Host "`nFeatures integrated:" -ForegroundColor Yellow
        Write-Host "  ✓ VirtIO drivers injected into install.wim (post-install)" -ForegroundColor White
        Write-Host "  ✓ VirtIO drivers injected into boot.wim (installer phase)" -ForegroundColor White
        if ($includeGuestTools) {
            Write-Host "  ✓ Guest tools auto-install via SetupComplete.cmd" -ForegroundColor White
            Write-Host "    (Runs after OOBE, installs QEMU Agent + SPICE components)" -ForegroundColor Gray
        }
        Write-Host "  ✓ Full hands-off VirtIO deployment - no IDE workaround needed!" -ForegroundColor White
        Write-Host "  ✓ Volume label: $volumeLabel" -ForegroundColor White
        
        Write-Host "`nCleaning up..." -ForegroundColor Yellow
        Remove-Item -Path $WorkDir -Recurse -Force
    } else {
        Write-Error "oscdimg.exe failed with exit code $LASTEXITCODE"
        Write-Host "oscdimg output:" -ForegroundColor Red
        $oscdimgOutput | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        $script:CleanupOnExit = $false
        throw "oscdimg failed"
    }

} catch {
    Write-Error "Script failed: $_"
    if (-not $script:CleanupOnExit) {
        Write-Host "`nWork directory preserved for debugging: $WorkDir" -ForegroundColor Yellow
    }
    throw
} finally {
    Write-Progress -Activity "VirtIO Integration" -Completed
    
    # Cleanup: Ensure all WIM files are unmounted and ISOs are dismounted
    if ($MountWIM -and (Test-Path $MountWIM)) {
        try {
            Dismount-WindowsImage -Path $MountWIM -Discard -ErrorAction SilentlyContinue
        } catch {}
    }
    
    if ($MountBootWIM -and (Test-Path $MountBootWIM)) {
        try {
            Dismount-WindowsImage -Path $MountBootWIM -Discard -ErrorAction SilentlyContinue
        } catch {}
    }
    
    if ($VirtIOMount) {
        Dismount-DiskImage -ImagePath $VirtIOISO -ErrorAction SilentlyContinue | Out-Null
    }
    
    if ($MicroWinMount) {
        Dismount-DiskImage -ImagePath $MicroWinISO -ErrorAction SilentlyContinue | Out-Null
    }
    
    # Only clean up working directory if script completed successfully
    if ($script:CleanupOnExit -and $WorkDir -and (Test-Path $WorkDir)) {
        Remove-Item -Path $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host "`nDone!" -ForegroundColor Green
}
