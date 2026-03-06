# =====================================================
# ULTIMATE Windows USB Creator (FAT32 Edition)
# FAT32 ONLY - UEFI + BIOS - WIM Auto Split
# Supports install.wim + install.esd
# Works with Windows 10 / 11 / Server / Tiny11
# =====================================================

param(
    [string]$ISOPath
)

$ErrorActionPreference = "Stop"

# -----------------------------------------------------
# Auto Elevate to Administrator
# -----------------------------------------------------
$currUser  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currUser)
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$ISOPath`"" -Verb RunAs
    exit
}

# -----------------------------------------------------
# Ask for ISO if not provided
# -----------------------------------------------------
if (-not $ISOPath) {
    Write-Host ""
    $ISOPath = Read-Host "Enter Windows ISO path (or drag the ISO file here)"
}

$ISOPath = $ISOPath.Trim('"').Trim("'").Trim()

if (-not (Test-Path $ISOPath)) {
    Write-Host "ISO file not found: $ISOPath" -ForegroundColor Red
    pause
    exit
}

$ISOPath = (Resolve-Path $ISOPath).Path

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "        ULTIMATE Windows USB Creator"
Write-Host "             FAT32 Bootable"
Write-Host "===========================================" -ForegroundColor Cyan

# -----------------------------------------------------
# Select USB Drive
# -----------------------------------------------------
$disk = Get-Disk | Where-Object { $_.BusType -eq "USB" } |
        Out-GridView -Title "Select USB Drive (THIS WILL BE ERASED)" -OutputMode Single

if (-not $disk) {
    Write-Host "No USB drive selected. Exiting." -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "Selected Disk : $($disk.Number)"
Write-Host "Size          : $([math]::Round($disk.Size / 1GB, 2)) GB"

$confirm = Read-Host "Type YES to completely erase this USB"
if ($confirm -ne "YES") {
    Write-Host "Cancelled." -ForegroundColor Yellow
    exit
}

# -----------------------------------------------------
# Partition Mode
# -----------------------------------------------------
Write-Host ""
Write-Host "Choose Partition Type:"
Write-Host "  1 - MBR (BIOS + UEFI)"
Write-Host "  2 - GPT (UEFI Only)"

$mode = Read-Host "Enter option (default: 1)"

if ($mode -eq "2") { $partitionStyle = "GPT" }
else               { $partitionStyle = "MBR" }

# -----------------------------------------------------
# Prepare USB
# -----------------------------------------------------
Write-Host ""
Write-Host "Cleaning USB..." -ForegroundColor Cyan

$disk | Clear-Disk -RemoveData -RemoveOEM -Confirm:$false

Write-Host "Initializing disk as $partitionStyle..."
try {
    Initialize-Disk -Number $disk.Number -PartitionStyle $partitionStyle -Confirm:$false
} catch {
    Write-Host "  (Disk already initialized - continuing)" -ForegroundColor DarkGray
}

$disk = Get-Disk -Number $disk.Number
if ($disk.PartitionStyle -ne $partitionStyle) {
    Set-Disk -Number $disk.Number -PartitionStyle $partitionStyle
}

Write-Host "Creating FAT32 partition..."

if ($partitionStyle -eq "GPT") {
    New-Partition -DiskNumber $disk.Number -Size 16MB -GptType "{e3c9e316-0b5c-4db8-817d-f92df00215ae}" 
    $partition = New-Partition -DiskNumber $disk.Number -UseMaximumSize -AssignDriveLetter `
                               -GptType "{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}"
} else {
    $partition = New-Partition -DiskNumber $disk.Number -UseMaximumSize -AssignDriveLetter -IsActive
}

$volume    = Format-Volume -Partition $partition -FileSystem FAT32 -NewFileSystemLabel "WINSETUP" -Confirm:$false
$usbLetter = $volume.DriveLetter

Write-Host "USB mounted as ${usbLetter}:" -ForegroundColor Green

# -----------------------------------------------------
# Mount ISO
# -----------------------------------------------------
Write-Host ""
Write-Host "Mounting ISO..." -ForegroundColor Cyan

Mount-DiskImage -ImagePath $ISOPath 
$isoLetter = (Get-DiskImage -ImagePath $ISOPath | Get-Volume).DriveLetter

if (-not $isoLetter) {
    Write-Host "Failed to determine ISO drive letter." -ForegroundColor Red
    Dismount-DiskImage -ImagePath $ISOPath
    exit
}

Write-Host "ISO mounted as ${isoLetter}:" -ForegroundColor Green

$wimPath = "${isoLetter}:\sources\install.wim"
$esdPath = "${isoLetter}:\sources\install.esd"

# -----------------------------------------------------
# Copy ALL files from ISO
# -----------------------------------------------------
Write-Host ""
Write-Host "Copying all files from ISO..." -ForegroundColor Cyan

robocopy "${isoLetter}:\" "${usbLetter}:\" /E /R:1 /W:1 /XF install.wim install.esd | ForEach-Object {
    Write-Host $_ -ForegroundColor Green
}

if ($LASTEXITCODE -gt 7) {
    Write-Host "Robocopy error (exit code $LASTEXITCODE)." -ForegroundColor Red
    Dismount-DiskImage -ImagePath $ISOPath
    exit
}

# -----------------------------------------------------
# CRITICAL: Strip ReadOnly + System + Hidden attributes
# from every file on the USB.
# Robocopy copies files preserving the ISO's attributes.
# Optical media marks everything ReadOnly, so bcdedit
# gets "Access is denied" when trying to open BCD files.
# We must remove -r -s -h before any bcdedit calls.
# -----------------------------------------------------
Write-Host ""
Write-Host "Stripping read-only attributes from USB files..." -ForegroundColor Cyan
$attribExe = "$env:SystemRoot\System32\attrib.exe"
& $attribExe -r -s -h "${usbLetter}:\*" /s /d
Write-Host "  Done." -ForegroundColor Green

# -----------------------------------------------------
# Handle install image - WIM or ESD
# -----------------------------------------------------
$swmDest = "${usbLetter}:\sources\install.swm"

if (Test-Path $wimPath) {

    Write-Host ""
    Write-Host "install.wim detected - splitting for FAT32..." -ForegroundColor Cyan
    dism /split-image /imagefile:"$wimPath" /swmfile:"$swmDest" /FileSize:4096

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM split-image failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        Dismount-DiskImage -ImagePath $ISOPath
        exit
    }

} elseif (Test-Path $esdPath) {

    Write-Host ""
    Write-Host "install.esd detected - converting to WIM..." -ForegroundColor Cyan

    $tempWim    = Join-Path $env:TEMP "winusb_temp.wim"
    $imageInfo  = & dism /get-imageinfo /imagefile:"$esdPath"
    $indexLines = $imageInfo | Select-String "Index\s*:\s*\d+"
    $indices    = $indexLines | ForEach-Object { ($_ -replace ".*Index\s*:\s*", "").Trim() }

    Write-Host "Found $($indices.Count) image(s) in ESD."

    foreach ($idx in $indices) {
        Write-Host "  Exporting index $idx..."
        dism /export-image /SourceImageFile:"$esdPath" /SourceIndex:$idx `
             /DestinationImageFile:"$tempWim" /Compress:max

        if ($LASTEXITCODE -ne 0) {
            Write-Host "DISM export failed at index $idx." -ForegroundColor Red
            Dismount-DiskImage -ImagePath $ISOPath
            Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
            exit
        }
    }

    Write-Host "Splitting WIM for FAT32..." -ForegroundColor Cyan
    dism /split-image /imagefile:"$tempWim" /swmfile:"$swmDest" /FileSize:4096

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM split-image failed." -ForegroundColor Red
        Dismount-DiskImage -ImagePath $ISOPath
        Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
        exit
    }

    Remove-Item $tempWim -Force -ErrorAction SilentlyContinue

} else {
    Write-Host "WARNING: No install image found." -ForegroundColor Yellow
}

# -----------------------------------------------------
# Ensure bootx64.efi exists for UEFI
# -----------------------------------------------------
Write-Host ""
Write-Host "Checking UEFI boot file..." -ForegroundColor Cyan

New-Item -ItemType Directory -Path "${usbLetter}:\efi\boot" -Force 
$bootx64 = "${usbLetter}:\efi\boot\bootx64.efi"

if (Test-Path $bootx64) {
    Write-Host "  bootx64.efi  : OK" -ForegroundColor Green
} else {
    $candidates = @(
        "${usbLetter}:\efi\microsoft\boot\bootmgfw.efi",
        "${isoLetter}:\efi\microsoft\boot\bootmgfw.efi",
        "${isoLetter}:\efi\boot\bootx64.efi"
    )
    $found = $false
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            Copy-Item $c $bootx64 -Force
            Write-Host "  bootx64.efi  : OK (from $c)" -ForegroundColor Green
            $found = $true
            break
        }
    }
    if (-not $found) {
        Write-Host "  bootx64.efi  : ERROR - not found" -ForegroundColor Red
    }
}

# -----------------------------------------------------
# BCD PATCH - fix device entries so USB boot works
# Now that attrib has been stripped, bcdedit can open
# the BCD files. We find every Windows Boot Loader
# entry by parsing bcdedit output and patch device +
# osdevice to point at boot.wim on the USB.
# -----------------------------------------------------
Write-Host ""
Write-Host "Patching BCD device entries..." -ForegroundColor Cyan

function Patch-BCD {
    param([string]$StorePath)

    if (-not (Test-Path $StorePath)) {
        Write-Host "  Skipping (not found): $StorePath" -ForegroundColor DarkGray
        return
    }

    Write-Host "  Store: $StorePath" -ForegroundColor Cyan

    # Read raw BCD output line by line
    $lines = & bcdedit /store "$StorePath" /enum all 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR reading BCD (exit $LASTEXITCODE):" -ForegroundColor Red
        $lines | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        return
    }

    Write-Host "  --- RAW BCD ---" -ForegroundColor DarkGray
    $lines | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    Write-Host "  --- END BCD ---" -ForegroundColor DarkGray

    # Parse GUIDs of all "Windows Boot Loader" entries
    # bcdedit output format:
    #   Windows Boot Loader
    #   -------------------
    #   identifier    {guid}
    #   device        ...
    $loaderGuids = @()
    $inLoader    = $false

    foreach ($line in $lines) {
        $trimmed = $line.Trim()

        if ($trimmed -match "^Windows Boot Loader") {
            $inLoader = $true
            continue
        }

        if ($inLoader -and $trimmed -eq "") {
            $inLoader = $false
            continue
        }

        if ($inLoader -and $trimmed -match "^identifier\s+(\{[0-9a-fA-F-]+\})") {
            $loaderGuids += $matches[1]
            $inLoader = $false
        }
    }

    if ($loaderGuids.Count -eq 0) {
        Write-Host "  WARNING: No Windows Boot Loader entries found - trying {default}" -ForegroundColor Yellow
        $loaderGuids = @("{default}")
    } else {
        Write-Host "  Found Boot Loader GUIDs: $($loaderGuids -join ', ')" -ForegroundColor Green
    }

    foreach ($guid in $loaderGuids) {
        Write-Host "  Patching $guid ..." -ForegroundColor DarkGray

        $d1 = & bcdedit /store "$StorePath" /set "$guid" device   "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1
        $d2 = & bcdedit /store "$StorePath" /set "$guid" osdevice "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1
        $d3 = & bcdedit /store "$StorePath" /set "$guid" path     "\windows\system32\boot\winload.efi"               2>&1

        Write-Host "    device  : $d1" -ForegroundColor DarkGray
        Write-Host "    osdevice: $d2" -ForegroundColor DarkGray
        Write-Host "    path    : $d3" -ForegroundColor DarkGray
    }

    & bcdedit /store "$StorePath" /set "{ramdiskoptions}" ramdisksdidevice boot           2>&1 
    & bcdedit /store "$StorePath" /set "{ramdiskoptions}" ramdisksdipath "\boot\boot.sdi" 2>&1 

    Write-Host "  --- VERIFY ---" -ForegroundColor Cyan
    $verify = & bcdedit /store "$StorePath" /enum all 2>&1
    $verify | Where-Object { $_ -match "identifier|device|osdevice|path" } | ForEach-Object {
        Write-Host "    $_" -ForegroundColor Green
    }
    Write-Host "  --- END VERIFY ---" -ForegroundColor Cyan
}

Patch-BCD -StorePath "${usbLetter}:\efi\microsoft\boot\BCD"
Patch-BCD -StorePath "${usbLetter}:\boot\BCD"

# -----------------------------------------------------
# BIOS Boot Support (MBR only)
# -----------------------------------------------------
if ($partitionStyle -eq "MBR") {
    Write-Host ""
    Write-Host "Writing BIOS boot sector..." -ForegroundColor Cyan

    $bootsect = "${isoLetter}:\boot\bootsect.exe"

    if (Test-Path $bootsect) {
        & $bootsect /nt60 "${usbLetter}:" 
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  WARNING: bootsect.exe exit code $LASTEXITCODE" -ForegroundColor Yellow
        } else {
            Write-Host "  Boot sector : OK" -ForegroundColor Green
        }
    } else {
        Write-Host "  WARNING: bootsect.exe not found." -ForegroundColor Yellow
    }
}

# -----------------------------------------------------
# Critical file check
# -----------------------------------------------------
Write-Host ""
Write-Host "Critical boot file check:" -ForegroundColor Cyan

$criticalFiles = @(
    "${usbLetter}:\efi\boot\bootx64.efi",
    "${usbLetter}:\efi\microsoft\boot\BCD",
    "${usbLetter}:\boot\BCD",
    "${usbLetter}:\boot\boot.sdi",
    "${usbLetter}:\sources\boot.wim"
)

$allOk = $true
foreach ($f in $criticalFiles) {
    if (Test-Path $f) {
        Write-Host "  [OK]      $f" -ForegroundColor Green
    } else {
        Write-Host "  [MISSING] $f" -ForegroundColor Red
        $allOk = $false
    }
}

if (-not $allOk) {
    Write-Host "WARNING: Missing files will prevent booting." -ForegroundColor Yellow
}

# -----------------------------------------------------
# Cleanup
# -----------------------------------------------------
Write-Host ""
Write-Host "Dismounting ISO..." -ForegroundColor Cyan
Dismount-DiskImage -ImagePath $ISOPath 

Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host "   Bootable FAT32 Windows USB Created!"
Write-Host ""
Write-Host "   USB Drive : ${usbLetter}:"
Write-Host "   Mode      : $partitionStyle"
if ($partitionStyle -eq "GPT") {
    Write-Host "   Supports  : UEFI"
} else {
    Write-Host "   Supports  : UEFI + BIOS"
}
Write-Host "===========================================" -ForegroundColor Green
Write-Host ""

pause
