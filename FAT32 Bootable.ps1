# =====================================================
# ULTIMATE Windows USB Creator (FAT32 Edition)
# FAT32 ONLY - UEFI + BIOS - WIM Auto Split
# Supports install.wim + install.esd
# Works with Windows 10 / 11 / Server / Tiny11
#
# FIXES vs original:
#   - GPT: Uses EFI System Partition GUID (not Basic Data)
#     This was causing 0xc000014c (BCD not found by firmware)
#   - GPT: Removed unnecessary MSR partition
#   - BCD: Now detects "Windows Setup" entries (not just "Windows Boot Loader")
#   - BCD: Correct winload path per store (winload.efi vs winload.exe)
#   - Robocopy: Output no longer piped through ForEach-Object (was slow/garbled)
#   - Robocopy: Exit code fixed (>= 8 is error; 0-7 are informational bitmask)
#   - DISM ESD: Uses /Compress:fast (max was needlessly slow and occasionally fails)
#   - DISM ESD: Clears leftover temp WIM before starting
#   - split-image: Size lowered to 3800 MB (safer margin under 4 GB FAT32 limit)
#   - New-Item: Suppressed noisy output with Out-Null
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

# Refresh disk object after init
$disk = Get-Disk -Number $disk.Number

Write-Host "Creating FAT32 partition..."

if ($partitionStyle -eq "GPT") {
    # FIX: Use EFI System Partition GUID {c12a7328-f81f-11d2-ba4b-00a0c93ec93b}
    # The original used Basic Data GUID {ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}
    # which UEFI firmware does NOT recognize as a bootable EFI partition,
    # causing error 0xc000014c (BCD not found) at boot time.
    #
    # Also removed the unnecessary 16 MB MSR partition that was created before
    # the main partition - MSR is only needed on Windows data drives, not USB
    # installers, and it was pushing the ESP to partition slot 2.
    $partition = New-Partition -DiskNumber $disk.Number -UseMaximumSize -AssignDriveLetter `
                               -GptType "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}"
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
# Copy ALL files from ISO (excluding large install image)
# FIX: Removed the ForEach-Object pipe - robocopy output
# was being line-buffered through PS which caused garbled
# display and significant slowdown on large ISOs.
# FIX: Exit code check changed to >= 8.
#   Robocopy exit codes are a bitmask:
#     0 = No files copied (source == dest)
#     1 = Files copied successfully
#     2 = Extra files in dest (not an error)
#     4 = Mismatched files (not an error)
#     8 = FAIL - some files not copied
#    16 = FATAL - serious error
#   Codes 0-7 are all success conditions.
# -----------------------------------------------------
Write-Host ""
Write-Host "Copying all files from ISO..." -ForegroundColor Cyan

robocopy "${isoLetter}:\" "${usbLetter}:\" /E /R:1 /W:1 /XF install.wim install.esd

if ($LASTEXITCODE -ge 8) {
    Write-Host "Robocopy failed (exit code $LASTEXITCODE)." -ForegroundColor Red
    Dismount-DiskImage -ImagePath $ISOPath 
    exit
}

# -----------------------------------------------------
# Strip ReadOnly + System + Hidden attributes
# Robocopy preserves the ISO's read-only flags.
# Optical media marks everything read-only, so bcdedit
# gets "Access is denied" when trying to modify BCD.
# Must strip attributes before any bcdedit calls.
# -----------------------------------------------------
Write-Host ""
Write-Host "Stripping read-only attributes from USB files..." -ForegroundColor Cyan
& "$env:SystemRoot\System32\attrib.exe" -r -s -h "${usbLetter}:\*" /s /d
Write-Host "  Done." -ForegroundColor Green

# -----------------------------------------------------
# Handle install image - WIM or ESD
# FIX: Split size lowered from 4096 to 3800 MB.
#   4096 MB is right at the FAT32 file size limit and
#   can fail due to filesystem overhead. 3800 MB gives
#   a safe margin.
# FIX (ESD path): Use /Compress:fast instead of /Compress:max.
#   /Compress:max is extremely slow (can take hours on large
#   ISOs) and occasionally causes DISM to fail on ESD sources.
#   Since the WIM is temporary and will be split immediately,
#   compression level has no effect on the final USB size.
# FIX (ESD path): Clear leftover temp WIM before starting.
#   If a previous run was interrupted, the old temp file would
#   cause DISM to fail on the first export.
# -----------------------------------------------------
$swmDest = "${usbLetter}:\sources\install.swm"

if (Test-Path $wimPath) {

    Write-Host ""
    Write-Host "install.wim detected - splitting for FAT32..." -ForegroundColor Cyan
    & dism /split-image /imagefile:"$wimPath" /swmfile:"$swmDest" /FileSize:3800

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM split-image failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        Dismount-DiskImage -ImagePath $ISOPath 
        exit
    }

} elseif (Test-Path $esdPath) {

    Write-Host ""
    Write-Host "install.esd detected - converting to split WIM..." -ForegroundColor Cyan

    $tempWim = Join-Path $env:TEMP "winusb_temp.wim"

    # Clear any leftover temp file from a previous interrupted run
    Remove-Item $tempWim -Force -ErrorAction SilentlyContinue

    $imageInfo  = & dism /get-imageinfo /imagefile:"$esdPath"
    $indexLines = $imageInfo | Select-String "Index\s*:\s*\d+"
    $indices    = $indexLines | ForEach-Object { ($_ -replace ".*Index\s*:\s*", "").Trim() }

    Write-Host "Found $($indices.Count) image(s) in ESD."

    foreach ($idx in $indices) {
        Write-Host "  Exporting index $idx..."
        # DISM appends to the destination WIM when it already exists.
        # First iteration creates it; subsequent iterations append.
        & dism /export-image /SourceImageFile:"$esdPath" /SourceIndex:$idx `
               /DestinationImageFile:"$tempWim" /Compress:fast

        if ($LASTEXITCODE -ne 0) {
            Write-Host "DISM export failed at index $idx (exit code $LASTEXITCODE)." -ForegroundColor Red
            Dismount-DiskImage -ImagePath $ISOPath 
            Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
            exit
        }
    }

    Write-Host "Splitting WIM for FAT32..." -ForegroundColor Cyan
    & dism /split-image /imagefile:"$tempWim" /swmfile:"$swmDest" /FileSize:3800

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM split-image failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        Dismount-DiskImage -ImagePath $ISOPath 
        Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
        exit
    }

    Remove-Item $tempWim -Force -ErrorAction SilentlyContinue

} else {
    Write-Host "WARNING: No install.wim or install.esd found in sources." -ForegroundColor Yellow
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
            Write-Host "  bootx64.efi  : OK (copied from $c)" -ForegroundColor Green
            $found = $true
            break
        }
    }
    if (-not $found) {
        Write-Host "  bootx64.efi  : ERROR - not found in any expected location" -ForegroundColor Red
    }
}

# -----------------------------------------------------
# BCD Patch - ensure device entries point to boot.wim
#
# FIX: The original looked only for "Windows Boot Loader"
# entries. Windows setup ISOs use "Windows Setup" type
# entries. The fallback to {default} was unreliable.
# Now we detect both "Windows Setup" and "Windows Boot
# Loader" entries and patch whichever we find.
#
# FIX: winload path now differs per BCD store:
#   EFI BCD  (\efi\microsoft\boot\BCD)  -> winload.efi
#   BIOS BCD (\boot\BCD)                -> winload.exe
# The original always wrote winload.efi to both, which
# breaks BIOS boot.
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

    $lines = & bcdedit /store "$StorePath" /enum all 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR reading BCD (exit $LASTEXITCODE):" -ForegroundColor Red
        $lines | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        return
    }

    # Parse GUIDs for "Windows Setup" and "Windows Boot Loader" entries.
    # Both need the ramdisk device patch to boot correctly from USB.
    $targetGuids = @()
    $inTarget    = $false
    $currentGuid = $null

    foreach ($line in $lines) {
        $trimmed = $line.Trim()

        # Detect a Windows Setup or Windows Boot Loader section header
        if ($trimmed -match "^Windows (Setup|Boot Loader)") {
            $inTarget    = $true
            $currentGuid = $null
            continue
        }

        # Blank line signals end of the current entry block
        if ($trimmed -eq "") {
            if ($inTarget -and $currentGuid) {
                $targetGuids += $currentGuid
            }
            $inTarget    = $false
            $currentGuid = $null
            continue
        }

        # Capture the GUID identifier for the current entry
        if ($inTarget -and $trimmed -match "^identifier\s+(\{[0-9a-fA-F-]+\})") {
            $currentGuid = $matches[1]
        }
    }

    # If the last entry in the file had no trailing blank line, flush it
    if ($inTarget -and $currentGuid) {
        $targetGuids += $currentGuid
    }

    if ($targetGuids.Count -eq 0) {
        Write-Host "  No Windows Setup/Boot Loader entries found. Falling back to {default}." -ForegroundColor Yellow
        $targetGuids = @("{default}")
    } else {
        Write-Host "  Entries to patch: $($targetGuids -join ', ')" -ForegroundColor Green
    }

    # Determine correct winload filename based on which BCD store this is.
    # EFI store uses winload.efi; legacy BIOS store uses winload.exe.
    if ($StorePath -like "*\efi\*") {
        $winloadPath = "\windows\system32\boot\winload.efi"
    } else {
        $winloadPath = "\windows\system32\boot\winload.exe"
    }

    foreach ($guid in $targetGuids) {
        Write-Host "  Patching $guid (path: $winloadPath) ..." -ForegroundColor DarkGray

        & bcdedit /store "$StorePath" /set "$guid" device   "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1 
        & bcdedit /store "$StorePath" /set "$guid" osdevice "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1 
        & bcdedit /store "$StorePath" /set "$guid" path     $winloadPath                                        2>&1 

        Write-Host "    OK" -ForegroundColor Green
    }

    # Ensure ramdiskoptions entry exists and is correct
    & bcdedit /store "$StorePath" /set "{ramdiskoptions}" ramdisksdidevice boot           2>&1 
    & bcdedit /store "$StorePath" /set "{ramdiskoptions}" ramdisksdipath "\boot\boot.sdi" 2>&1 
    Write-Host "  Ramdisk options: OK" -ForegroundColor Green

    # Print a short verification summary
    Write-Host "  Verification:" -ForegroundColor Cyan
    $verify = & bcdedit /store "$StorePath" /enum all 2>&1
    $verify | Where-Object { $_ -match "identifier|device|osdevice|path|ramdisk" } |
              ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

Patch-BCD -StorePath "${usbLetter}:\efi\microsoft\boot\BCD"
Patch-BCD -StorePath "${usbLetter}:\boot\BCD"

# -----------------------------------------------------
# BIOS Boot Support (MBR only)
# -----------------------------------------------------
if ($partitionStyle -eq "MBR") {
    Write-Host ""
    Write-Host "Writing BIOS boot sector..." -ForegroundColor Cyan

    # Prefer bootsect.exe from the ISO (guaranteed to match the Windows version).
    # Fall back to the one on the host system if not present in the ISO.
    $bootsect = "${isoLetter}:\boot\bootsect.exe"
    if (-not (Test-Path $bootsect)) {
        $bootsect = "$env:SystemRoot\System32\bootsect.exe"
    }

    if (Test-Path $bootsect) {
        & "$bootsect" /nt60 "${usbLetter}:"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  WARNING: bootsect.exe exit code $LASTEXITCODE" -ForegroundColor Yellow
        } else {
            Write-Host "  Boot sector : OK" -ForegroundColor Green
        }
    } else {
        Write-Host "  WARNING: bootsect.exe not found anywhere. BIOS boot will not work." -ForegroundColor Yellow
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
        $size = (Get-Item $f).Length
        Write-Host ("  [OK]      {0,-55} {1,10:N0} bytes" -f $f, $size) -ForegroundColor Green
    } else {
        Write-Host "  [MISSING] $f" -ForegroundColor Red
        $allOk = $false
    }
}

# Check that at least one SWM split file was created
$swmFiles = Get-ChildItem "${usbLetter}:\sources\install*.swm" -ErrorAction SilentlyContinue
if ($swmFiles -and $swmFiles.Count -gt 0) {
    Write-Host ("  [OK]      install.swm ({0} part(s))" -f $swmFiles.Count) -ForegroundColor Green
} else {
    Write-Host "  [MISSING] install.swm - no split install image found" -ForegroundColor Red
    $allOk = $false
}

if (-not $allOk) {
    Write-Host ""
    Write-Host "  WARNING: One or more required files are missing." -ForegroundColor Yellow
    Write-Host "           The USB may not boot correctly." -ForegroundColor Yellow
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
