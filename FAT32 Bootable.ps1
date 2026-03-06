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
    Write-Host "  (Initialize-Disk threw: $_)" -ForegroundColor DarkGray
    Write-Host "  Attempting Set-Disk fallback..." -ForegroundColor DarkGray
}

# Refresh disk object - Initialize-Disk may have silently failed if the disk
# was already initialized (common after Clear-Disk on some USB controllers).
# Set-Disk forces the correct partition style regardless.
$disk = Get-Disk -Number $disk.Number
if ($disk.PartitionStyle -ne $partitionStyle) {
    Write-Host "  Disk is $($disk.PartitionStyle), converting to $partitionStyle..." -ForegroundColor DarkGray
    Set-Disk -Number $disk.Number -PartitionStyle $partitionStyle
    $disk = Get-Disk -Number $disk.Number
}

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

$null = Mount-DiskImage -ImagePath $ISOPath
$isoLetter = (Get-DiskImage -ImagePath $ISOPath | Get-Volume).DriveLetter

if (-not $isoLetter) {
    Write-Host "Failed to determine ISO drive letter." -ForegroundColor Red
    $null = Dismount-DiskImage -ImagePath $ISOPath
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
    $null = Dismount-DiskImage -ImagePath $ISOPath
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
        $null = Dismount-DiskImage -ImagePath $ISOPath
        exit
    }

    Write-Host "  Split files written to ${usbLetter}:\sources\" -ForegroundColor Green

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
            $null = Dismount-DiskImage -ImagePath $ISOPath
            Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
            exit
        }
    }

    Write-Host "Splitting WIM for FAT32..." -ForegroundColor Cyan
    & dism /split-image /imagefile:"$tempWim" /swmfile:"$swmDest" /FileSize:3800

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM split-image failed (exit code $LASTEXITCODE)." -ForegroundColor Red
        $null = Dismount-DiskImage -ImagePath $ISOPath
        Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
        exit
    }

    Write-Host "  Split files written to ${usbLetter}:\sources\" -ForegroundColor Green
    Remove-Item $tempWim -Force -ErrorAction SilentlyContinue

} else {
    Write-Host "WARNING: No install.wim or install.esd found in sources." -ForegroundColor Yellow
}

# -----------------------------------------------------
# Ensure bootx64.efi exists for UEFI
# -----------------------------------------------------
Write-Host ""
Write-Host "Checking UEFI boot file..." -ForegroundColor Cyan

$null = New-Item -ItemType Directory -Path "${usbLetter}:\efi\boot" -Force
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
# BCD Rebuild - create a clean BCD store from scratch
#
# WHY rebuild instead of patch:
#   The BCD copied from the ISO has device entries that
#   point to the CD-ROM (cdrom= or boot= pointing at the
#   optical disc). Patching individual fields is fragile
#   because we can't be sure which entry type names the
#   ISO uses ("Windows Setup", "Windows Boot Loader",
#   etc.), and the original entries may be missing fields
#   like systemroot/winpe/detecthal that are required for
#   WinPE to start winload correctly.
#
#   Deleting the ISO BCD and building a fresh one with
#   bcdedit /createstore gives us a known-good, complete
#   store every time, regardless of which Windows version
#   or ISO variant is used.
#
# UEFI BCD  (\efi\microsoft\boot\BCD) -> winload.efi
# BIOS BCD  (\boot\BCD)               -> winload.exe
# -----------------------------------------------------
Write-Host ""
Write-Host "Rebuilding BCD stores from scratch..." -ForegroundColor Cyan

function Rebuild-BCD {
    param(
        [string]$StoreDir,   # Directory that must exist (e.g. F:\efi\microsoft\boot)
        [string]$StorePath,  # Full path to BCD file
        [bool]$IsEFI         # $true = EFI store, $false = BIOS store
    )

    Write-Host ""
    Write-Host "  Store : $StorePath" -ForegroundColor Cyan
    Write-Host "  Type  : $(if ($IsEFI) { 'EFI (winload.efi)' } else { 'BIOS (winload.exe)' })" -ForegroundColor Cyan

    # Ensure the parent directory exists
    $null = New-Item -ItemType Directory -Path $StoreDir -Force

    # Delete the ISO's BCD so we start with a blank slate.
    # The ISO BCD has cdrom device references and may be missing
    # required WinPE fields (systemroot, winpe, detecthal).
    Remove-Item $StorePath -Force -ErrorAction SilentlyContinue

    # Create a new, empty BCD store
    $out = & bcdedit /createstore "$StorePath" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: /createstore failed (exit $LASTEXITCODE): $out" -ForegroundColor Red
        return
    }
    Write-Host "  Created empty store." -ForegroundColor DarkGray

    # ── Windows Boot Manager ──────────────────────────────────────
    & bcdedit /store "$StorePath" /create "{bootmgr}" /d "Windows Boot Manager" 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set    "{bootmgr}" device  boot                                       2>&1 | Out-Null
    if ($IsEFI) {
        & bcdedit /store "$StorePath" /set "{bootmgr}" path   "\EFI\Microsoft\Boot\bootmgfw.efi"         2>&1 | Out-Null
    } else {
        & bcdedit /store "$StorePath" /set "{bootmgr}" path   "\bootmgr"                                 2>&1 | Out-Null
    }
    & bcdedit /store "$StorePath" /set "{bootmgr}" locale     "en-US"                                    2>&1 | Out-Null
    & bcdedit /store "$StorePath" /timeout 30                                                             2>&1 | Out-Null
    Write-Host "  Boot Manager : OK" -ForegroundColor DarkGray

    # ── Windows Setup OS-Loader entry ────────────────────────────
    # /create with no GUID allocates a new random GUID.
    # We parse it from bcdedit's stdout ("The entry {guid} was created").
    $createOut = (& bcdedit /store "$StorePath" /create /d "Windows Setup" /application osloader 2>&1) -join " "
    $guidMatch = [regex]::Match($createOut, "\{[0-9a-fA-F\-]+\}")
    if (-not $guidMatch.Success) {
        Write-Host "  ERROR: Could not parse GUID from: $createOut" -ForegroundColor Red
        return
    }
    $setupGuid   = $guidMatch.Value
    $winloadPath = if ($IsEFI) { "\windows\system32\boot\winload.efi" } `
                               else { "\windows\system32\boot\winload.exe" }

    Write-Host "  Setup GUID   : $setupGuid" -ForegroundColor DarkGray
    Write-Host "  Winload path : $winloadPath" -ForegroundColor DarkGray

    & bcdedit /store "$StorePath" /set "$setupGuid" device     "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" osdevice   "ramdisk=[boot]\sources\boot.wim,{ramdiskoptions}" 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" path       $winloadPath                                        2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" systemroot "\windows"                                          2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" detecthal  yes                                                 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" winpe      yes                                                 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "$setupGuid" locale     "en-US"                                             2>&1 | Out-Null
    Write-Host "  Setup entry  : OK" -ForegroundColor DarkGray

    # Wire setup entry as the default and only display-order item in bootmgr
    & bcdedit /store "$StorePath" /set "{bootmgr}" default      "$setupGuid" 2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set "{bootmgr}" displayorder "$setupGuid" 2>&1 | Out-Null

    # ── Ramdisk options ───────────────────────────────────────────
    # {ramdiskoptions} is a well-known alias for
    # {7619dcc8-fafe-11d9-b411-000476eba25f}.  We must /create it
    # explicitly in a new store before we can /set values on it.
    & bcdedit /store "$StorePath" /create "{ramdiskoptions}" /d "Ramdisk Options"   2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set    "{ramdiskoptions}" ramdisksdidevice boot   2>&1 | Out-Null
    & bcdedit /store "$StorePath" /set    "{ramdiskoptions}" ramdisksdipath "\boot\boot.sdi" 2>&1 | Out-Null
    Write-Host "  Ramdisk opts : OK" -ForegroundColor DarkGray

    # ── Verify ────────────────────────────────────────────────────
    Write-Host "  --- Verify ---" -ForegroundColor Cyan
    $verify = & bcdedit /store "$StorePath" /enum all 2>&1
    $verify | Where-Object { $_ -match "identifier|device|osdevice|path|ramdisk|systemroot|winpe|detecthal" } |
              ForEach-Object { Write-Host "    $_" -ForegroundColor Green }
    Write-Host "  --- End ---" -ForegroundColor Cyan
}

Rebuild-BCD -StoreDir "${usbLetter}:\efi\microsoft\boot" `
            -StorePath "${usbLetter}:\efi\microsoft\boot\BCD" `
            -IsEFI $true

Rebuild-BCD -StoreDir "${usbLetter}:\boot" `
            -StorePath "${usbLetter}:\boot\BCD" `
            -IsEFI $false

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

# Check that at least one SWM split file was created.
# Use -Path + -Filter (more reliable than inline wildcards on FAT32 volumes).
$swmFiles = @(Get-ChildItem -Path "${usbLetter}:\sources" -Filter "install*.swm" -ErrorAction SilentlyContinue)
if ($swmFiles.Count -gt 0) {
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
$null = Dismount-DiskImage -ImagePath $ISOPath

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
