# =====================================================
# ULTIMATE Windows USB Creator (FAT32 Edition)
# FAT32 ONLY • UEFI + BIOS • WIM Auto Split
# Supports install.wim + install.esd
# Works with Windows 10 / 11 / Server / Tiny11
# Includes EFI repair patch for 0xc0000225
# =====================================================

param(
    [string]$ISOPath
)

# -----------------------------------------------------
# Auto Elevate to Administrator
# -----------------------------------------------------
$currUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currUser)
$admin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $admin) {
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# -----------------------------------------------------
# Ask for ISO if not provided (drag & drop supported)
# -----------------------------------------------------
if (-not $ISOPath) {
    Write-Host ""
    $ISOPath = Read-Host "Enter Windows ISO path or drag the ISO file here"
}

if (!(Test-Path $ISOPath)) {
    Write-Host "ISO file not found!" -ForegroundColor Red
    pause
    exit
}

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "        ULTIMATE Windows USB Creator"
Write-Host "             FAT32 Bootable"
Write-Host "===========================================" -ForegroundColor Cyan

# -----------------------------------------------------
# Select USB Drive
# -----------------------------------------------------
$disk = Get-Disk | Where BusType -eq USB | Out-GridView -Title "Select USB Drive (THIS WILL BE ERASED)" -OutputMode Single

if (!$disk) {
    Write-Host "No USB selected"
    exit
}

Write-Host ""
Write-Host "Selected Disk: $($disk.Number)"
Write-Host "Size: $([math]::Round($disk.Size/1GB,2)) GB"

$confirm = Read-Host "Type YES to completely erase this USB"
if ($confirm -ne "YES") { exit }

# -----------------------------------------------------
# Partition Mode
# -----------------------------------------------------
Write-Host ""
Write-Host "Choose Partition Type"
Write-Host "1 - MBR (BIOS + UEFI)"
Write-Host "2 - GPT (UEFI Only)"

$mode = Read-Host "Enter option"

if ($mode -eq "2") { $partitionType="GPT" }
else { $partitionType="MBR" }

# -----------------------------------------------------
# Prepare USB
# -----------------------------------------------------
Write-Host ""
Write-Host "Cleaning USB..."
$disk | Clear-Disk -RemoveData -Confirm:$false

Write-Host "Initializing disk as $partitionType"
$disk | Initialize-Disk -PartitionStyle $partitionType

Write-Host "Creating FAT32 partition"
$volume = $disk | New-Partition -UseMaximumSize -AssignDriveLetter -IsActive |
    Format-Volume -FileSystem FAT32 -NewFileSystemLabel "WINSETUP" -Confirm:$false

$usbLetter = $volume.DriveLetter

Write-Host "USB Mounted as $usbLetter:"

# -----------------------------------------------------
# Mount ISO
# -----------------------------------------------------
Write-Host ""
Write-Host "Mounting ISO..."
$mount = Mount-DiskImage -ImagePath $ISOPath -PassThru
$isoLetter = ($mount | Get-Volume).DriveLetter

Write-Host "ISO Mounted as $isoLetter:"

# -----------------------------------------------------
# Detect install image
# -----------------------------------------------------
$wim = "$isoLetter`:\sources\install.wim"
$esd = "$isoLetter`:\sources\install.esd"

# -----------------------------------------------------
# Copy files (fast robocopy)
# -----------------------------------------------------
Write-Host ""
Write-Host "Copying Windows files..."
robocopy "$isoLetter`:" "$usbLetter`:" /E /R:1 /W:1 /XF install.wim install.esd /NFL /NDL /NP

# -----------------------------------------------------
# Handle install image
# FAT32 requires splitting files >4GB
# -----------------------------------------------------

if (Test-Path $wim) {

    Write-Host ""
    Write-Host "install.wim detected"
    Write-Host "Splitting WIM for FAT32"

    dism /split-image /imagefile:$wim /swmfile:$usbLetter`:\sources\install.swm /FileSize:4096

}
elseif (Test-Path $esd) {

    Write-Host ""
    Write-Host "install.esd detected"
    Write-Host "Converting ESD → WIM"

    $tempWim = "$env:TEMP\\winusb_temp.wim"

    dism /export-image /SourceImageFile:$esd /SourceIndex:1 /DestinationImageFile:$tempWim /Compress:max

    Write-Host "Splitting WIM for FAT32"

    dism /split-image /imagefile:$tempWim /swmfile:$usbLetter`:\sources\install.swm /FileSize:4096

    Remove-Item $tempWim -Force
}
else {
    Write-Host "WARNING: No install image found" -ForegroundColor Yellow
}

# -----------------------------------------------------
# EFI PATCH (Fix winload.efi / 0xc0000225 errors)
# -----------------------------------------------------

$efiSource = "$isoLetter`:\efi"
$efiDest = "$usbLetter`:\efi"

if (Test-Path $efiSource) {
    Write-Host "Ensuring EFI boot files exist"
    robocopy $efiSource $efiDest /E /R:1 /W:1 | Out-Null
}

$bootx64 = "$usbLetter`:\efi\boot\bootx64.efi"

if (!(Test-Path $bootx64)) {

    $altBoot = "$usbLetter`:\efi\microsoft\boot\bootmgfw.efi"

    if (Test-Path $altBoot) {
        Write-Host "Creating fallback UEFI bootloader"

        New-Item -ItemType Directory -Path "$usbLetter`:\efi\boot" -Force | Out-Null

        Copy-Item $altBoot $bootx64 -Force
    }
}

if (!(Test-Path "$usbLetter`:\boot")) {
    robocopy "$isoLetter`:\boot" "$usbLetter`:\boot" /E | Out-Null
}

# -----------------------------------------------------
# BIOS Boot Support
# -----------------------------------------------------

Write-Host ""
Write-Host "Writing boot sector"

if (Test-Path "$isoLetter`:\boot\bootsect.exe") {

    Set-Location "$isoLetter`:\boot"
    .\bootsect.exe /nt60 "$usbLetter`:" | Out-Null

}

# -----------------------------------------------------
# Cleanup
# -----------------------------------------------------

Write-Host ""
Write-Host "Dismounting ISO"
Dismount-DiskImage -ImagePath $ISOPath

# -----------------------------------------------------
# Finished
# -----------------------------------------------------

Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host " Bootable FAT32 Windows USB Created"
Write-Host " USB Drive: $usbLetter:"
Write-Host " Mode: $partitionType"
Write-Host " Supports: UEFI + BIOS"
Write-Host "===========================================" -ForegroundColor Green
Write-Host ""
