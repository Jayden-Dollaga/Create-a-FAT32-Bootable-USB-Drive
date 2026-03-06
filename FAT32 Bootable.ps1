# Configure the path of your ISO file
$iso = "C:\New folder\2019.ISO"

# =========================================================
# Do not edit below this line
# =========================================================

# Clean ! will clear the selected USB stick
$UsbDrive = Get-Disk | Where-Object BusType -eq USB | Out-GridView -Title 'Select USB Drive to Format' -OutputMode Single
$UsbDrive | Clear-Disk -RemoveData -Confirm:$false

# Convert the partition style to GPT
$UsbDrive = Get-Disk | Where-Object BusType -eq 'USB'
if (($UsbDrive).PartitionStyle -eq 'RAW') {
    $UsbDrive | Initialize-Disk -PartitionStyle 'MBR'
} else {
    $UsbDrive | Set-Disk -PartitionStyle 'MBR'
}

# Create a primary partition and format to FAT32
$volume = $UsbDrive | New-Partition -Size 8GB -IsActive -AssignDriveLetter | Format-Volume -FileSystem FAT32
$usbDriveLetter = $volume.DriveLetter

# Mount ISO and get driver letter of the mounted ISO drive
$mountISO = Mount-DiskImage -ImagePath $iso -StorageType ISO -PassThru
$isoDriveLetter = ($mountISO | Get-Volume).DriveLetter

# Copy mounted ISO content to the USB except install.wim
Copy-Item -Path "$isoDriveLetter`:\*" -Destination "$usbDriveLetter`:\" -Recurse -Force -Exclude install.wim

# Split the file install.wim as FAT32 has a maximum single file size limit of 4GB
dism /split-image /imagefile:$isoDriveLetter`:\sources\install.wim /swmfile:$usbDriveLetter`:\sources\install.swm /FileSize:4096

# Make the USB Bootable for BIOS/MBR compatibility
Set-Location -Path "$($isoDriveLetter):\boot"
.\bootsect.exe /nt60 "$($usbDriveLetter):"

# Dismount ISO
Dismount-DiskImage -ImagePath $iso
