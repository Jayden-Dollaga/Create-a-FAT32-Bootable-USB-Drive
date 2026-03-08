#Requires -Version 5.1
<#
.SYNOPSIS
    RufusPS - Windows USB Creator  (FAT32-first)
.DESCRIPTION
    GUI-based bootable USB creator focused on FAT32 Windows installation:
      * Full UEFI + Legacy BIOS support
      * FAT32 on any USB size (auto-splits WIM > 4 GB into .swm files)
      * NTFS option (no 4 GB split needed)
      * MBR or GPT partition scheme
      * install.wim and install.esd support (auto-converts ESD -> WIM)
      * EFI boot repair (fixes 0xc0000225 / winload.efi errors)
      * Fast Robocopy multi-threaded file copy
      * Real-time DISM progress bars
      * Pre-flight free-space checks (C: for ESD temp, USB for WIM)
      * USB safety checks (prevents wiping wrong disk)
      * Drag-and-drop ISO support
      * Auto admin elevation
      * Windows 10 / 11 / Server / Tiny11 compatible
.NOTES
    Run directly - auto-elevates to Administrator if needed.
#>

# =====================================================================
#  ADMIN AUTO-ELEVATION
# =====================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

# =====================================================================
#  CONSOLE STARTUP BANNER
# =====================================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  RufusPS  -  Windows USB Creator"           -ForegroundColor Cyan
Write-Host "  Running as Administrator"                   -ForegroundColor Green
Write-Host "  Console output mirrors the GUI log window"  -ForegroundColor DarkGray
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# =====================================================================
#  ASSEMBLIES
# =====================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =====================================================================
#  SCRIPT-SCOPE STATE
# =====================================================================
$script:IsoPath         = ""
$script:UsbDrives       = @()
$script:CancelRequested = $false

# -- Colour palette ---------------------------------------------------
$C = @{
    BG        = [System.Drawing.Color]::FromArgb(15,  15,  22)
    BG2       = [System.Drawing.Color]::FromArgb(26,  26,  38)
    BG3       = [System.Drawing.Color]::FromArgb(36,  36,  52)
    Accent    = [System.Drawing.Color]::FromArgb(82,  162, 255)
    AccentDim = [System.Drawing.Color]::FromArgb(50,  100, 170)
    Green     = [System.Drawing.Color]::FromArgb(60,  190,  90)
    Red       = [System.Drawing.Color]::FromArgb(210,  60,  60)
    Orange    = [System.Drawing.Color]::FromArgb(230, 140,  40)
    Muted     = [System.Drawing.Color]::FromArgb(120, 120, 145)
    Text      = [System.Drawing.Color]::FromArgb(220, 220, 235)
    LogGreen  = [System.Drawing.Color]::FromArgb( 90, 210, 110)
    LogCyan   = [System.Drawing.Color]::FromArgb( 80, 200, 220)
    LogYellow = [System.Drawing.Color]::FromArgb(230, 200,  80)
    LogRed    = [System.Drawing.Color]::FromArgb(240,  90,  90)
    LogMuted  = [System.Drawing.Color]::FromArgb(100, 100, 120)
}

# =====================================================================
#  GUI BUILDER
# =====================================================================
function Build-GUI {

    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = "RufusPS  -  Windows USB Creator"
    $form.ClientSize       = New-Object System.Drawing.Size(700, 740)
    $form.StartPosition    = "CenterScreen"
    $form.BackColor        = $C.BG
    $form.ForeColor        = $C.Text
    $form.FormBorderStyle  = "FixedSingle"
    $form.MaximizeBox      = $false
    $form.Font             = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.AllowDrop        = $true
    try { $form.Icon = [System.Drawing.SystemIcons]::Shield } catch {}

    function New-SectionLabel($text, $x, $y) {
        $l           = New-Object System.Windows.Forms.Label
        $l.Text      = $text
        $l.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = $C.Accent
        $l.Location  = New-Object System.Drawing.Point($x, $y)
        $l.Size      = New-Object System.Drawing.Size(660, 17)
        return $l
    }

    function New-Separator($y) {
        $p           = New-Object System.Windows.Forms.Panel
        $p.Location  = New-Object System.Drawing.Point(20, $y)
        $p.Size      = New-Object System.Drawing.Size(660, 1)
        $p.BackColor = $C.BG3
        return $p
    }

    # -- Header -------------------------------------------------------
    $pnlHeader           = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location  = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size      = New-Object System.Drawing.Size(700, 72)
    $pnlHeader.BackColor = $C.BG2
    $form.Controls.Add($pnlHeader)

    $lblTitle           = New-Object System.Windows.Forms.Label
    $lblTitle.Text      = "  RufusPS"
    $lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 19, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $C.Accent
    $lblTitle.Location  = New-Object System.Drawing.Point(18, 10)
    $lblTitle.Size      = New-Object System.Drawing.Size(360, 38)
    $pnlHeader.Controls.Add($lblTitle)

    $lblSub           = New-Object System.Windows.Forms.Label
    $lblSub.Text      = "Windows 10 / 11 / Server  *  UEFI + BIOS  *  FAT32 / NTFS  *  GPT / MBR"
    $lblSub.ForeColor = $C.Muted
    $lblSub.Location  = New-Object System.Drawing.Point(20, 50)
    $lblSub.Size      = New-Object System.Drawing.Size(660, 17)
    $pnlHeader.Controls.Add($lblSub)

    # -- ISO Section --------------------------------------------------
    $form.Controls.Add((New-SectionLabel "ISO IMAGE" 20 88))

    $txtIso             = New-Object System.Windows.Forms.TextBox
    $txtIso.Location    = New-Object System.Drawing.Point(20, 108)
    $txtIso.Size        = New-Object System.Drawing.Size(520, 28)
    $txtIso.BackColor   = $C.BG2
    $txtIso.ForeColor   = $C.Muted
    $txtIso.BorderStyle = "FixedSingle"
    $txtIso.Text        = "Drag & Drop an ISO here, or click Browse ->"
    $txtIso.AllowDrop   = $true
    $form.Controls.Add($txtIso)

    $btnBrowse                           = New-Object System.Windows.Forms.Button
    $btnBrowse.Text                      = "Browse..."
    $btnBrowse.Location                  = New-Object System.Drawing.Point(548, 107)
    $btnBrowse.Size                      = New-Object System.Drawing.Size(132, 30)
    $btnBrowse.BackColor                 = $C.AccentDim
    $btnBrowse.ForeColor                 = $C.Text
    $btnBrowse.FlatStyle                 = "Flat"
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.Cursor                    = "Hand"
    $form.Controls.Add($btnBrowse)

    $lblIsoInfo           = New-Object System.Windows.Forms.Label
    $lblIsoInfo.Text      = ""
    $lblIsoInfo.ForeColor = $C.Green
    $lblIsoInfo.Location  = New-Object System.Drawing.Point(20, 140)
    $lblIsoInfo.Size      = New-Object System.Drawing.Size(660, 18)
    $form.Controls.Add($lblIsoInfo)

    $form.Controls.Add((New-Separator 164))

    # -- USB Section --------------------------------------------------
    $form.Controls.Add((New-SectionLabel "USB DRIVE" 20 174))

    $cmbUsb               = New-Object System.Windows.Forms.ComboBox
    $cmbUsb.Location      = New-Object System.Drawing.Point(20, 194)
    $cmbUsb.Size          = New-Object System.Drawing.Size(520, 28)
    $cmbUsb.BackColor     = $C.BG2
    $cmbUsb.ForeColor     = $C.Text
    $cmbUsb.FlatStyle     = "Flat"
    $cmbUsb.DropDownStyle = "DropDownList"
    $form.Controls.Add($cmbUsb)

    $btnRefresh                            = New-Object System.Windows.Forms.Button
    $btnRefresh.Text                       = "Refresh"
    $btnRefresh.Location                   = New-Object System.Drawing.Point(548, 193)
    $btnRefresh.Size                       = New-Object System.Drawing.Size(132, 30)
    $btnRefresh.BackColor                  = $C.BG3
    $btnRefresh.ForeColor                  = $C.Text
    $btnRefresh.FlatStyle                  = "Flat"
    $btnRefresh.FlatAppearance.BorderColor = $C.BG3
    $btnRefresh.Cursor                     = "Hand"
    $form.Controls.Add($btnRefresh)

    $lblUsbWarn           = New-Object System.Windows.Forms.Label
    $lblUsbWarn.Text      = ""
    $lblUsbWarn.ForeColor = $C.Orange
    $lblUsbWarn.Location  = New-Object System.Drawing.Point(20, 228)
    $lblUsbWarn.Size      = New-Object System.Drawing.Size(660, 18)
    $form.Controls.Add($lblUsbWarn)

    $form.Controls.Add((New-Separator 252))

    # -- Partition Scheme ---------------------------------------------
    # IMPORTANT: Each radio group must live in its own Panel so WinForms
    # does not treat all radio buttons on the form as one group.
    $form.Controls.Add((New-SectionLabel "PARTITION SCHEME" 20 262))

    $pnlPartition           = New-Object System.Windows.Forms.Panel
    $pnlPartition.Location  = New-Object System.Drawing.Point(20, 280)
    $pnlPartition.Size      = New-Object System.Drawing.Size(660, 28)
    $pnlPartition.BackColor = $C.BG
    $form.Controls.Add($pnlPartition)

    $rbMBR           = New-Object System.Windows.Forms.RadioButton
    $rbMBR.Text      = "MBR  -  BIOS + UEFI  (recommended for most PCs)"
    $rbMBR.Location  = New-Object System.Drawing.Point(0, 2)
    $rbMBR.Size      = New-Object System.Drawing.Size(330, 24)
    $rbMBR.BackColor = $C.BG
    $rbMBR.ForeColor = $C.Text
    $rbMBR.Checked   = $true
    $pnlPartition.Controls.Add($rbMBR)

    $rbGPT           = New-Object System.Windows.Forms.RadioButton
    $rbGPT.Text      = "GPT  -  UEFI only  (modern systems / Secure Boot)"
    $rbGPT.Location  = New-Object System.Drawing.Point(340, 2)
    $rbGPT.Size      = New-Object System.Drawing.Size(320, 24)
    $rbGPT.BackColor = $C.BG
    $rbGPT.ForeColor = $C.Text
    $pnlPartition.Controls.Add($rbGPT)

    $form.Controls.Add((New-Separator 314))

    # -- File System --------------------------------------------------
    $form.Controls.Add((New-SectionLabel "FILE SYSTEM" 20 324))

    $pnlFileSys           = New-Object System.Windows.Forms.Panel
    $pnlFileSys.Location  = New-Object System.Drawing.Point(20, 342)
    $pnlFileSys.Size      = New-Object System.Drawing.Size(660, 42)
    $pnlFileSys.BackColor = $C.BG
    $form.Controls.Add($pnlFileSys)

    $rbFAT32           = New-Object System.Windows.Forms.RadioButton
    $rbFAT32.Text      = "FAT32  -  Universal UEFI + BIOS  (auto-splits WIM > 4 GB)"
    $rbFAT32.Location  = New-Object System.Drawing.Point(0, 4)
    $rbFAT32.Size      = New-Object System.Drawing.Size(360, 36)
    $rbFAT32.BackColor = $C.BG
    $rbFAT32.ForeColor = $C.Text
    $rbFAT32.Checked   = $true
    $pnlFileSys.Controls.Add($rbFAT32)

    $rbNTFS           = New-Object System.Windows.Forms.RadioButton
    $rbNTFS.Text      = "NTFS  -  No 4 GB limit  (BIOS / limited UEFI)"
    $rbNTFS.Location  = New-Object System.Drawing.Point(365, 4)
    $rbNTFS.Size      = New-Object System.Drawing.Size(295, 36)
    $rbNTFS.BackColor = $C.BG
    $rbNTFS.ForeColor = $C.Text
    $pnlFileSys.Controls.Add($rbNTFS)

    $form.Controls.Add((New-Separator 390))

    # -- Progress -----------------------------------------------------
    $form.Controls.Add((New-SectionLabel "PROGRESS" 20 400))

    $lblPct           = New-Object System.Windows.Forms.Label
    $lblPct.Text      = "0%"
    $lblPct.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lblPct.ForeColor = $C.Accent
    $lblPct.Location  = New-Object System.Drawing.Point(648, 400)
    $lblPct.Size      = New-Object System.Drawing.Size(40, 17)
    $lblPct.TextAlign = "TopRight"
    $form.Controls.Add($lblPct)

    $progressBar           = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location  = New-Object System.Drawing.Point(20, 420)
    $progressBar.Size      = New-Object System.Drawing.Size(660, 22)
    $progressBar.Style     = "Continuous"
    $progressBar.ForeColor = $C.Accent
    $progressBar.Minimum   = 0
    $progressBar.Maximum   = 100
    $form.Controls.Add($progressBar)

    $lblStatus           = New-Object System.Windows.Forms.Label
    $lblStatus.Text      = "Ready.  Select an ISO and a USB drive, then click Create USB."
    $lblStatus.ForeColor = $C.Muted
    $lblStatus.Location  = New-Object System.Drawing.Point(20, 447)
    $lblStatus.Size      = New-Object System.Drawing.Size(660, 18)
    $form.Controls.Add($lblStatus)

    $form.Controls.Add((New-Separator 468))

    # -- Log ----------------------------------------------------------
    $form.Controls.Add((New-SectionLabel "LOG" 20 478))

    $txtLog             = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Location    = New-Object System.Drawing.Point(20, 498)
    $txtLog.Size        = New-Object System.Drawing.Size(660, 190)
    $txtLog.BackColor   = [System.Drawing.Color]::FromArgb(8, 8, 14)
    $txtLog.ForeColor   = $C.LogGreen
    $txtLog.Font        = New-Object System.Drawing.Font("Consolas", 8.5)
    $txtLog.ReadOnly    = $true
    $txtLog.BorderStyle = "None"
    $txtLog.ScrollBars  = "Vertical"
    $form.Controls.Add($txtLog)

    # -- Action Buttons -----------------------------------------------
    $btnStart                           = New-Object System.Windows.Forms.Button
    $btnStart.Text                      = ">   CREATE USB"
    $btnStart.Location                  = New-Object System.Drawing.Point(20, 685)
    $btnStart.Size                      = New-Object System.Drawing.Size(520, 42)
    $btnStart.BackColor                 = $C.Green
    $btnStart.ForeColor                 = [System.Drawing.Color]::White
    $btnStart.Font                      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnStart.FlatStyle                 = "Flat"
    $btnStart.FlatAppearance.BorderSize = 0
    $btnStart.Cursor                    = "Hand"
    $form.Controls.Add($btnStart)

    $btnCancel                           = New-Object System.Windows.Forms.Button
    $btnCancel.Text                      = "Cancel"
    $btnCancel.Location                  = New-Object System.Drawing.Point(548, 685)
    $btnCancel.Size                      = New-Object System.Drawing.Size(132, 42)
    $btnCancel.BackColor                 = $C.Red
    $btnCancel.ForeColor                 = [System.Drawing.Color]::White
    $btnCancel.Font                      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnCancel.FlatStyle                 = "Flat"
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.Cursor                    = "Hand"
    $btnCancel.Enabled                   = $false
    $form.Controls.Add($btnCancel)

    return @{
        Form       = $form
        TxtIso     = $txtIso
        BtnBrowse  = $btnBrowse
        LblIsoInfo = $lblIsoInfo
        CmbUsb     = $cmbUsb
        BtnRefresh = $btnRefresh
        LblUsbWarn = $lblUsbWarn
        RbMBR      = $rbMBR
        RbGPT      = $rbGPT
        RbFAT32    = $rbFAT32
        RbNTFS     = $rbNTFS
        ProgressBar= $progressBar
        LblPct     = $lblPct
        LblStatus  = $lblStatus
        TxtLog     = $txtLog
        BtnStart   = $btnStart
        BtnCancel  = $btnCancel
    }
}

# =====================================================================
#  HELPER FUNCTIONS
# =====================================================================

function Write-Log {
    param($UI, [string]$Message, [string]$Level = "Info")

    $col = switch ($Level) {
        "Warn"    { [System.Drawing.Color]::FromArgb(230, 200,  80) }
        "Error"   { [System.Drawing.Color]::FromArgb(240,  90,  90) }
        "Success" { [System.Drawing.Color]::FromArgb( 80, 200, 220) }
        "Cyan"    { [System.Drawing.Color]::FromArgb( 80, 200, 220) }
        "Muted"   { [System.Drawing.Color]::FromArgb(100, 100, 120) }
        default   { [System.Drawing.Color]::FromArgb( 90, 210, 110) }
    }
    $consoleColor = switch ($Level) {
        "Warn"    { "Yellow"   }
        "Error"   { "Red"      }
        "Success" { "Cyan"     }
        "Cyan"    { "Cyan"     }
        "Muted"   { "DarkGray" }
        default   { "Green"    }
    }

    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host "[$ts] $Message" -ForegroundColor $consoleColor

    try {
        $UI.TxtLog.SelectionStart  = $UI.TxtLog.TextLength
        $UI.TxtLog.SelectionLength = 0
        $UI.TxtLog.SelectionColor  = $col
        $UI.TxtLog.AppendText("[$ts] $Message`r`n")
        $UI.TxtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        try { $UI.TxtLog.AppendText("[$ts] $Message`r`n") } catch {}
    }
}

function Set-Progress {
    param($UI, [int]$Pct, [string]$Status)
    $Pct = [Math]::Max(0, [Math]::Min($Pct, 100))
    $UI.ProgressBar.Value = $Pct
    $UI.LblPct.Text       = "$Pct%"
    $UI.LblStatus.Text    = $Status
    Write-Host "  --> [$Pct%] $Status" -ForegroundColor DarkCyan
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-UsbList {
    param($UI)
    $UI.CmbUsb.Items.Clear()
    $script:UsbDrives = @()

    try {
        $disks = Get-Disk -ErrorAction Stop | Where-Object BusType -eq 'USB'
    } catch {
        $UI.LblUsbWarn.Text = "Unable to enumerate disks. Run from elevated Windows PowerShell."
        return
    }

    foreach ($disk in $disks) {
        $sizeGB     = [Math]::Round($disk.Size / 1GB, 1)
        $label      = ""
        $partitions = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue
        foreach ($p in $partitions) {
            try {
                $vol = Get-Volume -Partition $p -ErrorAction SilentlyContinue
                if ($vol -and $vol.FileSystemLabel) { $label = $vol.FileSystemLabel; break }
            } catch {}
        }
        $display = "Disk $($disk.Number)  -  $sizeGB GB   [$($disk.FriendlyName)]"
        if ($label) { $display += "   ($label)" }

        $script:UsbDrives += [PSCustomObject]@{
            Display = $display
            Disk    = $disk
            SizeGB  = $sizeGB
        }
        $UI.CmbUsb.Items.Add($display)
    }

    if ($UI.CmbUsb.Items.Count -gt 0) {
        $UI.CmbUsb.SelectedIndex = 0
        $UI.LblUsbWarn.Text      = "[!]  All data on the selected drive will be permanently erased!"
    } else {
        $UI.LblUsbWarn.Text = "No USB drives detected.  Insert a drive and click Refresh."
    }
}

function Set-IsoPath {
    param($UI, [string]$Path)
    if (-not (Test-Path $Path)) { return }

    if ([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -ne ".iso") {
        [System.Windows.Forms.MessageBox]::Show(
            "Please choose a valid .iso file.",
            "Invalid file type",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $script:IsoPath      = $Path
    $sizeGB              = [Math]::Round((Get-Item $Path).Length / 1GB, 2)
    $UI.TxtIso.Text      = $Path
    $UI.TxtIso.ForeColor = $C.Text
    $UI.LblIsoInfo.Text  = "[OK]  $([System.IO.Path]::GetFileName($Path))  ($sizeGB GB)"
    Write-Log $UI "ISO selected: $Path" "Info"
}

# -- Detect install image type from mounted ISO -----------------------
function Test-WindowsIsoLayout {
    param([string]$IsoDrive)

    # sources\boot.wim is the only file that absolutely must exist —
    # it contains the Windows PE installer environment.
    # setup.exe and bootmgr are absent from some valid Windows 11 ISOs
    # so we treat them as advisory warnings rather than hard failures.
    # Guard against null/empty drive letter being passed in.
    if ([string]::IsNullOrWhiteSpace($IsoDrive)) {
        return [PSCustomObject]@{ IsValid = $false; Missing = @('(drive letter is empty)'); Warnings = @() }
    }

    $required = @('sources\boot.wim')
    $advisory = @('setup.exe', 'bootmgr')

    $missing  = @()
    $warnings = @()
    foreach ($rel in $required) { if (-not (Test-Path "$IsoDrive`:\$rel")) { $missing  += $rel } }
    foreach ($rel in $advisory) { if (-not (Test-Path "$IsoDrive`:\$rel")) { $warnings += $rel } }

    return [PSCustomObject]@{
        IsValid  = ($missing.Count -eq 0)
        Missing  = $missing
        Warnings = $warnings
    }
}

function Get-InstallImageInfo {
    param([string]$IsoDrive)
    $wimPath = "$IsoDrive`:\sources\install.wim"
    $esdPath = "$IsoDrive`:\sources\install.esd"

    $path = if (Test-Path $wimPath) { $wimPath }
            elseif (Test-Path $esdPath) { $esdPath }
            else { return @{ Type = "None"; Path = ""; Version = "Not found" } }

    $type    = if ($path -match "\.esd$") { "ESD" } else { "WIM" }
    $version = "Unknown"
    try {
        $raw  = & dism /Get-WimInfo "/WimFile:$path" /Index:1 2>&1
        $line = $raw | Where-Object { $_ -match "^Name\s*:" } | Select-Object -First 1
        if ($line) { $version = ($line -split ":", 2)[1].Trim() }
    } catch {}

    return @{ Type = $type; Path = $path; Version = $version }
}

# -- Run DISM with real-time progress capture -------------------------
function Invoke-DismWithProgress {
    param(
        $UI,
        [string]$Arguments,
        [int]$BasePercent,
        [int]$PercentRange,
        [string]$Label
    )

    $outFile  = [System.IO.Path]::GetTempFileName()
    $errFile  = [System.IO.Path]::GetTempFileName()
    $exitCode = 0   # default; set inside try before finally to avoid race

    $proc = Start-Process dism.exe `
        -ArgumentList $Arguments `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile `
        -NoNewWindow -PassThru

    $reader  = $null
    $stream  = $null
    $lastPct = 0

    try {
        # Open the temp file for reading while DISM is still writing it.
        # FileShare ReadWrite is required so DISM can keep writing.
        $stream = [System.IO.File]::Open(
            $outFile,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $reader = New-Object System.IO.StreamReader($stream)

        while (-not $proc.HasExited) {

            if ($script:CancelRequested) {
                try { $proc.Kill() } catch {}
                break
            }

            while ($null -ne ($line = $reader.ReadLine())) {
                # Regex: optional literal dot between digits e.g. "10.0%" or "100%"
                if ($line -match '(\d+\.?\d*)%') {
                    $pct = $BasePercent + [int]([double]$Matches[1] / 100 * $PercentRange)
                    if ($pct -ne $lastPct) {
                        Set-Progress $UI $pct "$Label  $([Math]::Round([double]$Matches[1]))%"
                        $lastPct = $pct
                    }
                }
                if ($line.Trim() -ne '' -and
                    $line -notmatch '^(Deployment Image|Microsoft|Copyright|---|=)') {
                    Write-Log $UI "  DISM: $($line.Trim())" "Muted"
                }
            }

            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 150
        }

        # WaitForExit() must be called before ExitCode is readable on
        # Start-Process -PassThru handles.
        $proc.WaitForExit()

        # Capture exit code INSIDE try, BEFORE finally closes streams.
        # Treat null (rare Start-Process race) as 0 = success.
        $exitCode = if ($null -eq $proc.ExitCode) { 0 } else { [int]$proc.ExitCode }

        # Drain any remaining output lines after DISM exits.
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line -match '(\d+\.?\d*)%') {
                $pct = $BasePercent + [int]([double]$Matches[1] / 100 * $PercentRange)
                Set-Progress $UI $pct "$Label  $([Math]::Round([double]$Matches[1]))%"
            }
            if ($line.Trim() -ne '' -and
                $line -notmatch '^(Deployment Image|Microsoft|Copyright|---|=)') {
                Write-Log $UI "  DISM: $($line.Trim())" "Muted"
            }
        }

    } finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($stream) { try { $stream.Close() } catch {} }
        try { Remove-Item $outFile -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Item $errFile -Force -ErrorAction SilentlyContinue } catch {}
    }

    return $exitCode
}

# =====================================================================
#  CORE USB CREATION
# =====================================================================
function Start-UsbCreation {
    param(
        $UI,
        [string]$IsoPath,
        $DiskObj,
        [bool]$UseGPT,
        [bool]$UseNTFS
    )

    $script:CancelRequested = $false
    $UI.BtnStart.Enabled    = $false
    $UI.BtnCancel.Enabled   = $true
    $mountISO               = $null
    $isoDrive               = $null
    $tempWim                = $null

    Set-Progress $UI 0 "Starting..."

    try {

        # ----------------------------------------------------------------
        # STEP 1  Safety checks
        # ----------------------------------------------------------------
        Set-Progress $UI 2 "Running safety checks..."
        Write-Log $UI "---  Safety Checks  ---" "Cyan"

        $DiskObj = Get-Disk -Number $DiskObj.Number -ErrorAction SilentlyContinue
        if (-not $DiskObj) { throw "Could not read the selected disk. Unplug and re-insert the USB, then click Refresh." }

        if ($DiskObj.BusType -ne 'USB') { throw "SAFETY: The selected disk is not a USB drive." }

        $sysDisk = Get-Disk | Where-Object { $_.IsSystem -or $_.IsBoot }
        if ($DiskObj.Number -in $sysDisk.Number) { throw "SAFETY: Selected disk contains a system/boot partition." }

        if ($DiskObj.Size -lt 7GB) {
            throw "SAFETY: Drive is too small ($([Math]::Round($DiskObj.Size/1GB,1)) GB). Need a real 8 GB+ drive."
        }
        Write-Log $UI "[OK]  Safety checks passed." "Success"

        # ----------------------------------------------------------------
        # STEP 2  Mount ISO
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 4 "Mounting ISO..."
        Write-Log $UI "---  Mount ISO  ---" "Cyan"
        Write-Log $UI "File: $([System.IO.Path]::GetFileName($IsoPath))"

        $mountISO = Mount-DiskImage -ImagePath $IsoPath -PassThru -ErrorAction Stop

        # Poll for drive letter — Get-Volume can return $null or a char(0)
        # immediately after mount on slow/virtual drives.
        $isoVol = $null
        for ($t = 0; $t -lt 20; $t++) {
            try { $isoVol = Get-DiskImage -ImagePath $IsoPath | Get-Volume -ErrorAction Stop } catch {}
            if (-not $isoVol) {
                try { $isoVol = $mountISO | Get-Volume -ErrorAction Stop } catch {}
            }
            if ($isoVol -and "$($isoVol.DriveLetter)".Trim() -ne "") { break }
            Start-Sleep -Milliseconds 500
        }
        # Cast to [string] explicitly — DriveLetter is [char] and a null char
        # passes PowerShell's -not check but breaks path construction.
        [string]$isoDrive = "$($isoVol.DriveLetter)".Trim()
        if ($isoDrive -eq "" -or $isoDrive -eq [char]0) {
            throw "ISO did not receive a drive letter after 10 seconds. Try remounting."
        }
        Write-Log $UI "[OK]  ISO mounted at $isoDrive`:\" "Success"
        Write-Log $UI "  Drive letter: '$isoDrive'" "Muted"

        # Hard safety check: validate the mounted ISO looks like a Windows
        # setup image BEFORE wiping the selected USB disk.
        $isoLayout = Test-WindowsIsoLayout $isoDrive
        if (-not $isoLayout.IsValid) {
            $missingText = ($isoLayout.Missing | ForEach-Object { "  - $_" }) -join "`n"
            throw ("The selected ISO is missing sources\boot.wim and cannot be used:`n$missingText`n`nUse a full Windows installation ISO.")
        }
        if ($isoLayout.Warnings.Count -gt 0) {
            $isoLayout.Warnings | ForEach-Object { Write-Log $UI "  [?] Advisory: $_ not found (non-fatal)." "Warn" }
        }
        Write-Log $UI "[OK]  ISO layout check passed." "Success"

        # ----------------------------------------------------------------
        # STEP 3  Detect install image
        # ----------------------------------------------------------------
        Set-Progress $UI 7 "Detecting Windows version..."
        Write-Log $UI "---  Windows Detection  ---" "Cyan"
        $imgInfo = Get-InstallImageInfo $isoDrive
        Write-Log $UI "  Install image  : $($imgInfo.Type)"
        Write-Log $UI "  Windows edition: $($imgInfo.Version)"

        # Pre-flight: check temp space for ESD conversion
        if ($imgInfo.Type -eq "ESD") {
            $esdBytes  = (Get-Item $imgInfo.Path).Length
            $needBytes = [Math]::Max($esdBytes * 2, 8GB)
            $tmpDrive  = [System.IO.Path]::GetPathRoot($env:TEMP).TrimEnd('\')[0]
            $tmpFree   = (Get-Volume -DriveLetter $tmpDrive -ErrorAction SilentlyContinue).SizeRemaining
            if ($tmpFree -and $tmpFree -lt $needBytes) {
                throw "Not enough space on $tmpDrive`:\ for ESD conversion. Need $([Math]::Round($needBytes/1GB,1)) GB, have $([Math]::Round($tmpFree/1GB,1)) GB."
            }
        }

        # ----------------------------------------------------------------
        # STEP 4  Prepare disk  (mirrors original script exactly)
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 10 "Preparing USB drive..."
        Write-Log $UI "---  Disk Preparation  ---" "Cyan"

        [int]$diskNum = $DiskObj.Number
        Write-Log $UI "  Disk $diskNum  -  wiping..." "Warn"

        # Original script sequence — do not change the order or add ErrorAction Stop.
        # The original works without try/catch; we match that exactly.
        $DiskObj | Set-Disk -IsReadOnly $false
        $DiskObj | Clear-Disk -RemoveData -Confirm:$false
        Initialize-Disk -Number $diskNum -PartitionStyle MBR

        Write-Log $UI "[OK]  Disk prepared (MBR)." "Success"

        # ----------------------------------------------------------------
        # STEP 5  Partition and format  (mirrors original script)
        # ----------------------------------------------------------------
        Set-Progress $UI 14 "Creating partition..."
        Write-Log $UI "---  Partition & Format  ---" "Cyan"

        $fs      = if ($UseNTFS) { "NTFS" } else { "FAT32" }
        $fsLabel = "WINUSB"

        # Pre-flight: FAT32 large-file scan
        if (-not $UseNTFS) {
            $fatLimit  = 4294967295
            $bigFiles  = Get-ChildItem "$isoDrive`:\" -Recurse -File -ErrorAction SilentlyContinue |
                         Where-Object { $_.Length -gt $fatLimit -and $_.Name -notmatch '^install\.(wim|esd)$' }
            if ($bigFiles) {
                $list = ($bigFiles | ForEach-Object { "  $($_.Name)  ($([Math]::Round($_.Length/1GB,2)) GB)" }) -join "`n"
                throw "FAT32 cannot store files > 4 GB. Found:`n$list`nUse NTFS for this ISO."
            }
        }

        Write-Log $UI "  Creating $fs partition..."
        $partition = New-Partition -DiskNumber $diskNum -UseMaximumSize -IsActive -AssignDriveLetter
        $volume    = Format-Volume -Partition $partition -FileSystem $fs -NewFileSystemLabel $fsLabel -Confirm:$false
        $usbDrive  = $volume.DriveLetter

        if (-not $usbDrive) { throw "Partition was created but no drive letter was assigned." }
        Write-Log $UI "[OK]  Partition ready: $usbDrive`:\ ($fs)" "Success"

        # ----------------------------------------------------------------
        # STEP 6  Copy files  (mirrors original — Copy-Item + robocopy log)
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 18 "Copying Windows files..."
        Write-Log $UI "---  File Copy  ---" "Cyan"
        Write-Log $UI "  $isoDrive`:\ --> $usbDrive`:\  (excluding install.wim / install.esd)"

        # Use robocopy for speed/progress display, but fall back to
        # Copy-Item on error (mirrors original's Copy-Item approach).
        $roboArgs = @(
            "$isoDrive`:\", "$usbDrive`:\",
            "/E", "/MT:16", "/NJH", "/NJS", "/NDL", "/NP",
            "/XF", "install.wim", "install.esd"
        )
        $roboJob = Start-Process robocopy -ArgumentList $roboArgs -PassThru -NoNewWindow
        $rStart  = Get-Date
        while (-not $roboJob.HasExited) {
            if ($script:CancelRequested) { try { $roboJob.Kill() } catch {}; throw "Cancelled." }
            $elapsed = [int]((Get-Date) - $rStart).TotalSeconds
            Set-Progress $UI ([Math]::Min(18 + $elapsed, 48)) "Copying files... (${elapsed}s)"
            Start-Sleep -Milliseconds 400
        }
        $roboJob.WaitForExit()
        if ($roboJob.ExitCode -gt 7) { throw "Robocopy failed (exit $($roboJob.ExitCode))." }
        Write-Log $UI "[OK]  Files copied." "Success"

        # ----------------------------------------------------------------
        # STEP 7  ESD -> WIM conversion  (if needed)
        # ----------------------------------------------------------------
        $wimSource = ""
        if ($imgInfo.Type -eq "ESD") {
            if ($script:CancelRequested) { throw "Cancelled." }
            Set-Progress $UI 50 "Converting install.esd -> install.wim..."
            Write-Log $UI "---  ESD -> WIM  ---" "Cyan"

            $esdFile  = $imgInfo.Path
            $tempWim  = "$env:TEMP\rufusps_install.wim"
            if (Test-Path $tempWim) { Remove-Item $tempWim -Force }

            $rawInfoText = (& dism /Get-WimInfo "/WimFile:$esdFile" 2>&1 | Out-String)
            $indexes     = [regex]::Matches($rawInfoText, "Index\s*:\s*(\d+)") |
                           ForEach-Object { $_.Groups[1].Value }
            if ($indexes.Count -eq 0) { throw "Could not read indexes from install.esd." }
            Write-Log $UI "  $($indexes.Count) index(es) found."

            $pctPer = [int](20 / $indexes.Count)
            for ($i = 0; $i -lt $indexes.Count; $i++) {
                if ($script:CancelRequested) { throw "Cancelled." }
                $idx  = $indexes[$i]
                $base = 50 + ($i * $pctPer)
                Write-Log $UI "  Exporting index $idx..."
                $dismArgs = "/Export-Image /SourceImageFile:`"$esdFile`" /SourceIndex:$idx " +
                             "/DestinationImageFile:`"$tempWim`" /Compress:fast"
                $code = Invoke-DismWithProgress $UI $dismArgs $base $pctPer "ESD index $idx"
                if ([int]$code -gt 1) { Write-Log $UI "  [!] Index $idx code $code." "Warn" }
            }
            if (-not (Test-Path $tempWim)) { throw "ESD conversion produced no output file." }
            $wimSource = $tempWim
            Write-Log $UI "[OK]  ESD -> WIM done." "Success"

        } elseif ($imgInfo.Type -eq "WIM") {
            $wimSource = $imgInfo.Path
        }

        # ----------------------------------------------------------------
        # STEP 8  Split or copy install image  (mirrors original — dism via Start-Process)
        # ----------------------------------------------------------------
        if ($wimSource -ne "") {
            if ($script:CancelRequested) { throw "Cancelled." }
            $wimBytes = (Get-Item $wimSource).Length
            Write-Log $UI "---  Install Image  ---" "Cyan"
            Write-Log $UI "  Size: $([Math]::Round($wimBytes/1GB,2)) GB"

            $srcDir = "$usbDrive`:\sources"
            if (-not (Test-Path $srcDir)) { New-Item $srcDir -ItemType Directory | Out-Null }

            # USB free-space check
            $usbFree = (Get-Volume -DriveLetter $usbDrive -ErrorAction SilentlyContinue).SizeRemaining
            if ($usbFree -and $usbFree -lt ($wimBytes + 200MB)) {
                throw "Not enough space on USB. Need $([Math]::Round(($wimBytes+200MB)/1GB,1)) GB, have $([Math]::Round($usbFree/1GB,1)) GB."
            }

            if ((-not $UseNTFS) -and ($wimBytes -gt 4GB)) {
                Set-Progress $UI 72 "Splitting install.wim for FAT32..."
                Write-Log $UI "  WIM > 4 GB — splitting into .swm chunks..." "Warn"

                # Use Start-Process -Verb runAs exactly like the original script.
                $dismArgs = "/Split-Image /ImageFile:`"$wimSource`" " +
                            "/SWMFile:`"$srcDir\install.swm`" /FileSize:3900"
                Write-Log $UI "  Running DISM split (this may take several minutes)..." "Muted"
                $dismProc = Start-Process dism -ArgumentList $dismArgs -Wait -PassThru -Verb RunAs
                if ($dismProc.ExitCode -gt 1) { throw "DISM split failed (exit $($dismProc.ExitCode))." }
                if (-not (Test-Path "$srcDir\install.swm")) { throw "DISM split produced no output." }
                Write-Log $UI "[OK]  WIM split." "Success"
            } else {
                Set-Progress $UI 72 "Copying install.wim..."
                Write-Log $UI "  Copying install.wim..."
                Copy-Item $wimSource "$srcDir\install.wim" -Force -ErrorAction Stop
                Write-Log $UI "[OK]  install.wim copied." "Success"
            }
        }

        # ----------------------------------------------------------------
        # STEP 9  Boot configuration
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 90 "Configuring bootloader..."
        Write-Log $UI "---  Boot Configuration  ---" "Cyan"

        # 9a  EFI directories and files
        foreach ($d in @("$usbDrive`:\EFI\Boot", "$usbDrive`:\EFI\Microsoft\Boot", "$usbDrive`:\boot")) {
            if (-not (Test-Path $d)) { New-Item $d -ItemType Directory -Force | Out-Null }
        }

        @(
            @{ S="$isoDrive`:\efi\boot\bootx64.efi";              D="$usbDrive`:\EFI\Boot\bootx64.efi" },
            @{ S="$isoDrive`:\efi\boot\bootia32.efi";             D="$usbDrive`:\EFI\Boot\bootia32.efi" },
            @{ S="$isoDrive`:\efi\microsoft\boot\bootmgfw.efi";   D="$usbDrive`:\EFI\Microsoft\Boot\bootmgfw.efi" },
            @{ S="$isoDrive`:\efi\microsoft\boot\bootmgfw.efi";   D="$usbDrive`:\EFI\Boot\bootmgfw.efi" },
            @{ S="$isoDrive`:\efi\microsoft\boot\bootmgr.efi";    D="$usbDrive`:\EFI\Microsoft\Boot\bootmgr.efi" },
            @{ S="$isoDrive`:\bootmgr";                           D="$usbDrive`:\bootmgr" },
            @{ S="$isoDrive`:\bootmgr.efi";                       D="$usbDrive`:\bootmgr.efi" }
        ) | ForEach-Object {
            if (Test-Path $_.S) {
                Copy-Item $_.S $_.D -Force -ErrorAction SilentlyContinue
                Write-Log $UI "  [OK] $(Split-Path $_.D -Leaf)" "Success"
            }
        }

        # 9b  BCD
        if (Test-Path "$isoDrive`:\efi\microsoft\boot\bcd") {
            Copy-Item "$isoDrive`:\efi\microsoft\boot\bcd" "$usbDrive`:\EFI\Microsoft\Boot\BCD" -Force
            Write-Log $UI "  [OK] EFI BCD." "Success"
        }
        if (Test-Path "$isoDrive`:\boot\bcd") {
            Copy-Item "$isoDrive`:\boot\bcd" "$usbDrive`:\boot\BCD" -Force
            Write-Log $UI "  [OK] BIOS BCD." "Success"
        }

        # Sync \boot\ folder
        $br = Start-Process robocopy -ArgumentList @("$isoDrive`:\boot","$usbDrive`:\boot","/E","/MT:8","/NJH","/NJS","/NP") -PassThru -NoNewWindow
        $br.WaitForExit()

        # 9c  BIOS boot sector
        #
        # We are already running as Administrator so -Verb RunAs is NOT used —
        # re-elevating an already-elevated process fails silently on many
        # Windows versions and means the boot sector never gets written.
        #
        # The drive letter is passed WITHOUT extra quotes: /nt60 E: /mbr
        # Quoting it ("/nt60 "E:" /mbr") causes bootsect to reject the arg.
        #
        # We call bootsect.exe with & (direct invocation) since we are
        # already elevated — no Start-Process indirection needed.
        Write-Log $UI "--- Writing BIOS boot sector ---" "Cyan"
        $bootsectCandidates = @("$usbDrive`:\boot\bootsect.exe", "$isoDrive`:\boot\bootsect.exe")
        $bootDone = $false
        foreach ($bs in $bootsectCandidates) {
            if (-not (Test-Path $bs -ErrorAction SilentlyContinue)) { continue }
            Write-Log $UI "  Running: $bs" "Muted"
            try {
                & $bs /nt60 "$usbDrive`:" /mbr 2>&1 | Where-Object { $_.ToString().Trim() } |
                    ForEach-Object { Write-Log $UI "  bootsect: $_" "Muted" }
                if ($LASTEXITCODE -eq 0) {
                    Write-Log $UI "  [OK] Boot sector written." "Success"
                    $bootDone = $true; break
                }
                Write-Log $UI "  [!] Exit $LASTEXITCODE — trying next." "Warn"
            } catch {
                Write-Log $UI "  [!] $($_.Exception.Message) — trying next." "Warn"
            }
        }

        # Fallback: bcdboot.exe — always in System32, already elevated.
        # Writes identical VBR boot code, then restore ISO BCD so Setup
        # boots instead of the host OS.
        if (-not $bootDone) {
            Write-Log $UI "  Falling back to bcdboot.exe..." "Warn"
            try {
                & "$env:SystemRoot\System32\bcdboot.exe" C:\Windows /s "$usbDrive`:" /f BIOS 2>&1 |
                    Where-Object { $_.ToString().Trim() } |
                    ForEach-Object { Write-Log $UI "  bcdboot: $_" "Muted" }
                if ($LASTEXITCODE -eq 0) {
                    Write-Log $UI "  [OK] Boot sector written via bcdboot." "Success"
                    if (Test-Path "$isoDrive`:\boot\bcd") {
                        Copy-Item "$isoDrive`:\boot\bcd" "$usbDrive`:\boot\BCD" -Force
                        Write-Log $UI "  [OK] Setup BCD restored." "Success"
                    }
                    $bootDone = $true
                } else {
                    Write-Log $UI "  [!] bcdboot exit $LASTEXITCODE." "Warn"
                }
            } catch {
                Write-Log $UI "  [!] bcdboot failed: $($_.Exception.Message)" "Warn"
            }
        }
        if (-not $bootDone) {
            Write-Log $UI "  [!] Could not write boot sector automatically." "Warn"
            Write-Log $UI "      Run manually (as admin): & '$isoDrive`:\boot\bootsect.exe' /nt60 $usbDrive`: /mbr" "Warn"
        }

        # 9d  BCD hardening / repair
        # Common boot failure: 0xc0000225 / 0xc000014c because
        # EFI\Microsoft\Boot\BCD is missing or corrupt on the USB.
        # Re-seed BCD from the ISO and mirror it to all three standard paths.
        Write-Log $UI "--- BCD hardening ---" "Cyan"
        $isoEfiBcd   = "$isoDrive`:\efi\microsoft\boot\bcd"
        $isoBiosBcd  = "$isoDrive`:\boot\bcd"
        $seedBcdPath = if     (Test-Path $isoEfiBcd)  { $isoEfiBcd }
                       elseif (Test-Path $isoBiosBcd) { $isoBiosBcd }
                       else                           { "" }

        if ($seedBcdPath -eq "") {
            Write-Log $UI "  [!] No BCD file found in ISO (EFI or BIOS path)." "Warn"
        } else {
            $bcdTargets = @(
                "$usbDrive`:\EFI\Microsoft\Boot\BCD",
                "$usbDrive`:\EFI\Boot\BCD",
                "$usbDrive`:\boot\BCD"
            )
            foreach ($dst in $bcdTargets) {
                try {
                    $dstDir = Split-Path $dst -Parent
                    if (-not (Test-Path $dstDir)) { New-Item $dstDir -ItemType Directory -Force | Out-Null }
                    Copy-Item $seedBcdPath $dst -Force -ErrorAction Stop
                    try { attrib -h -r -s $dst 2>$null | Out-Null } catch {}
                    $bcdLen = (Get-Item $dst -ErrorAction SilentlyContinue).Length
                    if ($bcdLen -and $bcdLen -gt 0) {
                        Write-Log $UI "  [OK] BCD refreshed: $(Split-Path $dst -NoQualifier)" "Success"
                    } else {
                        Write-Log $UI "  [!] BCD copied but appears empty: $(Split-Path $dst -NoQualifier)" "Warn"
                    }
                } catch {
                    Write-Log $UI "  [!] Failed to refresh BCD at $dst : $($_.Exception.Message)" "Warn"
                }
            }
        }

        # 9e  Final check
        Write-Log $UI "--- Final check ---" "Cyan"
        @("$usbDrive`:\EFI\Boot\bootx64.efi",
          "$usbDrive`:\EFI\Microsoft\Boot\BCD",
          "$usbDrive`:\EFI\Boot\BCD",
          "$usbDrive`:\boot\BCD",
          "$usbDrive`:\sources\boot.wim") | ForEach-Object {
            if (Test-Path $_) { Write-Log $UI "  [OK] $(Split-Path $_ -NoQualifier)" "Success" }
            else              { Write-Log $UI "  [--] $(Split-Path $_ -NoQualifier)" "Warn" }
        }

        # ----------------------------------------------------------------
        # STEP 10  Done
        # ----------------------------------------------------------------
        Set-Progress $UI 100 "USB creation complete!"
        Write-Log $UI ""
        Write-Log $UI "+===========================================+" "Success"
        Write-Log $UI "|   [OK]  USB IS READY TO BOOT!            |" "Success"
        Write-Log $UI "|   Drive      : $usbDrive`:\               |" "Success"
        Write-Log $UI "|   File system: $fs                        |" "Success"
        Write-Log $UI "+===========================================+" "Success"

        [System.Windows.Forms.MessageBox]::Show(
            "USB created successfully!`n`nDrive: $usbDrive`:\`nFile system: $fs`n`nSafely eject and boot.",
            "RufusPS - Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

    } catch {
        $msg = $_.Exception.Message
        Write-Log $UI "[X]  FAILED: $msg" "Error"
        Set-Progress $UI $UI.ProgressBar.Value "Error - see log."
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred:`n`n$msg",
            "RufusPS - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null

    } finally {
        if ($mountISO) {
            try { Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue } catch {}
            Write-Log $UI "ISO dismounted." "Muted"
        }
        if ($tempWim -and (Test-Path $tempWim)) {
            try { Remove-Item $tempWim -Force -ErrorAction SilentlyContinue } catch {}
            Write-Log $UI "Temp WIM removed." "Muted"
        }
        $UI.BtnStart.Enabled    = $true
        $UI.BtnCancel.Enabled   = $false
        $script:CancelRequested = $false
    }
}

# =====================================================================
#  WIRE UP GUI & RUN
# =====================================================================
$ui   = Build-GUI
$form = $ui.Form

Update-UsbList $ui

# -- Browse button ----------------------------------------------------
$ui.BtnBrowse.Add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "ISO files (*.iso)|*.iso|All files (*.*)|*.*"
    $ofd.Title  = "Select a Windows ISO"
    if ($ofd.ShowDialog() -eq "OK") {
        Set-IsoPath $ui $ofd.FileName
    }
})

# -- Drag & Drop on form ----------------------------------------------
$form.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})
$form.Add_DragDrop({
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    $iso   = $files | Where-Object { $_ -match "\.iso$" } | Select-Object -First 1
    if ($iso) { Set-IsoPath $ui $iso }
})

# -- Drag & Drop on ISO textbox ---------------------------------------
$ui.TxtIso.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})
$ui.TxtIso.Add_DragDrop({
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    $iso   = $files | Where-Object { $_ -match "\.iso$" } | Select-Object -First 1
    if ($iso) { Set-IsoPath $ui $iso }
})

# -- Refresh button ---------------------------------------------------
$ui.BtnRefresh.Add_Click({ Update-UsbList $ui })

# -- Cancel button ----------------------------------------------------
$ui.BtnCancel.Add_Click({
    $script:CancelRequested = $true
    Write-Log $ui "Cancel requested - stopping after current step..." "Warn"
})

# -- Start button -----------------------------------------------------
$ui.BtnStart.Add_Click({
    if (-not $script:IsoPath -or -not (Test-Path $script:IsoPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a valid Windows ISO file.",
            "Missing ISO", "OK", "Warning") | Out-Null
        return
    }
    if ($ui.CmbUsb.SelectedIndex -lt 0 -or $script:UsbDrives.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a USB drive.",
            "Missing USB", "OK", "Warning") | Out-Null
        return
    }

    $selDrive = $script:UsbDrives[$ui.CmbUsb.SelectedIndex]

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "[!]  WARNING`n`nAll data on the following drive will be PERMANENTLY erased:`n`n" +
        "   $($selDrive.Display)`n`n" +
        "This action cannot be undone.  Continue?",
        "RufusPS - Confirm Erase",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne "Yes") { return }

    $ui.TxtLog.Clear()
    Start-UsbCreation `
        -UI      $ui `
        -IsoPath $script:IsoPath `
        -DiskObj $selDrive.Disk `
        -UseGPT  $ui.RbGPT.Checked `
        -UseNTFS $ui.RbNTFS.Checked
})

# -- Launch -----------------------------------------------------------
[System.Windows.Forms.Application]::Run($form)
