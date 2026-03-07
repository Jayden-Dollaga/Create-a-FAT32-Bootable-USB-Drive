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

function Refresh-UsbList {
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
    # Declare $tempWim at function scope so finally always cleans it up
    # even if an exception fires deep inside the try block.
    $tempWim                = $null

    # Reset progress bar to 0 before every run (it retains its last value).
    Set-Progress $UI 0 "Starting..."

    try {

        # ----------------------------------------------------------------
        # STEP 1  Safety checks
        # ----------------------------------------------------------------
        Set-Progress $UI 2 "Running safety checks..."
        Write-Log $UI "---  Safety Checks  ---" "Cyan"

        # FIX 1: Re-acquire a fresh disk object every time.
        # $DiskObj was captured during Refresh-UsbList, potentially minutes
        # ago.  Stale CIM objects return $null for BusType, Size, Number,
        # etc., so safety checks and disk-number capture would silently
        # operate on garbage data.  Always re-read from the storage stack.
        $freshDisk = Get-Disk -Number $DiskObj.Number -ErrorAction SilentlyContinue
        if (-not $freshDisk) {
            throw "Could not re-acquire disk $($DiskObj.Number). Unplug and re-insert the USB drive, then click Refresh."
        }
        $DiskObj = $freshDisk

        if ($DiskObj.BusType -ne 'USB') {
            throw "SAFETY: The selected disk is not a USB drive."
        }
        $systemDisks = Get-Disk | Where-Object { $_.IsSystem -or $_.IsBoot }
        if ($DiskObj.Number -in $systemDisks.Number) {
            throw "SAFETY: The selected disk contains a system or boot partition - aborting."
        }

        # Minimum 7 GB (not 8 GB): USB manufacturers use 1 GB = 1,000,000,000
        # bytes but Windows reports in GiB (1,073,741,824 bytes), so a real
        # "8 GB" stick always shows up as ~7.45 GiB in Windows.  Using 8 GB
        # as the threshold would reject every standard 8 GB drive.
        if ($DiskObj.Size -lt 7GB) {
            throw "SAFETY: USB drive is too small ($([Math]::Round($DiskObj.Size/1GB,1)) GB). " +
                  "A genuine 8 GB or larger drive is required."
        }
        Write-Log $UI "[OK]  Safety checks passed." "Success"

        # ----------------------------------------------------------------
        # STEP 2  Mount ISO
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 4 "Mounting ISO..."
        Write-Log $UI "---  Mount ISO  ---" "Cyan"
        Write-Log $UI "File: $([System.IO.Path]::GetFileName($IsoPath))"

        $mountISO = Mount-DiskImage -ImagePath $IsoPath -StorageType ISO -PassThru -ErrorAction Stop

        # Poll up to 10 seconds for the mounted volume to receive a drive letter.
        $isoVol = $null
        for ($tries = 0; $tries -lt 20; $tries++) {
            $isoVol = $mountISO | Get-Volume -ErrorAction SilentlyContinue
            if ($isoVol -and $isoVol.DriveLetter) { break }
            Start-Sleep -Milliseconds 500
        }
        $isoDrive = $isoVol.DriveLetter
        if (-not $isoDrive) {
            throw "Mounted ISO did not receive a drive letter after 10 seconds. Try remounting and run again."
        }
        Write-Log $UI "[OK]  ISO mounted at $isoDrive`:\" "Success"

        # ----------------------------------------------------------------
        # STEP 3  Detect Windows version and install image type
        # ----------------------------------------------------------------
        Set-Progress $UI 7 "Detecting Windows version..."
        Write-Log $UI "---  Windows Detection  ---" "Cyan"
        $imgInfo = Get-InstallImageInfo $isoDrive
        Write-Log $UI "  Install image  : $($imgInfo.Type)"
        Write-Log $UI "  Windows edition: $($imgInfo.Version)"
        if ($imgInfo.Type -eq "None") {
            Write-Log $UI "[!]  No install.wim or install.esd found - proceeding anyway." "Warn"
        }

        # BUG C FIX: Before committing to ESD conversion, verify that
        # $env:TEMP (usually C:\) has enough free space.  DISM expands an
        # ESD into a WIM that is typically 1.5-2x the ESD file size.
        # We require at least  esdSize * 2  free, minimum 8 GB.
        if ($imgInfo.Type -eq "ESD") {
            $esdBytes    = (Get-Item $imgInfo.Path).Length
            $needBytes   = [Math]::Max($esdBytes * 2, 8GB)
            $tempDrive   = [System.IO.Path]::GetPathRoot($env:TEMP).TrimEnd('\')[0]
            $tempFree    = (Get-Volume -DriveLetter $tempDrive -ErrorAction SilentlyContinue).SizeRemaining
            if ($tempFree -and $tempFree -lt $needBytes) {
                throw ("Not enough free space on $tempDrive`:\ for ESD conversion. " +
                       "Need $([Math]::Round($needBytes/1GB,1)) GB, " +
                       "have $([Math]::Round($tempFree/1GB,1)) GB free. " +
                       "Free up space on $tempDrive`:\ and try again.")
            }
            # Guard: $tempFree is null if Get-Volume failed (e.g. temp is on a
            # RAM disk or network share).  Skip the log line in that case.
            if ($tempFree) {
                Write-Log $UI ("  [OK] Temp drive $tempDrive`:\ has " +
                               "$([Math]::Round($tempFree/1GB,1)) GB free " +
                               "(need $([Math]::Round($needBytes/1GB,1)) GB).") "Muted"
            } else {
                Write-Log $UI "  [!] Could not read free space on temp drive $tempDrive`:\  - proceeding anyway." "Warn"
            }
        }

        # ----------------------------------------------------------------
        # STEP 4  Wipe and prepare the USB disk
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 10 "Preparing USB drive..."
        Write-Log $UI "---  Disk Preparation  ---" "Cyan"

        # Capture disk number as a plain [int] RIGHT NOW before any
        # destructive operation so it survives re-enumeration events.
        [int]$diskNum = $DiskObj.Number
        $style = if ($UseGPT) { 'GPT' } else { 'MBR' }
        Write-Log $UI "Preparing disk $diskNum as $style..." "Warn"

        # Bring the disk online and writable first.
        # Ventoy / Rufus often leave sticks offline or read-only.
        $DiskObj | Set-Disk -IsOffline $false -ErrorAction SilentlyContinue
        $DiskObj | Set-Disk -IsReadOnly $false -ErrorAction SilentlyContinue

        # Use a SINGLE diskpart session that does clean + convert in one shot.
        #
        # WHY: Every previous approach relied on WMI reporting PartitionStyle
        # = RAW after a clean.  On many USB controllers and some Windows
        # versions the storage driver caches the old style (MBR/GPT) in WMI
        # indefinitely — it never becomes RAW no matter how long you poll.
        # By issuing "clean" and "convert mbr/gpt" in the SAME diskpart
        # session we bypass WMI entirely: diskpart owns the partition table
        # directly and the convert succeeds immediately after the clean,
        # before the WMI cache has any chance to interfere.
        $convertCmd = if ($UseGPT) { 'convert gpt' } else { 'convert mbr' }
        $dpScript = @"
select disk $diskNum
clean
$convertCmd
exit
"@
        Write-Log $UI "  Running diskpart: clean + $convertCmd..." "Muted"
        $dpResult = $dpScript | & diskpart.exe 2>&1
        $dpResult | Where-Object { $_.ToString().Trim() -ne '' } | ForEach-Object {
            Write-Log $UI "  diskpart: $_" "Muted"
        }

        # Verify diskpart succeeded by checking the output text.
        # "DiskPart succeeded in cleaning" and "DiskPart successfully converted"
        # are the exact strings diskpart.exe emits on success.
        $dpOut = ($dpResult | Out-String)
        if ($dpOut -notmatch 'succeeded in cleaning') {
            throw "diskpart clean did not report success. Output: $dpOut"
        }
        if ($dpOut -notmatch 'successfully converted') {
            throw "diskpart $convertCmd did not report success. Output: $dpOut"
        }

        # Wait for the storage stack to re-enumerate the newly initialised
        # disk.  We do NOT check PartitionStyle here — WMI is unreliable
        # immediately after diskpart.  We just wait for the disk to be
        # visible and for Initialize-Disk to accept it (or confirm it is
        # already correctly initialised by diskpart).
        Write-Log $UI "  Waiting for disk $diskNum to re-enumerate..." "Muted"
        $DiskObj = $null
        for ($tries = 0; $tries -lt 30; $tries++) {
            try {
                $DiskObj = Get-Disk -Number $diskNum -ErrorAction Stop
                break
            } catch {
                Start-Sleep -Milliseconds 500
            }
        }
        if (-not $DiskObj) {
            throw "Disk $diskNum did not re-appear after diskpart. Try re-inserting the USB drive."
        }

        # Initialize-Disk is not needed — diskpart "convert mbr/gpt" above
        # already wrote the partition table.  Calling it causes "already
        # initialized" errors on every run.  Removed.

        Start-Sleep -Milliseconds 800
        $DiskObj = Get-Disk -Number $diskNum
        Write-Log $UI "[OK]  Disk $diskNum prepared ($style, $($DiskObj.PartitionStyle))." "Success"

        # ----------------------------------------------------------------
        # STEP 5  Create partition and format
        # ----------------------------------------------------------------
        Set-Progress $UI 14 "Creating partition..."
        Write-Log $UI "---  Partition & Format  ---" "Cyan"

        $fsLabel  = "WINUSB"
        $fs       = if ($UseNTFS) { "NTFS" } else { "FAT32" }

        # FIX 2: Use -UseMaximumSize instead of a calculated -Size.
        # Computing $DiskObj.Size - 8MB and passing it as -Size fails on
        # some drives where the controller reports a slightly smaller usable
        # area than the disk size (GPT backup table, alignment reserves,
        # bad-block sparing).  -UseMaximumSize always succeeds because
        # Windows calculates the actual maximum internally.
        Write-Log $UI "Creating $fs partition (maximum size)..."

        # GPT: tag with EFI System Partition GUID so UEFI firmware sees it.
        # MBR: mark active (bootable) flag.
        if ($UseGPT) {
            $efiGuid = '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}'
            $newPart = $DiskObj | New-Partition -UseMaximumSize -GptType $efiGuid -ErrorAction Stop
        } else {
            $newPart = $DiskObj | New-Partition -UseMaximumSize -IsActive -ErrorAction Stop
        }

        Start-Sleep -Milliseconds 800

        # Assign a drive letter if Windows did not auto-assign one (common
        # with GPT ESP partitions).
        if (-not $newPart.DriveLetter) {
            Write-Log $UI "  Drive letter not auto-assigned - assigning now..." "Muted"
            $newPart | Add-PartitionAccessPath -AssignDriveLetter -ErrorAction Stop
            Start-Sleep -Milliseconds 500
            $newPart = Get-Partition -DiskNumber $diskNum `
                                     -PartitionNumber $newPart.PartitionNumber
        }

        $usbDrive = $newPart.DriveLetter
        if (-not $usbDrive) {
            throw "Failed to assign a drive letter to the USB partition."
        }

        Write-Log $UI "Formatting as $fs (label: $fsLabel)..."
        # FIX 3: Some Windows 10 builds still enforce the 32 GB FAT32 limit
        # in the Storage Management API that Format-Volume calls.  If it
        # throws, fall back to diskpart's format command which has no such
        # restriction on Windows 10/11.
        $formatOk = $false
        try {
            $newPart | Format-Volume -FileSystem $fs `
                                     -NewFileSystemLabel $fsLabel `
                                     -Force -Confirm:$false `
                                     -ErrorAction Stop | Out-Null
            $formatOk = $true
        } catch {
            if (-not $UseNTFS) {
                Write-Log $UI "  Format-Volume FAT32 failed ($($_.Exception.Message))." "Warn"
                Write-Log $UI "  Retrying with diskpart format (no 32 GB limit)..." "Warn"
                $dpFmt = @"
select disk $diskNum
select partition $($newPart.PartitionNumber)
format fs=fat32 label=$fsLabel quick
exit
"@
                $dpFmtOut = $dpFmt | & diskpart.exe 2>&1
                $dpFmtOut | Where-Object { $_.ToString().Trim() -ne '' } |
                    ForEach-Object { Write-Log $UI "  diskpart: $_" "Muted" }
                # Verify the format succeeded by reading the volume back.
                Start-Sleep -Milliseconds 800
                $vol = Get-Partition -DiskNumber $diskNum -PartitionNumber $newPart.PartitionNumber |
                       Get-Volume -ErrorAction SilentlyContinue
                if ($vol -and $vol.FileSystem -eq 'FAT32') {
                    $formatOk = $true
                    Write-Log $UI "  [OK] diskpart FAT32 format succeeded." "Success"
                } else {
                    throw "diskpart FAT32 format also failed. The drive may be faulty."
                }
            } else {
                throw
            }
        }
        if (-not $formatOk) { throw "Format failed with no further details." }

        Write-Log $UI "[OK]  Partition ready: $usbDrive`:\ ($fs)" "Success"

        # ----------------------------------------------------------------
        # STEP 6  Copy all ISO files except the install image (Robocopy)
        # ----------------------------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 18 "Copying Windows files via Robocopy (multi-threaded)..."
        Write-Log $UI "---  File Copy  ---" "Cyan"
        Write-Log $UI "Source -> $isoDrive`:\   Destination -> $usbDrive`:\"
        Write-Log $UI "Excluding install.wim / install.esd  (handled separately)"

        # FIX 4: On FAT32, no single file may exceed 4 GB - 1 byte.
        # Scan the ISO for any file (other than install.wim/esd which are
        # handled separately) that would exceed this limit.  Robocopy would
        # exit with code 8 on such files, but the error message is cryptic.
        # Give the user a clear, actionable error before wasting time.
        if (-not $UseNTFS) {
            $fatLimit   = 4294967295   # 4 GiB - 1 byte
            $largeFiles = Get-ChildItem -Path "$isoDrive`:\" -Recurse -File -ErrorAction SilentlyContinue |
                          Where-Object {
                              $_.Length -gt $fatLimit -and
                              $_.Name -notmatch '^install\.(wim|esd)$'
                          }
            if ($largeFiles) {
                $list = ($largeFiles | ForEach-Object {
                    "  $($_.FullName)  ($([Math]::Round($_.Length/1GB,2)) GB)"
                }) -join "`n"
                throw ("FAT32 cannot store files larger than 4 GB, but the ISO contains:`n$list`n`n" +
                       "These files are not install images and cannot be auto-split. " +
                       "Use NTFS format for this ISO.")
            }
        }

        $roboArgs = @(
            "$isoDrive`:\", "$usbDrive`:\",
            "/E",       # all subdirectories including empty
            "/MT:16",   # 16 parallel threads
            "/NJH",     # no job header
            "/NJS",     # no job summary
            "/NDL",     # no directory listing noise
            "/NP",      # no per-file percentage
            "/XF", "install.wim", "install.esd"
        )

        $roboJob = Start-Process robocopy `
            -ArgumentList $roboArgs `
            -PassThru -NoNewWindow

        $rStart = Get-Date
        while (-not $roboJob.HasExited) {
            if ($script:CancelRequested) { try { $roboJob.Kill() } catch {}; throw "Cancelled." }
            $elapsed = ((Get-Date) - $rStart).TotalSeconds
            $estPct  = [Math]::Min(18 + [int]($elapsed / 2), 48)
            Set-Progress $UI $estPct "Copying files... ($([int]$elapsed)s elapsed)"
            Start-Sleep -Milliseconds 400
        }

        # WaitForExit() required for ExitCode to be reliable on PassThru handles.
        $roboJob.WaitForExit()

        # Robocopy exit codes 0-7 are all success variants.
        if ($roboJob.ExitCode -gt 7) {
            throw "Robocopy failed with exit code $($roboJob.ExitCode)."
        }
        Write-Log $UI "[OK]  Files copied successfully." "Success"

        # ----------------------------------------------------------------
        # STEP 7  ESD -> WIM conversion (if needed)
        # ----------------------------------------------------------------
        $wimSource = ""

        if ($imgInfo.Type -eq "ESD") {
            if ($script:CancelRequested) { throw "Cancelled." }
            Set-Progress $UI 50 "Converting install.esd -> install.wim..."
            Write-Log $UI "---  ESD -> WIM Conversion  ---" "Cyan"
            Write-Log $UI "This may take 5-20 minutes depending on drive speed." "Warn"

            $esdFile = $imgInfo.Path
            $tempWim = "$env:TEMP\rufusps_install.wim"
            if (Test-Path $tempWim) { Remove-Item $tempWim -Force }

            # Get all indexes.  $rawInfo is a string[] so pipe through
            # Out-String before regex to preserve newlines between elements.
            $rawInfo     = & dism /Get-WimInfo "/WimFile:$esdFile" 2>&1
            $rawInfoText = ($rawInfo | Out-String)
            $indexes     = [regex]::Matches($rawInfoText, "Index\s*:\s*(\d+)") |
                           ForEach-Object { $_.Groups[1].Value }

            if ($indexes.Count -eq 0) {
                throw "Could not enumerate indexes in install.esd."
            }
            Write-Log $UI "  Found $($indexes.Count) image index(es) in ESD."

            # BUG H FIX: Distribute the 50->70 progress range evenly
            # across ALL indexes so each one moves the bar, not just the
            # last.  Previous code gave each index only 5% regardless of
            # count, leaving the bar stuck at 52% for single-index ESDs.
            $pctPerIndex = [int](20 / $indexes.Count)   # 20 points = 50..70

            for ($i = 0; $i -lt $indexes.Count; $i++) {
                if ($script:CancelRequested) { throw "Cancelled." }
                $idx     = $indexes[$i]
                $basePct = 50 + ($i * $pctPerIndex)
                Write-Log $UI "  Exporting index $idx / $($indexes.Count)..."

                $dismArgs = "/Export-Image " +
                            "/SourceImageFile:`"$esdFile`" " +
                            "/SourceIndex:$idx " +
                            "/DestinationImageFile:`"$tempWim`" " +
                            "/Compress:fast"

                $code    = Invoke-DismWithProgress $UI $dismArgs $basePct $pctPerIndex "Converting index $idx"
                $codeInt = if ($null -eq $code -or $code -eq '') { 0 } else { [int]$code }
                if ($codeInt -gt 1) {
                    Write-Log $UI "  [!]  Index $idx returned code $codeInt - continuing." "Warn"
                }
            }

            if (-not (Test-Path $tempWim)) {
                throw "ESD -> WIM conversion produced no output file."
            }
            $wimSource = $tempWim
            Write-Log $UI "[OK]  ESD -> WIM conversion complete." "Success"

        } elseif ($imgInfo.Type -eq "WIM") {
            $wimSource = $imgInfo.Path
        }

        # ----------------------------------------------------------------
        # STEP 8  Copy or split the install image onto the USB
        # ----------------------------------------------------------------
        if ($wimSource -ne "") {
            if ($script:CancelRequested) { throw "Cancelled." }
            $wimBytes = (Get-Item $wimSource).Length
            $wimGB    = [Math]::Round($wimBytes / 1GB, 2)
            Write-Log $UI "---  Install Image  ---" "Cyan"
            Write-Log $UI "  Path: $wimSource"
            Write-Log $UI "  Size: $wimGB GB"

            $srcDir = "$usbDrive`:\sources"
            if (-not (Test-Path $srcDir)) { New-Item $srcDir -ItemType Directory | Out-Null }

            # BUG D FIX: Check that the USB has enough free space before
            # starting a potentially long copy/split that would fail midway
            # with a cryptic "not enough disk space" DISM or Copy-Item error.
            $usbFree      = (Get-Volume -DriveLetter $usbDrive -ErrorAction SilentlyContinue).SizeRemaining
            $neededOnUsb  = $wimBytes + 200MB   # 200 MB safety headroom
            if ($usbFree -and $usbFree -lt $neededOnUsb) {
                throw ("Not enough free space on $usbDrive`:\ for the install image. " +
                       "Need $([Math]::Round($neededOnUsb/1GB,1)) GB, " +
                       "have $([Math]::Round($usbFree/1GB,1)) GB free.")
            }

            if ((-not $UseNTFS) -and ($wimBytes -gt 4GB)) {
                # FAT32 cannot store files > 4 GB.  Split the WIM into
                # 4096 MB .swm chunks.  Windows Setup automatically finds
                # install*.swm files in \sources\ at setup time.
                Set-Progress $UI 72 "Splitting install.wim for FAT32 (> 4 GB)..."
                Write-Log $UI "WIM > 4 GB on FAT32 - splitting into .swm chunks..." "Warn"

                $swmOut   = "$srcDir\install.swm"
                $dismArgs = "/Split-Image " +
                            "/ImageFile:`"$wimSource`" " +
                            "/SWMFile:`"$swmOut`" " +
                            "/FileSize:3900"
                # NOTE: /FileSize is in MB.  FAT32 max file size is exactly
                # 4 GiB - 1 byte (4,294,967,295 B = 4095.9999... MB).
                # Using 4096 MB would produce a chunk of 4,294,967,296 B
                # which is 1 byte over the FAT32 limit and unwritable.
                # 3900 MB gives a safe ~196 MB margin.

                $code      = Invoke-DismWithProgress $UI $dismArgs 72 18 "Splitting WIM"
                $codeInt   = if ($null -eq $code -or $code -eq '') { 0 } else { [int]$code }
                $swmExists = Test-Path "$srcDir\install.swm"

                if ($codeInt -gt 1 -and -not $swmExists) {
                    throw "DISM WIM split failed (exit code $codeInt). No .swm files were created."
                }
                if (-not $swmExists) {
                    throw "DISM reported success but no install.swm was created. Check disk space."
                }
                Write-Log $UI "[OK]  WIM split successfully." "Success"

            } else {
                # NTFS, or FAT32 with a WIM that fits in one file (< 4 GB).
                Set-Progress $UI 72 "Copying install.wim to USB..."
                Write-Log $UI "Copying install.wim..."
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

        # 9a  Ensure all critical EFI directories exist -----------------
        Write-Log $UI "--- Verifying EFI boot files ---" "Cyan"
        foreach ($d in @("$usbDrive`:\EFI\Boot", "$usbDrive`:\EFI\Microsoft\Boot", "$usbDrive`:\boot")) {
            if (-not (Test-Path $d)) {
                New-Item $d -ItemType Directory -Force | Out-Null
                Write-Log $UI "  Created dir: $d" "Muted"
            }
        }

        # Force-copy every critical EFI file from the ISO, overwriting
        # anything that Robocopy may have placed (ensures freshest copy).
        $efiFilemap = @(
            @{ Src = "$isoDrive`:\efi\boot\bootx64.efi";
               Dst = "$usbDrive`:\EFI\Boot\bootx64.efi";
               Label = "EFI\Boot\bootx64.efi" },
            @{ Src = "$isoDrive`:\efi\boot\bootia32.efi";
               Dst = "$usbDrive`:\EFI\Boot\bootia32.efi";
               Label = "EFI\Boot\bootia32.efi" },
            @{ Src = "$isoDrive`:\efi\microsoft\boot\bootmgfw.efi";
               Dst = "$usbDrive`:\EFI\Microsoft\Boot\bootmgfw.efi";
               Label = "EFI\Microsoft\Boot\bootmgfw.efi" },
            @{ Src = "$isoDrive`:\efi\microsoft\boot\bootmgfw.efi";
               Dst = "$usbDrive`:\EFI\Boot\bootmgfw.efi";
               Label = "EFI\Boot\bootmgfw.efi (fallback)" },
            @{ Src = "$isoDrive`:\efi\microsoft\boot\bootmgr.efi";
               Dst = "$usbDrive`:\EFI\Microsoft\Boot\bootmgr.efi";
               Label = "EFI\Microsoft\Boot\bootmgr.efi" },
            @{ Src = "$isoDrive`:\bootmgr";
               Dst = "$usbDrive`:\bootmgr";
               Label = "bootmgr" },
            @{ Src = "$isoDrive`:\bootmgr.efi";
               Dst = "$usbDrive`:\bootmgr.efi";
               Label = "bootmgr.efi" }
        )
        foreach ($entry in $efiFilemap) {
            if (Test-Path $entry.Src) {
                Copy-Item $entry.Src $entry.Dst -Force -ErrorAction SilentlyContinue
                Write-Log $UI "  [OK] $($entry.Label)" "Success"
            } else {
                Write-Log $UI "  [--] Not in ISO: $($entry.Label)" "Muted"
            }
        }

        # 9b  BCD store ------------------------------------------------
        Write-Log $UI "--- BCD store ---" "Cyan"

        $bcdSrc     = "$isoDrive`:\efi\microsoft\boot\bcd"
        $bcdDst     = "$usbDrive`:\EFI\Microsoft\Boot\BCD"
        $bcdBootSrc = "$isoDrive`:\boot\bcd"
        $bcdBootDst = "$usbDrive`:\boot\BCD"

        if (Test-Path $bcdSrc) {
            Copy-Item $bcdSrc $bcdDst -Force
            Write-Log $UI "  [OK] EFI BCD copied." "Success"
        } else {
            Write-Log $UI "  [!] EFI BCD not found in ISO." "Warn"
        }

        if (Test-Path $bcdBootSrc) {
            Copy-Item $bcdBootSrc $bcdBootDst -Force
            Write-Log $UI "  [OK] BIOS BCD copied." "Success"
        }

        # Sync entire \boot\ folder for full BIOS compatibility
        # (bootsect.exe, fonts, resources, memtest, etc.)
        Write-Log $UI "  Syncing \\boot\\ folder..." "Muted"
        $bootRobo = Start-Process robocopy `
            -ArgumentList @("$isoDrive`:\boot", "$usbDrive`:\boot",
                            "/E", "/MT:8", "/NJH", "/NJS", "/NP") `
            -PassThru -NoNewWindow
        # BUG G FIX: WaitForExit() required before ExitCode is reliable
        # on Start-Process -PassThru handles, even after -Wait is used.
        # Changed to not use -Wait so we can call WaitForExit() explicitly.
        $bootRobo.WaitForExit()
        if ($bootRobo.ExitCode -le 7) {
            Write-Log $UI "  [OK] \\boot\\ folder synced." "Success"
        } else {
            Write-Log $UI "  [!] \\boot\\ sync returned code $($bootRobo.ExitCode)." "Warn"
        }

        # 9c  MBR/VBR boot sector (BIOS boot only) ---------------------
        if (-not $UseGPT) {
            Write-Log $UI "--- Writing BIOS boot sector (bootsect) ---" "Cyan"

            $bootsectExe = @("$usbDrive`:\boot\bootsect.exe",
                             "$isoDrive`:\boot\bootsect.exe") |
                           Where-Object { Test-Path $_ } |
                           Select-Object -First 1

            if ($bootsectExe) {
                try {
                    $bsOut = & "$bootsectExe" /nt60 "$usbDrive`:" /force /mbr 2>&1
                    $bsOut | Where-Object { $_.Trim() -ne "" } |
                             ForEach-Object { Write-Log $UI "  bootsect: $_" "Muted" }
                    if ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE) {
                        Write-Log $UI "  [OK] VBR + MBR boot code written." "Success"
                    } else {
                        Write-Log $UI "  [!] bootsect exit code: $LASTEXITCODE (BIOS boot may not work)." "Warn"
                    }
                } catch {
                    # bootsect.exe itself is corrupted or unreadable on this ISO.
                    # This is non-fatal — UEFI boot via EFI files is unaffected.
                    # Only legacy BIOS boot is impaired.
                    Write-Log $UI "  [!] bootsect.exe could not run: $($_.Exception.Message)" "Warn"
                    Write-Log $UI "  [!] BIOS (legacy) boot may not work, but UEFI boot is unaffected." "Warn"
                    Write-Log $UI "      To fix BIOS boot manually, run on a working machine:" "Warn"
                    Write-Log $UI "      bootsect /nt60 $usbDrive`: /force /mbr" "Warn"
                }
            } else {
                Write-Log $UI "  [!] bootsect.exe not found - BIOS boot may not work." "Warn"
                Write-Log $UI "      Run manually: bootsect /nt60 $usbDrive`: /force /mbr" "Warn"
            }
        }

        # 9d  GPT: set EFI System Partition attributes via diskpart -----
        if ($UseGPT) {
            Write-Log $UI "--- Setting GPT ESP attributes ---" "Cyan"
            # BUG E FIX: Use the already-captured [int]$diskNum instead of
            # re-reading $DiskObj.Number (which shadows $diskNum and could
            # read a stale value if $DiskObj was not refreshed recently).
            $partNum = (Get-Partition -DiskNumber $diskNum |
                        Where-Object { $_.DriveLetter -eq $usbDrive }).PartitionNumber
            if ($partNum) {
                $dpScript = @"
select disk $diskNum
select partition $partNum
gpt attributes=0x8000000000000001
exit
"@
                $dpScript | & diskpart.exe | Out-Null
                Write-Log $UI "  [OK] GPT partition attributes set." "Success"
            }
        }

        # 9e  Final sanity check ----------------------------------------
        Write-Log $UI "--- Final boot file check ---" "Cyan"
        $mustExist = @(
            "$usbDrive`:\EFI\Boot\bootx64.efi",
            "$usbDrive`:\EFI\Microsoft\Boot\bootmgfw.efi",
            "$usbDrive`:\EFI\Microsoft\Boot\BCD",
            "$usbDrive`:\sources\boot.wim"
        )
        $allGood = $true
        foreach ($f in $mustExist) {
            if (Test-Path $f) {
                Write-Log $UI "  [OK] $(Split-Path $f -NoQualifier)" "Success"
            } else {
                Write-Log $UI "  [X] MISSING: $(Split-Path $f -NoQualifier)" "Error"
                $allGood = $false
            }
        }
        if (-not $allGood) {
            Write-Log $UI "[!] Some boot files are missing. The USB may not boot on all systems." "Warn"
            Write-Log $UI "    This is normal for stripped ISOs (e.g. Tiny11)." "Warn"
        }

        # ----------------------------------------------------------------
        # STEP 10  Done
        # ----------------------------------------------------------------
        Set-Progress $UI 100 "[OK]  USB creation complete!"
        Write-Log $UI ""
        Write-Log $UI "+==========================================+" "Success"
        Write-Log $UI "|   [OK]  USB IS READY TO BOOT!           |" "Success"
        Write-Log $UI "|   Drive      : $usbDrive`:\              |" "Success"
        Write-Log $UI "|   Partition  : $(if ($UseGPT)  {'GPT (UEFI only)'} else {'MBR (UEFI + BIOS)'})   |" "Success"
        Write-Log $UI "|   File system: $(if ($UseNTFS) {'NTFS'} else {'FAT32'})                    |" "Success"
        Write-Log $UI "+==========================================+" "Success"

        [System.Windows.Forms.MessageBox]::Show(
            "USB drive created successfully!`n`n" +
            "Drive       : $usbDrive`:\`n" +
            "Partition   : $(if ($UseGPT) {'GPT  (UEFI only)'} else {'MBR  (UEFI + BIOS)'})`n" +
            "File system : $(if ($UseNTFS) {'NTFS'} else {'FAT32'})`n`n" +
            "You can now safely eject and boot from this drive.",
            "RufusPS - Success",
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
        # Always dismount the ISO so it does not stay locked.
        if ($mountISO) {
            try { Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue } catch {}
            Write-Log $UI "ISO dismounted." "Muted"
        }

        # BUG B FIX: Clean up the temp WIM here in finally so it is
        # ALWAYS deleted even if an exception fires during Step 8 (WIM
        # split/copy) before the inline cleanup code could run.
        # Previously the cleanup was inline between Steps 8 and 9 meaning
        # any throw in Step 8 left a 6-9 GB file in $env:TEMP permanently.
        if ($tempWim -and (Test-Path $tempWim)) {
            try { Remove-Item $tempWim -Force -ErrorAction SilentlyContinue } catch {}
            Write-Log $UI "Temporary WIM removed." "Muted"
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

Refresh-UsbList $ui

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
$ui.BtnRefresh.Add_Click({ Refresh-UsbList $ui })

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
