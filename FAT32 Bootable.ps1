#Requires -Version 5.1
<#
.SYNOPSIS
    RufusPS - Advanced Windows USB Creator
.DESCRIPTION
    GUI-based bootable USB creator with:
      * Full UEFI + Legacy BIOS support
      * FAT32 (auto WIM-split) or NTFS
      * MBR or GPT partition scheme
      * install.wim and install.esd support (auto-converts ESD->WIM)
      * Automatic WIM splitting for FAT32 (>4 GB files)
      * EFI boot repair (fixes 0xc0000225 / winload.efi)
      * Fast Robocopy multi-threaded file copy
      * Real-time DISM progress bars
      * USB safety checks (prevents wiping wrong disk)
      * Drag-and-drop ISO support
      * Auto admin elevation
      * Windows 10 / 11 / Server / Tiny11 compatible
.NOTES
    Run this script directly - it will auto-elevate to Administrator if needed.
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
    $rbFAT32.Text      = "FAT32  -  Universal  (UEFI + BIOS, auto-splits WIM > 4 GB)"
    $rbFAT32.Location  = New-Object System.Drawing.Point(0, 4)
    $rbFAT32.Size      = New-Object System.Drawing.Size(360, 36)
    $rbFAT32.BackColor = $C.BG
    $rbFAT32.ForeColor = $C.Text
    $rbFAT32.Checked   = $true
    $pnlFileSys.Controls.Add($rbFAT32)

    $rbNTFS           = New-Object System.Windows.Forms.RadioButton
    $rbNTFS.Text      = "NTFS  -  No 4 GB limit  (BIOS only)"
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

# BUG 11 FIX: Added try/catch around Get-Disk so that a missing Storage
# module shows a friendly warning instead of crashing the script.
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

# BUG 12 FIX: Added .iso extension validation - non-ISO files were
# silently accepted and would fail much later with a cryptic error.
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

    $outFile = [System.IO.Path]::GetTempFileName()
    $errFile = [System.IO.Path]::GetTempFileName()

    $proc = Start-Process dism.exe `
        -ArgumentList $Arguments `
        -RedirectStandardOutput $outFile `
        -RedirectStandardError  $errFile `
        -NoNewWindow -PassThru

    $reader   = $null
    $stream   = $null
    $lastPct  = 0
    # BUG 2 FIX: Use a local variable (not $script:LastDismExitCode) so
    # the exit code is captured inside try before finally closes streams.
    # The original read a script-scope var after finally ran - a race that
    # could return null when DISM finished faster than the polling loop.
    $exitCode = 0

    try {
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
                # BUG 1 FIX: Original regex was '(\d+\....\d*)%'.
                # In regex '\.' means "any character" not a literal dot,
                # and '....' means any 4 characters - so the pattern never
                # matched real DISM lines like "10.0%" or "100%".
                # Correct pattern uses '\.?' for an optional literal dot.
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

        # Capture exit code here, inside try, BEFORE finally closes streams.
        $exitCode = if ($null -eq $proc.ExitCode) { 0 } else { [int]$proc.ExitCode }

        # Drain any remaining output after DISM exits.
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

    try {

        # -- 1. Safety checks -----------------------------------------
        Set-Progress $UI 2 "Running safety checks..."
        Write-Log $UI "---  Safety Checks  ---" "Cyan"

        if ($DiskObj.BusType -ne 'USB') {
            throw "SAFETY: The selected disk is not a USB drive."
        }
        $systemDisk = Get-Disk | Where-Object { $_.IsSystem -or $_.IsBoot }
        if ($DiskObj.Number -in $systemDisk.Number) {
            throw "SAFETY: The selected disk contains a system or boot partition - aborting."
        }
        if ($DiskObj.Size -lt 4GB) {
            throw "SAFETY: USB drive is too small (< 4 GB)."
        }
        Write-Log $UI "[OK]  Safety checks passed." "Success"

        # -- 2. Mount ISO ---------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 4 "Mounting ISO..."
        Write-Log $UI "---  Mount ISO  ---" "Cyan"
        Write-Log $UI "File: $([System.IO.Path]::GetFileName($IsoPath))"

        $mountISO = Mount-DiskImage -ImagePath $IsoPath -StorageType ISO -PassThru -ErrorAction Stop

        # BUG 10 FIX: Original code slept 1 s then read DriveLetter once.
        # On slower systems the volume isn't registered yet and $isoDrive
        # is null, crashing all downstream paths. Poll up to 10 s instead.
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

        # -- 3. Detect Windows version --------------------------------
        Set-Progress $UI 7 "Detecting Windows version..."
        Write-Log $UI "---  Windows Detection  ---" "Cyan"
        $imgInfo = Get-InstallImageInfo $isoDrive
        Write-Log $UI "  Install image : $($imgInfo.Type)"
        Write-Log $UI "  Windows edition: $($imgInfo.Version)"
        if ($imgInfo.Type -eq "None") {
            Write-Log $UI "[!]  No install.wim or install.esd found - proceeding anyway." "Warn"
        }

        # -- 4. Prepare disk ------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 10 "Preparing USB drive..."
        Write-Log $UI "---  Disk Preparation  ---" "Cyan"

        # CRITICAL: Capture the disk number as a plain [int] RIGHT NOW,
        # before any destructive operation.  Clear-Disk causes Windows to
        # briefly remove and re-enumerate the USB device.  If we read
        # $DiskObj.Number after that momentary disappearance the property
        # is null, which breaks every subsequent Get-Disk / diskpart call.
        [int]$diskNum = $DiskObj.Number
        Write-Log $UI "Clearing disk $diskNum..." "Warn"

        # Ensure the disk is online and writable first.  USB sticks left
        # offline or read-only by Ventoy / Rufus will make Clear-Disk throw.
        $DiskObj | Set-Disk -IsOffline $false -ErrorAction SilentlyContinue
        $DiskObj | Set-Disk -IsReadOnly $false -ErrorAction SilentlyContinue

        # Stage 1 – PowerShell Clear-Disk (non-fatal; it often warns on
        # drives that are about to be re-enumerated).
        Write-Log $UI "  Running Clear-Disk..." "Muted"
        try {
            $DiskObj | Clear-Disk -RemoveData -Confirm:$false -ErrorAction Stop
        } catch {
            Write-Log $UI "  Clear-Disk warning: $($_.Exception.Message)" "Warn"
        }

        # Stage 2 – Re-acquire the disk object with a retry loop.
        # After Clear-Disk the USB device disappears from the storage stack
        # for up to ~3 seconds while Windows re-enumerates it.  Polling
        # here prevents the "No MSFT_Disk objects found" crash that occurred
        # when Get-Disk ran while the device was still offline.
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
            throw "Disk $diskNum did not re-appear after 15 seconds. Try re-inserting the USB drive and run again."
        }
        Write-Log $UI "  [OK] Disk $diskNum re-enumerated (style: $($DiskObj.PartitionStyle))." "Muted"

        # Stage 3 – If still not RAW, run diskpart clean which is the
        # definitive command for resetting a partition table.  GPT drives
        # keep a protective MBR entry and won't be RAW after Clear-Disk alone.
        if ($DiskObj.PartitionStyle -ne 'RAW') {
            Write-Log $UI "  Not RAW (style: $($DiskObj.PartitionStyle)) - running diskpart clean..." "Warn"

            $dpScript = "select disk $diskNum`nclean`nexit"
            $dpResult = $dpScript | & diskpart.exe 2>&1
            $dpResult | Where-Object { $_.ToString().Trim() -ne '' } | ForEach-Object {
                Write-Log $UI "  diskpart: $_" "Muted"
            }

            # Poll until the disk is both visible AND shows PartitionStyle RAW.
            # Two separate delays are needed:
            #   1. diskpart clean briefly takes the device offline (~0-2 s)
            #      so Get-Disk may throw "No MSFT_Disk objects found" first.
            #   2. Once the device is back, WMI can still report the old
            #      partition style (MBR/GPT) for another ~2-4 s while the
            #      storage driver flushes its metadata cache.
            # Polling for RAW specifically handles both conditions in one loop.
            Write-Log $UI "  Waiting for disk $diskNum to report RAW..." "Muted"
            $DiskObj = $null
            for ($tries = 0; $tries -lt 40; $tries++) {        # up to 20 s
                try {
                    $d = Get-Disk -Number $diskNum -ErrorAction Stop
                    if ($d.PartitionStyle -eq 'RAW') {
                        $DiskObj = $d
                        break
                    }
                } catch {
                    # Disk not yet visible - keep waiting.
                }
                Start-Sleep -Milliseconds 500
            }

            if (-not $DiskObj) {
                # Last-chance read: if the disk appeared but WMI is still
                # lagging, grab whatever state it reports and let the check
                # below throw a descriptive error.
                try { $DiskObj = Get-Disk -Number $diskNum -ErrorAction Stop } catch {}
            }

            if (-not $DiskObj) {
                throw "Disk $diskNum did not re-appear after diskpart clean. Try re-inserting the USB drive."
            }

            if ($DiskObj.PartitionStyle -ne 'RAW') {
                throw "Disk $diskNum is still not RAW after diskpart clean (style: $($DiskObj.PartitionStyle)). " +
                      "Try ejecting and re-inserting the USB drive, then run again."
            }
            Write-Log $UI "  [OK] diskpart clean succeeded - disk is now RAW." "Success"
        } else {
            Write-Log $UI "  [OK] Disk is RAW." "Success"
        }

        $style = if ($UseGPT) { 'GPT' } else { 'MBR' }
        Write-Log $UI "Initializing as $style..."
        $DiskObj | Initialize-Disk -PartitionStyle $style -ErrorAction Stop
        Start-Sleep -Milliseconds 800
        # Use the saved [int]$diskNum — never read .Number from a stale object.
        $DiskObj = Get-Disk -Number $diskNum
        Write-Log $UI "[OK]  Disk initialized ($style)." "Success"

        # -- 5. Create partition & format -----------------------------
        Set-Progress $UI 14 "Creating partition..."
        Write-Log $UI "---  Partition & Format  ---" "Cyan"

        $fsLabel = "WINUSB"
        $fs      = if ($UseNTFS) { "NTFS" } else { "FAT32" }

        # BUG 5 FIX: The original used $DiskObj.Size - 8MB unconditionally.
        # Windows cannot format FAT32 partitions > 32 GB and will fail.
        # Cap at 32 GB when the user has chosen FAT32 on a large drive.
        if (-not $UseNTFS -and $DiskObj.Size -gt 32GB) {
            $partSize = 32GB
            Write-Log $UI "[!] FAT32 capped at 32 GB (Windows FAT32 format limit)." "Warn"
        } else {
            $partSize = $DiskObj.Size - 8MB   # small alignment safety margin
        }

        Write-Log $UI "Creating $fs partition ($([Math]::Round($partSize / 1GB, 1)) GB)..."

        # BUG 6 FIX: GPT branch used plain New-Partition with no GptType.
        # Without the EFI System Partition GUID, UEFI firmware does not
        # recognise the partition as a boot target and skips the drive.
        #
        # BUG 7 FIX: Piping New-Partition directly to Format-Volume gives
        # no handle to call Add-PartitionAccessPath when Windows doesn't
        # auto-assign a drive letter (common with GPT ESP partitions).
        # Store the new partition in $newPart and assign the letter
        # explicitly before formatting.
        if ($UseGPT) {
            $efiGuid = '{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}'
            $newPart = $DiskObj | New-Partition -Size $partSize -GptType $efiGuid -ErrorAction Stop
        } else {
            $newPart = $DiskObj | New-Partition -Size $partSize -IsActive -ErrorAction Stop
        }

        Start-Sleep -Milliseconds 800

        # Assign a drive letter if Windows did not auto-assign one.
        if (-not $newPart.DriveLetter) {
            Write-Log $UI "  Drive letter not auto-assigned - assigning now..." "Muted"
            $newPart | Add-PartitionAccessPath -AssignDriveLetter -ErrorAction Stop
            Start-Sleep -Milliseconds 500
            $newPart = Get-Partition -DiskNumber $DiskObj.Number `
                                     -PartitionNumber $newPart.PartitionNumber
        }

        $usbDrive = $newPart.DriveLetter
        if (-not $usbDrive) {
            throw "Failed to assign a drive letter to the USB partition."
        }

        Write-Log $UI "Formatting as $fs (label: $fsLabel)..."
        $newPart | Format-Volume -FileSystem $fs `
                                 -NewFileSystemLabel $fsLabel `
                                 -Force -Confirm:$false `
                                 -ErrorAction Stop | Out-Null

        Write-Log $UI "[OK]  Partition ready: $usbDrive`:\ ($fs)" "Success"

        # -- 6. Copy files with Robocopy ------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 18 "Copying Windows files via Robocopy (multi-threaded)..."
        Write-Log $UI "---  File Copy  ---" "Cyan"
        Write-Log $UI "Source -> $isoDrive`:\   Destination -> $usbDrive`:\"
        Write-Log $UI "Excluding install.wim / install.esd  (handled separately)"

        $roboArgs = @(
            "$isoDrive`:\", "$usbDrive`:\",
            "/E",      # all subdirectories including empty
            "/MT:16",  # 16 parallel threads
            "/NJH",    # no job header
            "/NJS",    # no job summary
            "/NDL",    # no directory listing noise
            "/NP",     # no per-file percentage
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

        # BUG 8 FIX: WaitForExit() must be called before ExitCode is
        # reliable on Start-Process -PassThru handles. Without it, ExitCode
        # can return null even after the process has finished.
        $roboJob.WaitForExit()

        # Robocopy exit codes 0-7 are all success variants.
        if ($roboJob.ExitCode -gt 7) {
            throw "Robocopy failed with exit code $($roboJob.ExitCode)."
        }
        Write-Log $UI "[OK]  Files copied successfully." "Success"

        # -- 7. Handle ESD -> WIM conversion --------------------------
        $wimSource = ""
        if ($imgInfo.Type -eq "ESD") {
            if ($script:CancelRequested) { throw "Cancelled." }
            Set-Progress $UI 52 "Converting install.esd -> install.wim..."
            Write-Log $UI "---  ESD -> WIM Conversion  ---" "Cyan"
            Write-Log $UI "This may take 5 - 20 minutes depending on drive speed." "Warn"

            $esdFile = $imgInfo.Path
            $tempWim = "$env:TEMP\rufusps_install.wim"
            if (Test-Path $tempWim) { Remove-Item $tempWim -Force }

            $rawInfo = & dism /Get-WimInfo "/WimFile:$esdFile" 2>&1

            # BUG 9 FIX: $rawInfo is a string[] (one element per output line).
            # Passing an array to [regex]::Matches() coerces it to a single
            # string WITHOUT newlines between elements, causing the "Index :"
            # pattern to span what were originally separate lines and fail to
            # match. Pipe through Out-String first to get a properly
            # newline-delimited string before running the regex.
            $rawInfoText = ($rawInfo | Out-String)
            $indexes = [regex]::Matches($rawInfoText, "Index\s*:\s*(\d+)") |
                       ForEach-Object { $_.Groups[1].Value }

            if ($indexes.Count -eq 0) {
                throw "Could not enumerate indexes in install.esd."
            }
            Write-Log $UI "  Found $($indexes.Count) image index(es) in ESD."

            for ($i = 0; $i -lt $indexes.Count; $i++) {
                if ($script:CancelRequested) { throw "Cancelled." }
                $idx     = $indexes[$i]
                $basePct = 52 + [int](($i / $indexes.Count) * 18)
                Write-Log $UI "  Exporting index $idx / $($indexes.Count)..."

                $dismArgs = "/Export-Image " +
                            "/SourceImageFile:`"$esdFile`" " +
                            "/SourceIndex:$idx " +
                            "/DestinationImageFile:`"$tempWim`" " +
                            "/Compress:fast"

                $code    = Invoke-DismWithProgress $UI $dismArgs $basePct 5 "Converting index $idx"
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

        # -- 8. Split or copy WIM -------------------------------------
        if ($wimSource -ne "") {
            if ($script:CancelRequested) { throw "Cancelled." }
            $wimBytes = (Get-Item $wimSource).Length
            $wimGB    = [Math]::Round($wimBytes / 1GB, 2)
            Write-Log $UI "---  Install Image  ---" "Cyan"
            Write-Log $UI "  Path: $wimSource"
            Write-Log $UI "  Size: $wimGB GB"

            $srcDir = "$usbDrive`:\sources"
            if (-not (Test-Path $srcDir)) { New-Item $srcDir -ItemType Directory | Out-Null }

            if ((-not $UseNTFS) -and ($wimBytes -gt 4GB)) {
                Set-Progress $UI 72 "Splitting install.wim for FAT32 compatibility..."
                Write-Log $UI "WIM > 4 GB on FAT32 - splitting into .swm files..." "Warn"

                $swmOut   = "$srcDir\install.swm"
                $dismArgs = "/Split-Image " +
                            "/ImageFile:`"$wimSource`" " +
                            "/SWMFile:`"$swmOut`" " +
                            "/FileSize:4096"

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
                Set-Progress $UI 72 "Copying install.wim to USB..."
                Write-Log $UI "Copying install.wim..."
                Copy-Item $wimSource "$srcDir\install.wim" -Force
                Write-Log $UI "[OK]  install.wim copied." "Success"
            }
        }

        if ($tempWim -and (Test-Path $tempWim)) {
            Remove-Item $tempWim -Force -ErrorAction SilentlyContinue
            Write-Log $UI "Temporary WIM removed." "Muted"
        }

        # -- 9. Boot setup --------------------------------------------
        if ($script:CancelRequested) { throw "Cancelled." }
        Set-Progress $UI 90 "Configuring bootloader..."
        Write-Log $UI "---  Boot Configuration  ---" "Cyan"

        # STEP 9a - EFI file verification -----------------------------
        Write-Log $UI "--- Verifying EFI boot files ---" "Cyan"

        $efiDirs = @(
            "$usbDrive`:\EFI\Boot",
            "$usbDrive`:\EFI\Microsoft\Boot",
            "$usbDrive`:\boot"
        )
        foreach ($d in $efiDirs) {
            if (-not (Test-Path $d)) {
                New-Item $d -ItemType Directory -Force | Out-Null
                Write-Log $UI "  Created dir: $d" "Muted"
            }
        }

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

        # STEP 9b - BCD store -----------------------------------------
        Write-Log $UI "--- BCD store verification ---" "Cyan"

        $bcdSrc     = "$isoDrive`:\efi\microsoft\boot\bcd"
        $bcdDst     = "$usbDrive`:\EFI\Microsoft\Boot\BCD"
        $bcdBoot    = "$usbDrive`:\boot\BCD"
        $bcdBootSrc = "$isoDrive`:\boot\bcd"

        if (Test-Path $bcdSrc) {
            Copy-Item $bcdSrc $bcdDst -Force
            Write-Log $UI "  [OK] EFI BCD copied from ISO." "Success"
        } else {
            Write-Log $UI "  [!] EFI BCD not found in ISO - this is unusual." "Warn"
        }

        if (Test-Path $bcdBootSrc) {
            Copy-Item $bcdBootSrc $bcdBoot -Force
            Write-Log $UI "  [OK] BIOS BCD copied from ISO." "Success"
        }

        Write-Log $UI "  Syncing \\boot\\ folder..." "Muted"
        $bootRobo = Start-Process robocopy `
            -ArgumentList @("$isoDrive`:\boot", "$usbDrive`:\boot", "/E", "/MT:8", "/NJH", "/NJS", "/NP") `
            -PassThru -NoNewWindow -Wait
        if ($bootRobo.ExitCode -le 7) {
            Write-Log $UI "  [OK] \\boot\\ folder synced." "Success"
        }

        # STEP 9c - BIOS boot sector ----------------------------------
        if (-not $UseGPT) {
            Write-Log $UI "--- Writing BIOS boot sector (bootsect) ---" "Cyan"

            $bootsectPaths = @(
                "$usbDrive`:\boot\bootsect.exe",
                "$isoDrive`:\boot\bootsect.exe"
            )
            $bootsectExe = $bootsectPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

            if ($bootsectExe) {
                $bsOut = & "$bootsectExe" /nt60 "$usbDrive`:" /force /mbr 2>&1
                $bsOut | Where-Object { $_.Trim() -ne "" } | ForEach-Object {
                    Write-Log $UI "  bootsect: $_" "Muted"
                }
                if ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE) {
                    Write-Log $UI "  [OK] VBR + MBR boot code written." "Success"
                } else {
                    Write-Log $UI "  [!] bootsect exit code: $LASTEXITCODE" "Warn"
                }
            } else {
                Write-Log $UI "  [!] bootsect.exe not found - BIOS boot may not work." "Warn"
                Write-Log $UI "      Try running: bootsect /nt60 $usbDrive`: /force /mbr" "Warn"
            }
        }

        # STEP 9d - GPT ESP attributes --------------------------------
        if ($UseGPT) {
            Write-Log $UI "--- Setting GPT ESP attributes via diskpart ---" "Cyan"
            $diskNum = $DiskObj.Number
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

        # STEP 9e - Final sanity check --------------------------------
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
            Write-Log $UI "    This usually means the ISO is non-standard (e.g. Tiny11 stripped build)." "Warn"
        }

        # -- 10. Done -------------------------------------------------
        Set-Progress $UI 100 "[OK]  USB creation complete!"
        Write-Log $UI ""
        Write-Log $UI "+==========================================+" "Success"
        Write-Log $UI "|   [OK]  USB IS READY!                   |" "Success"
        Write-Log $UI "|   Drive      : $usbDrive`:\              |" "Success"
        Write-Log $UI "|   Partition  : $(if ($UseGPT) {'GPT (UEFI only)'} else {'MBR (UEFI + BIOS)'})   |" "Success"
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
        if ($mountISO) {
            try { Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue } catch {}
            Write-Log $UI "ISO dismounted." "Muted"
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

# BUG 13 FIX: Drag & Drop events on the ISO textbox were never wired up,
# so dropping a file onto the text box had no effect whatsoever.
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

    # BUG 14 FIX: Dialog text ended with "Continue..." instead of "Continue?"
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
