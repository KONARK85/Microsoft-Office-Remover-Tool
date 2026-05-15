#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Comprehensive Microsoft Office Removal Script

.DESCRIPTION
    Safely removes all versions and editions of Microsoft Office from a Windows PC,
    including Office 365, Microsoft 365, Office 2021, 2019, 2016, 2013, 2010, 2007,
    and standalone apps like Visio, Project, and OneNote.

.NOTES
    - Must be run as Administrator
    - Creates a restore point before making changes
    - Logs all actions to C:\OfficeRemoval\removal_log.txt
    - Reboot may be required after completion
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
$LogDir  = "C:\OfficeRemoval"
$LogFile = "$LogDir\removal_log.txt"
$SaraUrl = "https://aka.ms/SaRA_CommandLineBeta"   # Microsoft Support and Recovery Assistant

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line -ForegroundColor Cyan }
    }
}

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "  ║       Microsoft Office Complete Removal Tool         ║" -ForegroundColor Magenta
    Write-Host "  ║   Supports: 2007 / 2010 / 2013 / 2016 / 2019 /      ║" -ForegroundColor Magenta
    Write-Host "  ║             2021 / 365 / M365 + Visio / Project      ║" -ForegroundColor Magenta
    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""
}

function Confirm-Action {
    param([string]$Prompt)
    $ans = Read-Host "$Prompt [Y/N]"
    return $ans -match '^[Yy]'
}

function New-SystemRestorePoint {
    Write-Log "Creating system restore point..."
    try {
        Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue
        Checkpoint-Computer -Description "Before Office Removal - $(Get-Date -Format 'yyyyMMdd_HHmmss')" `
                            -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-Log "Restore point created successfully." "SUCCESS"
    } catch {
        Write-Log "Could not create restore point: $_" "WARN"
    }
}

# ─────────────────────────────────────────────
# STEP 1 – Stop Office Processes
# ─────────────────────────────────────────────
function Stop-OfficeProcesses {
    Write-Log "Stopping Office-related processes..."
    $procs = @(
        "WINWORD","EXCEL","POWERPNT","OUTLOOK","ONENOTE","MSPUB",
        "MSACCESS","LYNC","GROOVE","INFOPATH","VISIO","WINPROJ",
        "MSOUC","OfficeClickToRun","AppVShNotify","officec2rclient",
        "Teams","OneDrive"
    )
    foreach ($p in $procs) {
        Get-Process -Name $p -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Office processes stopped." "SUCCESS"
}

# ─────────────────────────────────────────────
# STEP 2 – Uninstall via Windows Installer (MSI)
# ─────────────────────────────────────────────
function Remove-OfficeViaMsi {
    Write-Log "Scanning for MSI-based Office installations..."

    $officePattern = "Microsoft Office|Microsoft 365|Microsoft Visio|Microsoft Project|" +
                     "Microsoft OneNote|Microsoft Teams|Microsoft OneDrive"

    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $found = @()
    foreach ($path in $regPaths) {
        $found += Get-ItemProperty $path -ErrorAction SilentlyContinue |
                  Where-Object { $_.DisplayName -match $officePattern -and $_.UninstallString }
    }

    if ($found.Count -eq 0) {
        Write-Log "No MSI-based Office products found." "WARN"
        return
    }

    Write-Log "Found $($found.Count) MSI Office product(s)."
    foreach ($pkg in $found) {
        Write-Log "  Removing: $($pkg.DisplayName)"
        try {
            if ($pkg.UninstallString -match "MsiExec|msiexec") {
                # Extract product code
                $prodCode = $pkg.UninstallString -replace '.*({[A-Z0-9\-]+}).*','$1'
                if ($prodCode -match '^\{') {
                    $args = "/x $prodCode /qn /norestart REBOOT=ReallySuppress"
                    Start-Process "msiexec.exe" -ArgumentList $args -Wait -NoNewWindow
                } else {
                    # Fallback: run uninstall string silently
                    $cmd = $pkg.UninstallString + " /qn /norestart"
                    Start-Process "cmd.exe" -ArgumentList "/c $cmd" -Wait -NoNewWindow
                }
                Write-Log "  Removed: $($pkg.DisplayName)" "SUCCESS"
            }
        } catch {
            Write-Log "  Failed to remove $($pkg.DisplayName): $_" "ERROR"
        }
    }
}

# ─────────────────────────────────────────────
# STEP 3 – Uninstall Click-to-Run (C2R)
# ─────────────────────────────────────────────
function Remove-OfficeC2R {
    Write-Log "Checking for Click-to-Run Office installations..."

    $c2rPaths = @(
        "C:\Program Files\Microsoft Office 15\ClientX64\officec2rclient.exe",
        "C:\Program Files\Microsoft Office 15\ClientX86\officec2rclient.exe",
        "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe",
        "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
    )

    $c2rExe = $c2rPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($c2rExe) {
        Write-Log "Found C2R installation at: $c2rExe"
        try {
            # Scenario=Uninstall removes all C2R products
            $args = "scenario=Uninstall DisplayLevel=False AcceptEULA=True"
            Start-Process $c2rExe -ArgumentList $args -Wait -NoNewWindow
            Write-Log "Click-to-Run uninstall completed." "SUCCESS"
        } catch {
            Write-Log "C2R uninstall failed: $_" "ERROR"
        }
    } else {
        Write-Log "No Click-to-Run installation detected." "WARN"
    }

    # Also try via registry-based C2R uninstaller
    $c2rUninstall = Get-ItemProperty `
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*" `
        -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty UninstallString -ErrorAction SilentlyContinue

    if ($c2rUninstall) {
        Write-Log "Running registry C2R uninstaller..."
        try {
            Start-Process "cmd.exe" -ArgumentList "/c $c2rUninstall DisplayLevel=False" `
                          -Wait -NoNewWindow
            Write-Log "Registry C2R uninstall done." "SUCCESS"
        } catch {
            Write-Log "Registry C2R uninstall failed: $_" "ERROR"
        }
    }
}

# ─────────────────────────────────────────────
# STEP 4 – Remove Office Store / UWP Apps
# ─────────────────────────────────────────────
function Remove-OfficeUwpApps {
    Write-Log "Removing Office UWP / Store apps..."

    $uwpPatterns = @(
        "Microsoft.Office*",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.Office.OneNote",
        "Microsoft.MicrosoftTeams",
        "Microsoft.OneDriveSync",
        "Microsoft.OutlookForWindows",
        "Microsoft.Todos"
    )

    foreach ($pat in $uwpPatterns) {
        $pkgs = Get-AppxPackage -AllUsers -Name $pat -ErrorAction SilentlyContinue
        foreach ($pkg in $pkgs) {
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                Write-Log "  Removed UWP: $($pkg.Name)" "SUCCESS"
            } catch {
                Write-Log "  Could not remove UWP $($pkg.Name): $_" "WARN"
            }
        }
    }

    # Remove provisioned packages so they don't reinstall for new users
    foreach ($pat in $uwpPatterns) {
        Get-AppxProvisionedPackage -Online |
            Where-Object { $_.DisplayName -like $pat } |
            ForEach-Object {
                try {
                    Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction Stop
                    Write-Log "  Removed provisioned: $($_.DisplayName)" "SUCCESS"
                } catch {
                    Write-Log "  Could not remove provisioned $($_.DisplayName): $_" "WARN"
                }
            }
    }
}

# ─────────────────────────────────────────────
# STEP 5 – Remove Leftover Files & Folders
# ─────────────────────────────────────────────
function Remove-OfficeFiles {
    Write-Log "Removing leftover Office files and folders..."

    $folders = @(
        "$env:ProgramFiles\Microsoft Office",
        "$env:ProgramFiles\Microsoft Office 15",
        "$env:ProgramFiles\Microsoft Office\root",
        "${env:ProgramFiles(x86)}\Microsoft Office",
        "${env:ProgramFiles(x86)}\Microsoft Office 15",
        "$env:ProgramData\Microsoft\Office",
        "$env:CommonProgramFiles\microsoft shared\Office*",
        "$env:CommonProgramFiles\Microsoft Shared\ClickToRun",
        "${env:CommonProgramFiles(x86)}\microsoft shared\Office*",
        "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun"
    )

    foreach ($folder in $folders) {
        # Expand wildcards
        $resolved = Resolve-Path $folder -ErrorAction SilentlyContinue
        if ($resolved) {
            foreach ($r in $resolved) {
                if (Test-Path $r.Path) {
                    try {
                        Remove-Item -Path $r.Path -Recurse -Force -ErrorAction Stop
                        Write-Log "  Deleted: $($r.Path)" "SUCCESS"
                    } catch {
                        Write-Log "  Could not delete $($r.Path): $_" "WARN"
                    }
                }
            }
        }
    }
}

# ─────────────────────────────────────────────
# STEP 6 – Clean Registry Keys
# ─────────────────────────────────────────────
function Remove-OfficeRegistryKeys {
    Write-Log "Cleaning Office registry keys..."

    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
        "HKCU:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
        "HKLM:\SOFTWARE\Microsoft\ClickToRun",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\ClickToRun",
        "HKLM:\SYSTEM\CurrentControlSet\Services\ClickToRunSvc"
    )

    foreach ($key in $keys) {
        $resolved = Resolve-Path $key -ErrorAction SilentlyContinue
        if ($resolved) {
            foreach ($r in $resolved) {
                try {
                    Remove-Item -Path $r.Path -Recurse -Force -ErrorAction Stop
                    Write-Log "  Deleted reg key: $($r.Path)" "SUCCESS"
                } catch {
                    Write-Log "  Could not delete reg key $($r.Path): $_" "WARN"
                }
            }
        }
    }
}

# ─────────────────────────────────────────────
# STEP 7 – Remove Scheduled Tasks
# ─────────────────────────────────────────────
function Remove-OfficeScheduledTasks {
    Write-Log "Removing Office scheduled tasks..."
    $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
             Where-Object { $_.TaskName -match "Office|OfficeBackgroundTask|OfficeTelemetry|OneDrive" }

    foreach ($task in $tasks) {
        try {
            Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction Stop
            Write-Log "  Removed task: $($task.TaskName)" "SUCCESS"
        } catch {
            Write-Log "  Could not remove task $($task.TaskName): $_" "WARN"
        }
    }
}

# ─────────────────────────────────────────────
# STEP 8 – Remove Start Menu / Desktop Shortcuts
# ─────────────────────────────────────────────
function Remove-OfficeShortcuts {
    Write-Log "Removing Office shortcuts..."

    $shortcutDirs = @(
        "$env:PUBLIC\Desktop",
        "$env:USERPROFILE\Desktop",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    )

    $officeApps = @(
        "Access","Excel","InfoPath","OneNote","Outlook","PowerPoint",
        "Project","Publisher","Visio","Word","Skype for Business","Teams"
    )

    foreach ($dir in $shortcutDirs) {
        if (Test-Path $dir) {
            foreach ($app in $officeApps) {
                Get-ChildItem -Path $dir -Filter "*$app*" -Recurse -ErrorAction SilentlyContinue |
                    Remove-Item -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Write-Log "Shortcuts removed." "SUCCESS"
}

# ─────────────────────────────────────────────
# STEP 9 – Remove ClickToRun Service
# ─────────────────────────────────────────────
function Remove-C2RService {
    Write-Log "Stopping and removing ClickToRun service..."
    $svc = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
    if ($svc) {
        try {
            Stop-Service -Name "ClickToRunSvc" -Force -ErrorAction SilentlyContinue
            sc.exe delete "ClickToRunSvc" | Out-Null
            Write-Log "ClickToRun service removed." "SUCCESS"
        } catch {
            Write-Log "Could not remove ClickToRun service: $_" "WARN"
        }
    } else {
        Write-Log "ClickToRun service not found." "WARN"
    }
}

# ─────────────────────────────────────────────
# STEP 10 – Remove User-Profile Office Data (Optional)
# ─────────────────────────────────────────────
function Remove-OfficeUserData {
    param([bool]$RemoveUserData)
    if (-not $RemoveUserData) {
        Write-Log "Skipping user profile Office data (user chose to keep)." "WARN"
        return
    }

    Write-Log "Removing user-profile Office data..."
    $userPaths = @(
        "$env:APPDATA\Microsoft\Office",
        "$env:APPDATA\Microsoft\Templates",
        "$env:APPDATA\Microsoft\AddIns",
        "$env:LOCALAPPDATA\Microsoft\Office",
        "$env:LOCALAPPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Microsoft\OneDrive"
    )
    foreach ($p in $userPaths) {
        if (Test-Path $p) {
            try {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Log "  Deleted: $p" "SUCCESS"
            } catch {
                Write-Log "  Could not delete $p: $_" "WARN"
            }
        }
    }
}

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
Show-Banner

# Ensure log directory exists
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
Write-Log "=== Office Removal Script Started ==="

# Pre-flight checks
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n  [ERROR] Please re-run this script as Administrator.`n" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  This script will:" -ForegroundColor White
Write-Host "   • Create a system restore point" -ForegroundColor Gray
Write-Host "   • Stop all Office processes" -ForegroundColor Gray
Write-Host "   • Uninstall MSI and Click-to-Run Office editions" -ForegroundColor Gray
Write-Host "   • Remove UWP/Store Office apps" -ForegroundColor Gray
Write-Host "   • Delete leftover files, registry keys, and shortcuts" -ForegroundColor Gray
Write-Host ""

if (-not (Confirm-Action "  Do you want to continue?")) {
    Write-Log "User aborted at confirmation prompt." "WARN"
    Write-Host "`n  Cancelled. No changes were made.`n" -ForegroundColor Yellow
    exit 0
}

$removeUserData = Confirm-Action "`n  Also remove personal Office data (templates, settings, add-ins)?"

Write-Host ""
Write-Log "Starting removal process..."

New-SystemRestorePoint
Stop-OfficeProcesses
Remove-OfficeViaMsi
Remove-OfficeC2R
Remove-OfficeUwpApps
Remove-C2RService
Remove-OfficeFiles
Remove-OfficeRegistryKeys
Remove-OfficeScheduledTasks
Remove-OfficeShortcuts
Remove-OfficeUserData -RemoveUserData $removeUserData

Write-Log "=== Office Removal Script Completed ===" "SUCCESS"

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║          Removal process complete!               ║" -ForegroundColor Green
Write-Host "  ║  Log saved to: C:\OfficeRemoval\removal_log.txt  ║" -ForegroundColor Green
Write-Host "  ║  A reboot is recommended to finish cleanup.      ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""

if (Confirm-Action "  Restart the computer now?") {
    Write-Log "User chose to restart."
    Restart-Computer -Force
} else {
    Write-Host "`n  Please restart manually when convenient.`n" -ForegroundColor Yellow
}
