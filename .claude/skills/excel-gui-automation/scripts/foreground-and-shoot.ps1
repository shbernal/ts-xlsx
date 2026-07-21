<#
.SYNOPSIS
  Foreground the running desktop Excel window and take a window-scoped
  screenshot. Use this as the first sanity check for this skill (confirms a
  visible Excel window exists and can be foregrounded from this shell) and any
  time you just want to "look" at current state without driving an action.
.PARAMETER TargetPid
  Pin to one EXCEL.EXE instance's PID (e.g. the one open-verdict.ps1 spawned)
  so an unrelated Excel window is never foregrounded/captured. 0 = first found.
#>
param(
  [string]$Shot = (Join-Path ([IO.Path]::GetTempPath()) 'xl-gui-shot.png'),
  [int]$TargetPid = 0
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\xl-window-lib.ps1"

$hwnd = Set-ExcelForeground -TargetPid $TargetPid
$size = Save-WindowScreenshot -Hwnd $hwnd -Path $Shot
Write-Host "Foregrounded hwnd=$hwnd; SHOT ($size): $Shot"
