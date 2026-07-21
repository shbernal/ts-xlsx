<#
.SYNOPSIS
  Open a .xlsx in a VISIBLE, INTERACTIVE desktop Excel and report the verdict
  the headless COM oracle cannot see: does Excel open it clean, prompt to
  repair, auto-repair, warn on a format mismatch, or refuse it outright.

.DESCRIPTION
  tools/excel-oracle/observe.ps1 opens files with DisplayAlerts=$false, which
  SUPPRESSES the modal repair dialog - it can infer a repair happened (Open
  throws, or the name carries "[Repaired]") but never reproduces the dialog a
  real user sees. This script is the interactive counterpart: it launches Excel
  with /x (a separate process, for clean PID isolation), watches that PID's own
  top-level windows via UI Automation, classifies the first decisive state,
  screenshots it, and emits ONE JSON verdict blob on stdout.

  It is a PROBE, not a test (same contract as excel-oracle): its output is a
  recorded fact that SEEDS a corpus case. Never wire it into CI - it needs a
  licensed Excel, an interactive desktop session, and is inherently slower and
  less deterministic than the Node corpus runner.

  By default it is READ-ONLY: on a repair/mismatch prompt it clicks the safe
  negative ('No') so nothing is rewritten. Pass -AcceptRepair to click 'Yes',
  capture the post-repair notification + any repair log Excel writes, and
  (with -SaveRepairedTo) persist Excel's recovered canonical output as a fixture.

.PARAMETER Path
  The .xlsx to open.

.PARAMETER AcceptRepair
  Click 'Yes' on a repair/format-mismatch prompt instead of 'No'. Captures the
  [Repaired] workbook title, the post-repair notification text, and scans for a
  repair-log file. Off by default (read-only).

.PARAMETER SaveRepairedTo
  With -AcceptRepair, save the recovered workbook to this path so you can commit
  Excel's own canonicalization of the repaired content as a fixture.

.PARAMETER TimeoutSec
  Max seconds to wait for a decisive window state before giving up. Default 40.

.PARAMETER CloseAfter
  Quit the Excel instance this run spawned when done. Off by default so you can
  inspect the window; when off, remember it leaves an EXCEL.EXE running.

.PARAMETER Shot
  Screenshot output path (defaults to a temp file - never the repo/skill dir).

.EXAMPLE
  # Read-only verdict on a suspect file
  & open-verdict.ps1 -Path .\test\corpus\fixtures\suspect.xlsx -Shot .tmp\verdict.png

.EXAMPLE
  # Accept the repair, keep Excel's recovered output, capture the log
  & open-verdict.ps1 -Path .\broken.xlsx -AcceptRepair -SaveRepairedTo .\recovered.xlsx
#>
param(
  [Parameter(Mandatory)][string]$Path,
  [switch]$AcceptRepair,
  [string]$SaveRepairedTo = '',
  [int]$TimeoutSec = 40,
  [switch]$CloseAfter,
  [string]$Shot = (Join-Path ([IO.Path]::GetTempPath()) 'xl-open-verdict.png')
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\xl-window-lib.ps1"
. "$PSScriptRoot\uia-lib.ps1"

$Path = (Resolve-Path -LiteralPath $Path).Path
$stem = [IO.Path]::GetFileNameWithoutExtension($Path)
$fileDir = [IO.Path]::GetDirectoryName($Path)

# Classification patterns (case-insensitive substring on a window's combined
# visible text). These are Excel's English-locale strings; on first real run,
# -Dump the actual dialog text and widen these if the wording differs by build.
$reRepairPrompt = 'we found a problem with some content|recover as much as we can'
$reMismatch     = "file format and extension of.*don't match|format and extension .* don't match"
$reRejected     = 'cannot open the file|file format or file extension is not valid|is corrupt and cannot be opened|unable to read (the )?file'
$rePostRepair   = 'repaired or removed the unreadable content|repaired records|excel was able to open the file'

# Only ever touch an EXCEL.EXE this run spawned.
$preExisting = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
if ($preExisting.Count -gt 0) {
  Write-Host "WARN: $($preExisting.Count) EXCEL.EXE already running; using /x to spawn an isolated instance."
}

# /x = separate process (clean PID isolation). Quote the path as one argument.
Start-Process -FilePath 'excel.exe' -ArgumentList '/x', "`"$Path`"" | Out-Null

$verdict = [ordered]@{
  path         = $Path
  verdict      = 'unknown'   # clean | repaired | repair-prompt | format-mismatch | rejected | timeout
  windowTitle  = $null
  dialogText   = $null
  accepted     = [bool]$AcceptRepair
  repairLog    = @()
  repairedPath = $null
  spawnedPid   = $null
  screenshot   = $Shot
}

function Get-SpawnedExcelPids {
  @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) |
    Where-Object { $preExisting -notcontains $_ }
}

# Poll for the first decisive window state.
$deadline = (Get-Date).AddSeconds($TimeoutSec)
$decided = $false
$cleanStreak = 0
while ((Get-Date) -lt $deadline -and -not $decided) {
  Start-Sleep -Milliseconds 600
  $pids = Get-SpawnedExcelPids
  if (-not $pids) { continue }
  $verdict.spawnedPid = $pids[0]

  foreach ($xlPid in $pids) {
    foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
      $title = $w.Current.Name
      $text = Get-UiaVisibleText -Scope $w

      if ($text -imatch $reRepairPrompt) {
        $verdict.verdict = 'repair-prompt'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      if ($text -imatch $reMismatch) {
        $verdict.verdict = 'format-mismatch'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      if ($text -imatch $reRejected) {
        $verdict.verdict = 'rejected'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      # A workbook frame window (title contains the file stem) with no blocking
      # dialog. Require it to persist a couple of polls so a transient splash
      # doesn't read as 'clean'.
      if ($title -and $title -like "*$stem*") {
        if ($title -match '\[Repaired\]') {
          $verdict.verdict = 'repaired'; $verdict.windowTitle = $title; $decided = $true; break
        }
        $cleanStreak++
        if ($cleanStreak -ge 2) { $verdict.verdict = 'clean'; $verdict.windowTitle = $title; $decided = $true; break }
      }
    }
    if ($decided) { break }
  }
}
if (-not $decided) { $verdict.verdict = 'timeout' }

# Foreground + screenshot the spawned instance for the visual record.
try {
  if ($verdict.spawnedPid) {
    $hwnd = Set-ExcelForeground -TargetPid $verdict.spawnedPid
    $verdict.screenshot = $Shot
    Save-WindowScreenshot -Hwnd $hwnd -Path $Shot | Out-Null
  }
} catch { Write-Host "SCREENSHOT SKIPPED: $($_.Exception.Message)" }

# Resolve a prompt.
if ($verdict.verdict -in @('repair-prompt', 'format-mismatch')) {
  $pids = Get-SpawnedExcelPids
  $scope = $null
  foreach ($xlPid in $pids) {
    foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
      $t = Get-UiaVisibleText -Scope $w
      if ($t -imatch $reRepairPrompt -or $t -imatch $reMismatch) { $scope = $w; break }
    }
    if ($scope) { break }
  }
  if ($scope) {
    $btn = if ($AcceptRepair) { 'Yes' } else { 'No' }
    Invoke-UiaElement -Name $btn -ControlType $script:CT::Button -Scope $scope | Out-Null
    Start-Sleep -Milliseconds 1500

    if ($AcceptRepair) {
      # Post-repair notification (if any) + [Repaired] title.
      $before = @(Get-ChildItem -Path $fileDir -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
      foreach ($xlPid in (Get-SpawnedExcelPids)) {
        foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
          $t = Get-UiaVisibleText -Scope $w
          if ($t -imatch $rePostRepair) { $verdict.dialogText = $t }
          if ($w.Current.Name -match '\[Repaired\]') { $verdict.windowTitle = $w.Current.Name; $verdict.verdict = 'repaired' }
          # Dismiss the notification if it has a Close/OK.
          Invoke-UiaElement -Name 'Close' -ControlType $script:CT::Button -Scope $w | Out-Null
          Invoke-UiaElement -Name 'OK' -ControlType $script:CT::Button -Scope $w | Out-Null
        }
      }
      if ($SaveRepairedTo) {
        try {
          $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
          $wb = $null
          for ($i = 1; $i -le $xl.Workbooks.Count; $i++) {
            if ($xl.Workbooks.Item($i).Name -like "*$stem*") { $wb = $xl.Workbooks.Item($i); break }
          }
          if ($wb) { $wb.SaveAs($SaveRepairedTo, 51); $verdict.repairedPath = $SaveRepairedTo; Write-Host "SAVED REPAIRED: $SaveRepairedTo" }
        } catch { Write-Host "SAVE-REPAIRED SKIPPED: $($_.Exception.Message)" }
      }
      # Excel writes error*.xml (the repair log) next to the file when the
      # repaired copy is materialized; report any that appeared.
      $after = @(Get-ChildItem -Path $fileDir -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
      $verdict.repairLog = @($after | Where-Object { $before -notcontains $_ })
    }
  }
}

if ($CloseAfter) {
  foreach ($xlPid in (Get-SpawnedExcelPids)) {
    Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
  }
}

$verdict | ConvertTo-Json -Depth 5
