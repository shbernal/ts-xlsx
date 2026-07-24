<#
.SYNOPSIS
  Open a macro-enabled workbook with macros ENABLED and report whether its VBA project actually loads
  and RUNS the authored source - the verdict no other oracle can produce.

.DESCRIPTION
  open-verdict.ps1 is structurally BLIND to VBA authoring defects: it dismisses every prompt with the
  safe negative and never enables macros, so it only ever catches package-level (structural) repair,
  never a module that fails to compile - or, worse, one that loads clean but silently runs the wrong
  (stale) p-code - once macros are enabled. This script closes that gap.

  It opens the file on a private, headless Excel COM instance with AutomationSecurity = Low (the
  automation equivalent of a user clicking "Enable Content": the VBA project is loaded and COMPILED on
  open, exactly the step that reveals authoring defects), then - given -RunMacro - calls
  Application.Run(<macro>) to prove a KNOWN macro from the authored source executes and returns what it
  should.

  WHY EXECUTION MATTERS (do not skip -RunMacro): a VBA module can pair valid-but-mismatched p-code with
  new source; Excel then loads clean AND silently runs the stale p-code. "Opens without error" is
  therefore NOT proof of correctness for any VBA-authoring feature. The definitive checks are:
    - a KNOWN authored macro RUNS and returns the expected value  -> the source is live and correct;
    - that same call THROWS "macro may not be available"          -> the module did not load/compile
                                                                      (bad p-code) or Excel is running
                                                                      stale p-code that lacks it.
  (Recorded finding, 2026-07-24; MODULEOFFSET=0 and version-cookie tricks do NOT force a recompile.)

  Because AutomationSecurity=Low + DisplayAlerts=$false suppresses the interactive "Invalid data format"
  modal a user would see, brokenness here surfaces as an Open exception or a Run failure, NOT a dialog -
  so -RunMacro (a known authored macro) is the definitive signal, not the absence of a popup.

  Runs under pwsh 7 or Windows PowerShell: it only needs New-Object -ComObject (no GetActiveObject, no
  UI Automation). PROBE, not a test (same contract as open-verdict.ps1 / tools/excel-oracle): its output
  is a recorded fact that seeds a corpus case. Never wire it into CI.

.PARAMETER Path
  The .xlsm/.xltm to open with macros enabled.

.PARAMETER RunMacro
  Macro to Application.Run after opening (e.g. 'AddThem' or 'Module1.AddThem'). Choose one from the
  AUTHORED source with a deterministic return value and NO UI (a plain Function, not a Sub that pops a
  MsgBox - a MsgBox blocks Run indefinitely and hangs the probe).

.PARAMETER RunArgs
  Arguments passed to the macro, in order (up to 3).

.PARAMETER ExpectResult
  Optional expected return value; when given, the verdict's `passed` is true only if the macro ran AND
  its result equals this (string-compared).

.PARAMETER ReadModule
  Optional module name; if given, its CodeModule.Lines are read back as Excel sees them - a direct way
  to confirm the source Excel holds matches what was authored.

.EXAMPLE
  pwsh -File execute-verdict.ps1 -Path .\authored.xlsm -RunMacro AddThem -RunArgs 4,5 -ExpectResult 9
#>
param(
  [Parameter(Mandatory)][string]$Path,
  [string]$RunMacro = '',
  [object[]]$RunArgs = @(),
  [string]$ExpectResult = '',
  [string]$ReadModule = ''
)
$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path

$verdict = [ordered]@{
  path        = $resolved
  openThrew   = $false
  openError   = $null
  ran         = $false
  runResult   = $null
  runError    = $null
  moduleLines = $null
  expected    = if ($ExpectResult -ne '') { $ExpectResult } else { $null }
  passed      = $null       # true/false only when RunMacro (and, if given, ExpectResult) can be judged
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
# COM-created Excel defaults AutomationSecurity to ForceDisable, which blocks macros BEFORE the VBA
# engine loads - masking authoring defects. Low = "as if the user enabled content": load + compile now.
$xl.AutomationSecurity = 1
$wb = $null
try {
  try {
    $wb = $xl.Workbooks.Open($resolved)
  } catch {
    $verdict.openThrew = $true
    $verdict.openError = $_.Exception.Message
  }

  if ($wb -and $ReadModule) {
    try {
      $comp = $wb.VBProject.VBComponents.Item($ReadModule)
      $verdict.moduleLines = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    } catch { $verdict.moduleLines = "READ FAILED: $($_.Exception.Message)" }
  }

  if ($wb -and $RunMacro) {
    try {
      $a = @($RunArgs)
      switch ($a.Count) {
        0 { $verdict.runResult = $xl.Run($RunMacro) }
        1 { $verdict.runResult = $xl.Run($RunMacro, $a[0]) }
        2 { $verdict.runResult = $xl.Run($RunMacro, $a[0], $a[1]) }
        3 { $verdict.runResult = $xl.Run($RunMacro, $a[0], $a[1], $a[2]) }
        default { throw "execute-verdict supports up to 3 macro args; got $($a.Count)" }
      }
      $verdict.ran = $true
    } catch { $verdict.runError = $_.Exception.Message }
  }
} finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $xl.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

if ($RunMacro) {
  if ($ExpectResult -ne '') {
    $verdict.passed = ($verdict.ran -and "$($verdict.runResult)" -eq $ExpectResult)
  } else {
    $verdict.passed = $verdict.ran
  }
}

$verdict | ConvertTo-Json -Depth 5
