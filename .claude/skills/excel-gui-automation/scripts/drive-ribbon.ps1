<#
.SYNOPSIS
  Drive a desktop-Excel ribbon action end to end: foreground the app, open a
  ribbon command via sequential KeyTips, invoke the resulting popup menu item
  via UI Automation, optionally dump/toggle/invoke dialog controls, and take a
  window-scoped screenshot to confirm the result.

.DESCRIPTION
  For the rare Excel feature with NO COM surface, where a fixture can only be
  authored by driving the real GUI. Reach for this only after confirming the
  feature truly has no COM equivalent - check the real Application/Workbook/
  Worksheet/Shapes object model first (Excel's COM surface is broad; most
  features ARE reachable). This path is slower and less deterministic than COM;
  for ordinary authoring and readback use tools/excel-oracle instead.

  The KeyTip -> UIA-invoke machinery is shared with the verified PowerPoint
  ribbon driver: Office ribbon dropdown popups run on a separate input queue
  from SendKeys and synthetic mouse, so keyboard opens the menu but only
  InvokePattern/TogglePattern reliably activates the item.

.PARAMETER KeyTips
  Semicolon-separated SendKeys tokens sent SEQUENTIALLY (not chorded) to walk
  the ribbon's KeyTip overlay, e.g. '{ESC};{ESC};%;N' = clear stray state, then
  Alt, then N (Insert tab). Discover the letters by screenshotting after each
  Alt press - KeyTips are drawn as small overlay badges on the ribbon.

.PARAMETER UiaMenuItem
  Name of a MenuItem control to Invoke via UI Automation after the KeyTips open
  its popup - for submenu entries keyboard KeyTips do not reach.

.PARAMETER DialogTitleLike
  Wildcard to scope the FOLLOW-ON dialog. Required with -Dump/-Toggle/-InvokeButton
  so lookups don't enumerate the whole desktop (privacy + reliability).

.PARAMETER Dump
  Print the real UIA accessible names of every interactive control in the dialog
  scope. Always run this once before guessing -Toggle/-InvokeButton names.

.PARAMETER Toggle
  Semicolon list of checkbox accessible Names to toggle ON.

.PARAMETER InvokeButton
  Accessible Name of a button to Invoke at the end (e.g. 'OK', 'Insert').

.EXAMPLE
  # Discover Insert-tab KeyTips
  & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%' -Shot .tmp\keytips.png
#>
param(
  [string]$KeyTips = "",
  [int]$KeyTipGapMs = 700,
  [string]$UiaMenuItem = "",
  [string]$DialogTitleLike = "",
  [switch]$Dump,
  [string]$Toggle = "",
  [string]$InvokeButton = "",
  [int]$SettleMs = 1300,
  [int]$TargetPid = 0,
  [string]$Shot = (Join-Path ([IO.Path]::GetTempPath()) 'xl-drive-ribbon.png')
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\xl-window-lib.ps1"

$hwnd = Set-ExcelForeground -TargetPid $TargetPid
Write-Host "Foregrounded Excel hwnd=$hwnd"

if ($KeyTips) {
  Send-KeyTipSequence -Keys ($KeyTips -split ';') -GapMs $KeyTipGapMs
}

if ($UiaMenuItem -or $Dump -or $Toggle -or $InvokeButton) {
  . "$PSScriptRoot\uia-lib.ps1"

  if ($UiaMenuItem) {
    Start-Sleep -Milliseconds 500
    Invoke-UiaElement -Name $UiaMenuItem -ControlType $script:CT::MenuItem | Out-Null
    Start-Sleep -Milliseconds 1500
  }

  $scope = $null
  if ($DialogTitleLike) {
    $dlg = Find-UiaDialog -TitleLike $DialogTitleLike
    if ($dlg) { Write-Host "DIALOG SCOPE: '$($dlg.Current.Name)'"; $scope = $dlg }
    else { Write-Host "DIALOG SCOPE: '$DialogTitleLike' NOT FOUND - falling back to full desktop (unscoped)" }
  }

  if ($Dump) { Get-UiaControlDump -Scope $scope }

  if ($Toggle) {
    foreach ($nm in ($Toggle -split ';')) {
      if ($nm.Trim()) { Set-UiaToggleOn -Name $nm.Trim() -Scope $scope | Out-Null; Start-Sleep -Milliseconds 250 }
    }
  }

  if ($InvokeButton) {
    Invoke-UiaElement -Name $InvokeButton -ControlType $script:CT::Button -Scope $scope | Out-Null
  }
}

Start-Sleep -Milliseconds $SettleMs
$size = Save-WindowScreenshot -Hwnd $hwnd -Path $Shot
Write-Host "SHOT ($size): $Shot"
