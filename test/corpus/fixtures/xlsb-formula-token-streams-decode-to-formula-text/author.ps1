# Authors this case's two fixtures with real Excel: `source.xlsb` (binary BIFF12, formulas stored as
# `Ptg` token streams) and `source.xlsx` (OOXML XML, formulas stored as text), saved from one and the
# same in-memory workbook. The XML twin is the oracle: whatever text Excel writes into `<f>` is what
# the binary token stream must decode back to, character for character.
#
# The workbook is deliberately a grammar tour rather than a realistic sheet — one formula per token
# class the decoder must handle (operators and their precedence, every operand kind, every reference
# shape including 3-D and whole-column/row, defined names, shared and array formulas).
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/xlsb-formula-token-streams-decode-to-formula-text/author.ps1

$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
$xlsb = Join-Path $outDir 'source.xlsb'
$xlsx = Join-Path $outDir 'source.xlsx'
Remove-Item -Force -ErrorAction SilentlyContinue $xlsb, $xlsx

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

  $calc = $wb.Worksheets.Item(1)
  $calc.Name = 'Calc'

  # The other sheets exist to be referenced: `Data` and `More` are adjacent so a 3-D span can name
  # them as a range, and `Odd Name` forces the single-quoted sheet-name spelling.
  $data = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $calc)
  $data.Name = 'Data'
  $more = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $data)
  $more.Name = 'More'
  $odd = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $more)
  $odd.Name = 'Odd Name'

  1..5 | ForEach-Object { $calc.Cells.Item($_, 1).Value2 = $_ }
  $calc.Range('E1').Value2 = 0.25
  $data.Range('A1').Value2 = 10
  $data.Range('B1').Value2 = 20
  $data.Range('A2').Value2 = 30
  $data.Range('B2').Value2 = 40
  $more.Range('A1').Value2 = 7
  $odd.Range('A1').Value2 = 9

  # Defined names: one pointing at a cell, one holding a constant, one scoped to a single sheet.
  $wb.Names.Add('Rate', '=Calc!$E$1') | Out-Null
  $wb.Names.Add('Factor', '=2') | Out-Null
  $calc.Names.Add('Local', '=Calc!$A$1') | Out-Null

  # One formula per token class. Column B labels what column C is exercising, so a failing assertion
  # can be read straight off the sheet.
  $cases = [ordered]@{
    'int literals and precedence'   = '=1+2*3'
    'parenthesised subexpression'   = '=(1+2)*3'
    'float literal'                 = '=1.5+2'
    'unary minus'                   = '=-A1'
    'percent postfix'               = '=A1%'
    'exponent'                      = '=A2^3'
    'string concatenation'          = '=A1&" x"'
    'string with doubled quote'     = '="say ""hi"""'
    'comparison'                    = '=A1<=A2'
    'not equal'                     = '=A1<>A2'
    'boolean literal'               = '=TRUE'
    'error literal'                 = '=#N/A'
    'mixed absolute reference'      = '=$A$1+A$2+$A3'
    'range argument'                = '=SUM(A1:A5)'
    'nested calls'                  = '=SUM(A1:A5)/COUNT(A1:A5)'
    'multi-argument call'           = '=IF(A1>0,"pos","neg")'
    'argument list with a range'    = '=SUM(A1:A2,A4)'
    'zero-argument call'            = '=PI()'
    'cross-sheet range'             = '=SUM(Data!A1:B2)'
    'cross-sheet single cell'       = '=Data!$A$1'
    'quoted sheet name'             = '=''Odd Name''!A1'
    'three-dimensional span'        = '=SUM(Data:More!A1)'
    'whole column'                  = '=SUM(A:A)'
    'whole row on another sheet'    = '=SUM(Data!2:2)'
    'defined name'                  = '=Rate*A1'
    'defined name holding constant' = '=Factor*A2'
    'sheet-scoped defined name'     = '=Local+1'
    'range intersection'            = '=SUM(A1:A3 A2:A5)'
    'future function'               = '=TEXTJOIN(",",TRUE,A1:A3)'
    'parenthesised union'           = '=SUM((A1:A2,A4:A5))'
    'array constant'                = '=SUM({1,2;3,4})'
    'non-square array constant'     = '=SUM({1,2,3;4,5,6})'
    'mixed-type array constant'     = '=COUNTA({"a",TRUE;#N/A,5})'
    'omitted argument'              = '=IF(A1>0,,1)'
    'reference error'               = '=#REF!+1'
  }
  $row = 0
  foreach ($label in $cases.Keys) {
    $row++
    $calc.Cells.Item($row, 2).Value2 = [string]$label
    $calc.Cells.Item($row, 3).Formula = [string]$cases[$label]
  }

  # A shared formula: one formula filled down a column, which Excel stores once and clones.
  $calc.Range('D1:D5').Formula = '=A1*2'
  # An array (CSE) formula over a range.
  $calc.Range('F1').FormulaArray = '=SUM(A1:A5*2)'

  $wb.Worksheets.Item('Calc').Activate()
  $wb.SaveAs($xlsb, 50)   # xlExcel12
  $wb.SaveAs($xlsx, 51)   # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsb, $xlsx | Select-Object Name, Length
