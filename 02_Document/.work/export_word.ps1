param([string]$Source,[string]$Pdf)
$ErrorActionPreference='Stop'
$wordApp=New-Object -ComObject Word.Application
$wordApp.Visible=$false
$wordApp.DisplayAlerts=0
try {
  $doc=$wordApp.Documents.Open($Source,$false,$true)
  $doc.ExportAsFixedFormat($Pdf,17)
  Write-Output ('PAGES '+$doc.ComputeStatistics(2))
  $doc.Close(0)
} finally {$wordApp.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp)|Out-Null}
