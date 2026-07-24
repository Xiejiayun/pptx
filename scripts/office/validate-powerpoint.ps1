param(
  [Parameter(Mandatory = $true)][string]$Corpus,
  [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $Output | Out-Null
$powerpoint = New-Object -ComObject PowerPoint.Application
$failures = @()
try {
  Get-ChildItem -Path $Corpus -Filter *.pptx -Recurse | ForEach-Object {
    $source = $_.FullName
    try {
      $presentation = $powerpoint.Presentations.Open($source, $true, $false, $false)
      $copy = Join-Path $Output $_.Name
      $presentation.SaveCopyAs($copy)
      $presentation.Close()
    } catch {
      $failures += @{ file = $source; error = $_.Exception.Message }
    }
  }
} finally {
  $powerpoint.Quit()
}
if ($failures.Count -gt 0) {
  $failures | ConvertTo-Json -Depth 4
  exit 1
}
@{ ok = $true; files = (Get-ChildItem -Path $Output -Filter *.pptx).Count } | ConvertTo-Json
