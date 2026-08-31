param(
  [Parameter(Mandatory = $true)]
  [string]$DocumentPath
)

$ErrorActionPreference = "Stop"
$resolvedDocument = [System.IO.Path]::GetFullPath($DocumentPath)
$word = $null
try {
  $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
} catch {
  Write-Output "Word is already closed."
  exit 0
}

for ($index = $word.Documents.Count; $index -ge 1; $index -= 1) {
  $document = $word.Documents.Item($index)
  if ([System.IO.Path]::GetFullPath($document.FullName) -eq $resolvedDocument) {
    $document.Close($false)
    Write-Output "Closed disposable document: $resolvedDocument"
  }
}

if ($word.Documents.Count -eq 0) {
  $word.Quit()
  Write-Output "Closed the empty Word instance."
} else {
  Write-Output "Left Word open because other documents are present."
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

