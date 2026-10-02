# Audit-only startup comparison. Never reads real application data.
$ErrorActionPreference = 'Stop'
$auditRoot = Join-Path $PSScriptRoot 'output\perf-audit-v72'
$env:RX_APP_DATA_DIR = Join-Path $auditRoot 'isolated-data'
$env:RX_PERF_DEBUG = '0'
$results = @()
foreach ($entry in @(
  @{ Name = 'onefile'; Path = (Join-Path $PSScriptRoot 'dist\RxPrescription-v7.2.exe') },
  @{ Name = 'onedir'; Path = (Join-Path $auditRoot 'package-dist\RxPerfAudit\RxPerfAudit.exe') }
)) {
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $process = Start-Process -FilePath $entry.Path -PassThru -WindowStyle Hidden
    $ownedIds = [Collections.Generic.HashSet[int]]::new()
    [void]$ownedIds.Add($process.Id)
    $elapsed = $null
    try {
      while ($watch.Elapsed.TotalSeconds -lt 30) {
        $all = Get-CimInstance Win32_Process
        foreach ($item in $all) {
          if ($ownedIds.Contains([int]$item.ParentProcessId)) { [void]$ownedIds.Add([int]$item.ProcessId) }
        }
        foreach ($ownedId in $ownedIds) {
          $candidate = Get-Process -Id $ownedId -ErrorAction SilentlyContinue
          if ($candidate -and $candidate.MainWindowHandle -ne 0 -and $candidate.Responding) {
            $elapsed = [math]::Round($watch.Elapsed.TotalMilliseconds, 3)
            break
          }
        }
        if ($null -ne $elapsed) { break }
        Start-Sleep -Milliseconds 100
      }
      $results += [pscustomobject]@{ package = $entry.Name; attempt = $attempt; window_detected_ms = $elapsed }
      Write-Output "$($entry.Name) attempt $attempt : $elapsed ms"
    } finally {
      foreach ($ownedId in $ownedIds) { Stop-Process -Id $ownedId -Force -ErrorAction SilentlyContinue }
    }
  }
}
$results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $auditRoot 'packaging.json') -Encoding utf8
# Window discovery is a coarse proxy, not proof that every page is usable.
