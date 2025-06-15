 $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$reportDir = "C:\scripts\PingCastle\reports"
$logDir = "C:\scripts\PingCastle\logs"
$transcriptLog = Join-Path $logDir "RunPingCastle-$timestamp.transcript.txt"
$stdoutLog = Join-Path $logDir "PingCastle-$timestamp.stdout.txt"
$stderrLog = Join-Path $logDir "PingCastle-$timestamp.stderr.txt"

# Ensure required folders exist
foreach ($dir in @($reportDir, $logDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

Start-Transcript -Path $transcriptLog -Append

Write-Host "[$(Get-Date)] Starting Run-PingCastle.ps1"
Write-Host "Running as: $env:USERNAME"
Write-Host "Setting current directory to: $reportDir"

# Set working directory so PingCastle drops files in the right place
Set-Location -Path $reportDir

Write-Host "Launching PingCastle..."
$process = Start-Process -FilePath "C:\scripts\PingCastle\PingCastle.exe" `
  -ArgumentList @(
      "--healthcheck",
      "--level", "Full",
      "--no-enum-limit",
      "--datefile"
  ) `
  -RedirectStandardOutput $stdoutLog `
  -RedirectStandardError $stderrLog `
  -NoNewWindow `
  -Wait `
  -PassThru

Write-Host "PingCastle exited with code: $($process.ExitCode)"

# List all newly written report files (optional)
$outputFiles = Get-ChildItem -Path $reportDir -Filter "*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 3
if ($outputFiles) {
    Write-Host "Recent PingCastle report(s):"
    $outputFiles | ForEach-Object { Write-Host " - $($_.FullName)" }
} else {
    Write-Warning "No XML report files found in $reportDir"
}

Stop-Transcript
 
