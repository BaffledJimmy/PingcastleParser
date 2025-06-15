 param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Slack", "MSTeams")]
    [string]$Method
)

function Get-SecureWebhook {
    $path = "C:\Scripts\PingCastleWebhook.txt"
    if (-not (Test-Path $path)) {
        throw "Webhook credential file not found at $path"
    }
    $secure = Get-Content $path | ConvertTo-SecureString
    return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    )
}

try {
    $webhookUrl = Get-SecureWebhook
    $reportPath = "C:\scripts\PingCastle\reports"
    $files = Get-ChildItem -Path $reportPath -Filter "*.xml" | Sort-Object LastWriteTime -Descending

    if ($files.Count -lt 2) {
        Write-Host "Not enough PingCastle XML files to compare."
        exit 1
    }

    $latest = [xml](Get-Content $files[0].FullName)
    $previous = [xml](Get-Content $files[1].FullName)

    $latestScore = [int]$latest.HealthcheckData.GlobalScore
    $previousScore = [int]$previous.HealthcheckData.GlobalScore
    $scoreChange = $latestScore - $previousScore

    $latestRisks = $latest.HealthcheckData.RiskRules.HealthcheckRiskRule
    $previousRisks = $previous.HealthcheckData.RiskRules.HealthcheckRiskRule

    $prevScores = @{}
    foreach ($r in $previousRisks) {
        $prevScores[$r.RiskId] = [int]$r.Points
    }

    $delta = @()
    foreach ($r in $latestRisks) {
        $prev = if ($prevScores.ContainsKey($r.RiskId)) { $prevScores[$r.RiskId] } else { 0 }
        $diff = [int]$r.Points - $prev
        if ($diff -ne 0) {
            $delta += [PSCustomObject]@{
                RiskId    = $r.RiskId
                Change    = $diff
                Rationale = $r.Rationale
            }
        }
    }

    $top3 = $delta | Sort-Object -Property { [math]::Abs($_.Change) } -Descending | Select-Object -First 3
 $reasonsText = ( $top3 | ForEach-Object {
        $sign = if ($_.Change -gt 0) { "+" } else { "-" }
        "* $($_.Rationale) (${sign}$($_.Change))"
    }
) -join "`n"


   switch ($true) {
    ($scoreChange -gt 0) { $summary = ":x: Risk score has *increased*" }
    ($scoreChange -lt 0) { $summary = ":white_check_mark: Risk score has *decreased*" }
    default              { $summary = ":pause_button: Risk score is *unchanged*" }
}


    $htmlLink = $files[0].FullName -replace ".xml$", ".html"
    $htmlLink = "file:///" + $htmlLink.Replace('\', '/')

 if ($Method -eq "Slack") {
    $messageLines = @()
    $messageLines += "*PingCastle Risk Score Update*"
    $messageLines += ""
    $messageLines += "$summary : $previousScore -> $latestScore ($scoreChange)"
    $messageLines += ""
    $messageLines += "*Top 3 Score Changes:*"
    $messageLines += $reasonsText
    $messageLines += ""
    $messageLines += "Report: <$htmlLink>"

    $payload = @{ text = ($messageLines -join "`n") } | ConvertTo-Json -Depth 3
}

    elseif ($Method -eq "MSTeams") {
        $payload = @{
            "@type"      = "MessageCard"
            "@context"   = "http://schema.org/extensions"
            "summary"    = "PingCastle Risk Score Report"
            "themeColor" = if ($scoreChange -gt 0) { "ff0000" } elseif ($scoreChange -lt 0) { "00cc66" } else { "cccccc" }
            "sections"   = @(
                @{
                    "activityTitle" = "**$summary**"
                    "text"          = "Previous: ${previousScore}, Current: ${latestScore}, Change: ${scoreChange}"
                },
                @{
                    "title" = "Top 3 Score Changes"
                    "text"  = $reasonsText
                },
                @{
                    "title" = "Report Link"
                    "text"  = "<a href='$htmlLink'>Open HTML Report</a>"
                }
            )
        } | ConvertTo-Json -Depth 10
    }

    Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $payload -ContentType 'application/json'
}
catch {
    Write-Error "Failed to send PingCastle score report: $_"
    exit 2
}
 
