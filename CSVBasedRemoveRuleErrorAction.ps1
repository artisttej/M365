# Ask for CSV file
$inputCsv = Read-Host "📄 Enter path to CSV of rules to reset (e.g. exported from Step 1)"

if (-Not (Test-Path $inputCsv)) {
    Write-Host "❌ File not found: $inputCsv" -ForegroundColor Red
    exit
}

# Setup output paths
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logCsv = ".\Dlp_ResetRuleErrorAction_Log_$timestamp.csv"
$htmlPath = ".\Dlp_ResetRuleErrorAction_Report_$timestamp.html"

Write-Host "`n♻️  Resetting RuleErrorAction to '' for rules in: $inputCsv" -ForegroundColor Cyan

$rules = Import-Csv -Path $inputCsv
$logRows = @()

foreach ($rule in $rules) {
    $result = "Skipped"
    $highlight = "<span class='highlight-false'>Unchanged ❌</span>"

try {
    if ($rule.RuleErrorAction -ne "") {
        $params = @{
            Identity = $rule.Identity
            RuleErrorAction = $null
            Confirm = $false
        }
        Set-DlpComplianceRule @params

        $result = "Reset"
        $highlight = "<span class='highlight-true'>Cleared ✅</span>"
        Write-Host "✅ Cleared: $($rule.RuleName)" -ForegroundColor Green
    } else {
        Write-Host "🟡 Already empty: $($rule.RuleName)" -ForegroundColor Yellow
    }
} catch {
    $result = "Failed"
    $highlight = "<span class='highlight-false'>Error ❌</span>"
    Write-Host "❌ Failed: $($rule.RuleName) - $_" -ForegroundColor Red
}


    $logRows += [PSCustomObject]@{
        PolicyName      = $rule.PolicyName
        RuleName        = $rule.RuleName
        Identity        = $rule.Identity
        Result          = $result
        RuleErrorAction = $highlight
    }
}

# Export CSV log
$logRows | Export-Csv -Path $logCsv -NoTypeInformation -Encoding UTF8
Write-Host "📁 Log CSV saved to: $logCsv" -ForegroundColor Green

# HTML style
$htmlStyle = @"
<style>
    body { font-family: Segoe UI; font-size: 14px; margin: 20px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { padding: 8px; border: 1px solid #ccc; }
    th { background-color: #f2f2f2; }
    .highlight-true { color: green; font-weight: bold; }
    .highlight-false { color: red; font-weight: bold; }
</style>
"@

# Build HTML table rows
$htmlRows = foreach ($r in $logRows) {
    "<tr><td>$($r.PolicyName)</td><td>$($r.RuleName)</td><td>$($r.Result)</td><td>$($r.RuleErrorAction)</td></tr>"
}

# Full HTML content
$htmlContent = @"
<html>
<head><title>DLP Reset RuleErrorAction Report</title>$htmlStyle</head>
<body>
<h1>DLP RuleErrorAction Reset Report</h1>
<p>Generated: $(Get-Date)</p>
<table>
<tr><th>Policy Name</th><th>Rule Name</th><th>Result</th><th>RuleErrorAction</th></tr>
$($htmlRows -join "`n")
</table>
</body>
</html>
"@

# Export HTML
$htmlContent | Out-File -Path $htmlPath -Encoding UTF8
Write-Host "📄 HTML report saved to: $htmlPath" -ForegroundColor Green
Write-Host "`n✅ Done resetting RuleErrorAction for all rules listed." -ForegroundColor Green
