# Retrieve all DLP policies
$allPolicies = Get-DlpCompliancePolicy

# Filter for policies with 'Pending' distribution status
$pendingPolicies = $allPolicies | Where-Object { $_.DistributionStatus -eq 'Pending' }

# Display results in the console
if ($pendingPolicies) {
    Write-Host "DLP Policies with Pending Distribution Status:" -ForegroundColor Green
    $pendingPolicies | Format-Table Name, DistributionStatus, Mode, Enabled -AutoSize
} else {
    Write-Host "No DLP policies with 'Pending' distribution status found." -ForegroundColor Yellow
}

# Optionally, export the results to a CSV file
$csvFilePath = ".\DlpPolicies_PendingStatus.csv"
$pendingPolicies | Select-Object Name, DistributionStatus, Mode, Enabled | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

Write-Host "Results exported to $csvFilePath" -ForegroundColor Green