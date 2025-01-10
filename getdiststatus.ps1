# Retrieve all policy names
$allPolicyNames = Get-DlpCompliancePolicy | Select-Object -ExpandProperty Name

# Initialize an array to store policy distribution details
$policyDetails = @()

# Iterate through each policy name and fetch accurate distribution status using -DistributionDetail
foreach ($policyName in $allPolicyNames) {
    try {
        $policy = Get-DlpCompliancePolicy -Identity $policyName -DistributionDetail
        $policyDetails += [PSCustomObject]@{
            Name               = $policy.Name
            DistributionStatus = $policy.DistributionStatus
            Mode               = $policy.Mode
            Enabled            = $policy.Enabled
            LastModified       = $policy.WhenChanged
        }
    } catch {
        Write-Host "Failed to retrieve distribution status for policy: $policyName" -ForegroundColor Red
    }
}

# Display all policies with their distribution statuses in the console
Write-Host "DLP Policies and their Distribution Statuses:" -ForegroundColor Green
$policyDetails | Format-Table Name, DistributionStatus, Mode, Enabled, LastModified -AutoSize

# Export the results to a CSV file
$csvFilePath = ".\DlpDistributionStatus.csv"
$policyDetails | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

Write-Host "Detailed policy distribution information exported to $csvFilePath" -ForegroundColor Green


