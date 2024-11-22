# File path for output
$outputFile = "C:\Temp\DlpSITNames.csv"

$dlpPolicies = Get-DLPCompliancePolicy

$dlpRule = Get-DlpComplianceRule | Select-Object -ExpandProperty AdvancedRule

$allPolicyDetails = @()

# Iterate through each policy and its rules
foreach ($policy in $dlpPolicies) {
    $policyName = $policy.Name

    # Retrieve all rules for the policy
    $dlpRules = Get-DlpComplianceRule -Policy $policyName

    foreach ($rule in $dlpRules) {
        $ruleName = $rule.Name

        # Extract the AdvancedRule property
        if ($rule.AdvancedRule) {
            $advancedRule = $rule.AdvancedRule | ConvertFrom-Json

            # Extract SITs from the AdvancedRule
            $sitDetails = $advancedRule.Condition.SubConditions |
                Where-Object { $_.ConditionName -eq "ContentContainsSensitiveInformation" } |
                ForEach-Object {
                    $_.Value.Groups | ForEach-Object {
                        $_.Sensitivetypes | ForEach-Object {
                            [PSCustomObject]@{
                                PolicyName     = $policyName
                                RuleName       = $ruleName
                                SITName        = $_.Name
                                Confidence     = $_.Confidencelevel
                                MinConfidence  = $_.Minconfidence
                                MaxConfidence  = $_.Maxconfidence
                            }
                        }
                    }
                }

            # Add SIT details to the array
            if ($sitDetails) {
                $allPolicyDetails += $sitDetails
            }
        }
    }
}


# Export the details to a CSV file
$allPolicyDetails | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Output "All DLP Policies, Rules, and SITs exported to $outputFile"
