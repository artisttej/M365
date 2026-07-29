[CmdletBinding()]
param(
    [Parameter()]
    [ValidateRange(1, 30)]
    [int]$Days = 30,

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$PageSize = 5000,

    [Parameter()]
    [ValidateRange(1, 24)]
    [int]$SliceHours = 24,

    [Parameter()]
    [ValidateRange(1, 60)]
    [int]$SafetyMarginMinutes = 5,

    [Parameter()]
    [string[]]$PolicyNameContains = @('D-USB', 'D-PRT'),

    [Parameter()]
    [string[]]$Workloads = @('Endpoint'),

    [Parameter()]
    [string]$OutputPath = (
        Join-Path -Path (Get-Location) -ChildPath (
            'ActivityExplorer-DUSB-DPRT-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
        )
    ),

    [Parameter()]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-FirstPropertyValue {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$PropertyName
    )

    foreach ($name in $PropertyName) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Get-NestedPropertyValue {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$PropertyName,

        [Parameter()]
        [int]$Depth = 0
    )

    if ($null -eq $InputObject -or $Depth -gt 6) {
        return $null
    }

    if ($InputObject -is [string]) {
        $text = ([string]$InputObject).Trim()
        if (($text.StartsWith('{') -and $text.EndsWith('}')) -or
            ($text.StartsWith('[') -and $text.EndsWith(']'))) {
            try {
                $parsed = $text | ConvertFrom-Json -ErrorAction Stop
                return Get-NestedPropertyValue -InputObject $parsed `
                    -PropertyName $PropertyName -Depth ($Depth + 1)
            }
            catch {
                return $null
            }
        }

        return $null
    }

    foreach ($name in $PropertyName) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    # Some activity fields are represented as Name/Value or Key/Value pairs.
    foreach ($labelPropertyName in @('Name', 'Key')) {
        $labelProperty = $InputObject.PSObject.Properties[$labelPropertyName]
        $valueProperty = $InputObject.PSObject.Properties['Value']

        if ($null -ne $labelProperty -and $null -ne $valueProperty) {
            foreach ($name in $PropertyName) {
                if (([string]$labelProperty.Value) -ieq $name) {
                    return $valueProperty.Value
                }
            }
        }
    }

    if ($InputObject -is [System.Array]) {
        foreach ($item in $InputObject) {
            $nestedValue = Get-NestedPropertyValue -InputObject $item `
                -PropertyName $PropertyName -Depth ($Depth + 1)
            if ($null -ne $nestedValue) {
                return $nestedValue
            }
        }

        return $null
    }

    foreach ($property in $InputObject.PSObject.Properties) {
        if ($null -eq $property.Value -or
            $property.Value -is [ValueType]) {
            continue
        }

        $nestedValue = Get-NestedPropertyValue -InputObject $property.Value `
            -PropertyName $PropertyName -Depth ($Depth + 1)
        if ($null -ne $nestedValue) {
            return $nestedValue
        }
    }

    return $null
}

function Get-ActivityPropertyValue {
    param(
        [Parameter(Mandatory)]
        [object]$Record,

        [Parameter(Mandatory)]
        [string[]]$PropertyName
    )

    $value = Get-FirstPropertyValue -InputObject $Record `
        -PropertyName $PropertyName

    if ($null -ne $value) {
        return $value
    }

    return Get-NestedPropertyValue -InputObject $Record `
        -PropertyName $PropertyName
}

function ConvertTo-FlatString {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [System.Array]) {
        return (($Value | ForEach-Object { [string]$_ }) -join ';')
    }

    return [string]$Value
}

function ConvertTo-UtcIsoString {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal `
        -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal

    $parsed = [DateTimeOffset]::Parse(
        [string]$Value,
        [System.Globalization.CultureInfo]::InvariantCulture,
        $styles
    )

    return $parsed.ToUniversalTime().ToString(
        'yyyy-MM-ddTHH:mm:ss.fffZ',
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

function ConvertTo-MatchKey {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    return ([regex]::Replace(
        [string]$Value,
        '[^A-Za-z0-9]',
        ''
    )).ToUpperInvariant()
}

function Test-ContainsAny {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [string[]]$SearchText
    )

    $normalizedValue = ConvertTo-MatchKey -Value $Value

    foreach ($text in $SearchText) {
        $normalizedSearchText = ConvertTo-MatchKey -Value $text

        if (-not [string]::IsNullOrWhiteSpace($normalizedSearchText) -and
            $normalizedValue.IndexOf(
                $normalizedSearchText,
                [System.StringComparison]::OrdinalIgnoreCase
            ) -ge 0) {
            return $true
        }
    }

    return $false
}

if (-not (Get-Command -Name Export-ActivityExplorerData -ErrorAction SilentlyContinue)) {
    throw @'
Export-ActivityExplorerData is unavailable in the current PowerShell session.
Run this script from an existing Security & Compliance PowerShell session.
No sign-in or connection is performed by this script.
'@
}

foreach ($requiredCommand in @('Get-DlpCompliancePolicy', 'Get-DlpComplianceRule')) {
    if (-not (Get-Command -Name $requiredCommand -ErrorAction SilentlyContinue)) {
        throw @"
$requiredCommand is unavailable in the current PowerShell session.
Run this script from an existing Security & Compliance PowerShell session.
No sign-in or connection is performed by this script.
"@
    }
}

$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutputPath
)
$outputDirectory = Split-Path -Path $resolvedOutputPath -Parent

if (-not (Test-Path -LiteralPath $outputDirectory)) {
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force
}

if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
    throw "Output file already exists: $resolvedOutputPath. Use -Force to overwrite it."
}

$policyFragments = @(
    $PolicyNameContains |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
)

if ($policyFragments.Count -eq 0) {
    throw 'Specify at least one non-empty value in -PolicyNameContains.'
}

$workloadValues = @(
    $Workloads |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
)

Write-Host (
    'Policy name contains: {0}' -f ($policyFragments -join ' OR ')
) -ForegroundColor Cyan

if ($workloadValues.Count -gt 0) {
    Write-Host (
        'Server-side workload filter: {0}' -f ($workloadValues -join ', ')
    ) -ForegroundColor Cyan
}

Write-Host 'Server-side activity filter: DLPRuleMatch' -ForegroundColor Cyan

$candidatePolicies = @(
    foreach ($policy in @(Get-DlpCompliancePolicy -ErrorAction Stop)) {
        $candidatePolicyName = ConvertTo-FlatString (
            Get-FirstPropertyValue -InputObject $policy -PropertyName @(
                'Name',
                'DisplayName'
            )
        )

        if (-not [string]::IsNullOrWhiteSpace($candidatePolicyName) -and
            (Test-ContainsAny -Value $candidatePolicyName -SearchText $policyFragments)) {
            [pscustomobject]@{
                Name   = $candidatePolicyName
                Source = $policy
            }
        }
    }
)

if ($candidatePolicies.Count -eq 0) {
    throw (
        'No configured DLP policy name contains: {0}' -f
        ($policyFragments -join ' OR ')
    )
}

Write-Host 'Matched policy configurations:' -ForegroundColor Cyan
$candidatePolicies.Name | Sort-Object | ForEach-Object { Write-Host "  $_" }

$policyLookup = @{}

foreach ($policy in $candidatePolicies) {
    foreach ($propertyName in @(
        'Name',
        'DisplayName',
        'Guid',
        'ImmutableId',
        'Identity',
        'Id',
        'ObjectGuid'
    )) {
        $property = $policy.Source.PSObject.Properties[$propertyName]
        if ($null -eq $property -or $null -eq $property.Value) {
            continue
        }

        $identifier = ConvertTo-FlatString -Value $property.Value
        if ([string]::IsNullOrWhiteSpace($identifier)) {
            continue
        }

        $identifiers = @($identifier)
        $identifiers += @(
            foreach ($guidMatch in [regex]::Matches(
                $identifier,
                '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
            )) {
                $guidMatch.Value
            }
        )

        foreach ($policyIdentifier in $identifiers) {
            $lookupKey = ConvertTo-MatchKey -Value $policyIdentifier
            if (-not [string]::IsNullOrWhiteSpace($lookupKey)) {
                $policyLookup[$lookupKey] = $policy.Name
            }
        }
    }
}

# DLPRuleMatch records can omit PolicyName. Build a lookup from rule names and
# IDs to their parent policy so those records can still be filtered correctly.
$ruleLookup = @{}

foreach ($policy in $candidatePolicies) {
    $rules = @(Get-DlpComplianceRule -Policy $policy.Name -ErrorAction Stop)

    foreach ($rule in $rules) {
        $resolvedRuleName = ConvertTo-FlatString (
            Get-FirstPropertyValue -InputObject $rule -PropertyName @(
                'Name',
                'DisplayName'
            )
        )

        $ruleIdentifiers = @(
            foreach ($propertyName in @(
                'Name',
                'DisplayName',
                'Guid',
                'ImmutableId',
                'Identity',
                'Id',
                'ObjectGuid'
            )) {
                $property = $rule.PSObject.Properties[$propertyName]
                if ($null -ne $property -and $null -ne $property.Value) {
                    $identifier = ConvertTo-FlatString -Value $property.Value
                    if (-not [string]::IsNullOrWhiteSpace($identifier)) {
                        $identifier

                        foreach ($guidMatch in [regex]::Matches(
                            $identifier,
                            '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
                        )) {
                            $guidMatch.Value
                        }
                    }
                }
            }
        )

        foreach ($identifier in @($ruleIdentifiers | Select-Object -Unique)) {
            $lookupKey = ConvertTo-MatchKey -Value $identifier
            if ([string]::IsNullOrWhiteSpace($lookupKey)) {
                continue
            }

            $lookupValue = [pscustomobject]@{
                PolicyName = $policy.Name
                RuleName   = $resolvedRuleName
            }

            if (-not $ruleLookup.ContainsKey($lookupKey)) {
                $ruleLookup[$lookupKey] = $lookupValue
            }
            elseif ($null -ne $ruleLookup[$lookupKey] -and
                $ruleLookup[$lookupKey].PolicyName -ne $policy.Name) {
                # A duplicate rule name across policies is ambiguous. Unique
                # GUID-based keys remain usable.
                $ruleLookup[$lookupKey] = $null
            }
        }
    }
}

$nowUtc = [DateTime]::UtcNow

# Avoid the service rejecting an exact 30-day boundary because of processing
# time or minor clock differences. Also avoid an EndTime that the service
# could interpret as a few seconds in the future.
$startUtc = $nowUtc.AddDays(-$Days).AddMinutes($SafetyMarginMinutes)
$endUtc = $nowUtc.AddMinutes(-$SafetyMarginMinutes)

if ($startUtc -ge $endUtc) {
    throw 'The calculated date range is empty. Reduce -SafetyMarginMinutes.'
}

$sliceStartUtc = $startUtc
$totalExported = 0L
$seenRecordIds = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$matchedPolicyNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# Write the header once so every subsequent page can be appended immediately.
'"PolicyName","RuleName","User","DateUTC"' |
    Set-Content -LiteralPath $resolvedOutputPath -Encoding utf8 -Force

while ($sliceStartUtc -lt $endUtc) {
    $nextBoundaryUtc = $sliceStartUtc.AddHours($SliceHours)
    if ($nextBoundaryUtc -gt $endUtc) {
        $nextBoundaryUtc = $endUtc
    }

    # Keep adjacent slices non-overlapping while retaining the final endpoint.
    $sliceEndUtc = if ($nextBoundaryUtc -lt $endUtc) {
        $nextBoundaryUtc.AddTicks(-1)
    }
    else {
        $endUtc
    }

    Write-Host (
        'Querying {0:u} to {1:u}' -f $sliceStartUtc, $sliceEndUtc
    ) -ForegroundColor Yellow

    $baseParameters = @{
        StartTime    = $sliceStartUtc
        EndTime      = $sliceEndUtc
        PageSize     = $PageSize
        OutputFormat = 'Json'
        ErrorAction  = 'Stop'
    }

    $filterNumber = 1

    if ($workloadValues.Count -gt 0) {
        $baseParameters["Filter$filterNumber"] = @('Workload') + $workloadValues
        $filterNumber++
    }

    $baseParameters["Filter$filterNumber"] = @('Activity', 'DLPRuleMatch')

    $pageCookie = $null
    $pageNumber = 0

    do {
        $parameters = $baseParameters.Clone()
        if (-not [string]::IsNullOrWhiteSpace([string]$pageCookie)) {
            $parameters.PageCookie = $pageCookie
        }

        try {
            $response = Export-ActivityExplorerData @parameters
        }
        catch {
            throw (
                'Activity Explorer query failed for {0:u} to {1:u}: {2}' -f
                $sliceStartUtc,
                $sliceEndUtc,
                $_.Exception.Message
            )
        }

        $pageNumber++

        if ($null -eq $response) {
            throw (
                'Activity Explorer returned no response for {0:u} to {1:u}. Confirm the existing IPPS session is still active.' -f
                $sliceStartUtc,
                $sliceEndUtc
            )
        }

        $resultData = Get-FirstPropertyValue -InputObject $response -PropertyName @(
            'ResultData'
        )

        $records = @()

        if ($resultData -is [string]) {
            if ([string]::IsNullOrWhiteSpace($resultData)) {
                $records = @()
            }
            else {
                $records = @($resultData | ConvertFrom-Json)
            }
        }
        elseif ($null -ne $resultData) {
            $records = @($resultData)
        }

        $observedOnPage = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )

        $exportRows = @(
            foreach ($record in $records) {
                $policyName = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @(
                        'PolicyName',
                        'DLPPolicyName'
                    )
                )
                $policyId = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @(
                        'PolicyId',
                        'DLPPolicyId'
                    )
                )
                $ruleName = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @(
                        'RuleName',
                        'PolicyRuleName'
                    )
                )
                $ruleId = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @(
                        'RuleId',
                        'PolicyRuleId',
                        'DLPPolicyRuleId'
                    )
                )
                $activity = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @('Activity')
                )

                if ([string]::IsNullOrWhiteSpace($policyName)) {
                    $policyLookupKey = ConvertTo-MatchKey -Value $policyId
                    if (-not [string]::IsNullOrWhiteSpace($policyLookupKey) -and
                        $policyLookup.ContainsKey($policyLookupKey)) {
                        $policyName = $policyLookup[$policyLookupKey]
                    }
                }

                if ([string]::IsNullOrWhiteSpace($policyName)) {
                    foreach ($ruleIdentifier in @($ruleId, $ruleName)) {
                        $lookupKey = ConvertTo-MatchKey -Value $ruleIdentifier
                        if (-not [string]::IsNullOrWhiteSpace($lookupKey) -and
                            $ruleLookup.ContainsKey($lookupKey) -and
                            $null -ne $ruleLookup[$lookupKey]) {
                            $policyName = $ruleLookup[$lookupKey].PolicyName

                            if ([string]::IsNullOrWhiteSpace($ruleName)) {
                                $ruleName = $ruleLookup[$lookupKey].RuleName
                            }

                            break
                        }
                    }
                }

                $observedPolicyName = if ([string]::IsNullOrWhiteSpace($policyName)) {
                    '<missing>'
                }
                else {
                    $policyName
                }

                $observedPolicyId = if ([string]::IsNullOrWhiteSpace($policyId)) {
                    '<missing>'
                }
                else {
                    $policyId
                }

                $observedActivity = if ([string]::IsNullOrWhiteSpace($activity)) {
                    '<missing>'
                }
                else {
                    $activity
                }

                $observedRuleName = if ([string]::IsNullOrWhiteSpace($ruleName)) {
                    '<missing>'
                }
                else {
                    $ruleName
                }

                $observedRuleId = if ([string]::IsNullOrWhiteSpace($ruleId)) {
                    '<missing>'
                }
                else {
                    $ruleId
                }

                $null = $observedOnPage.Add(
                    "Activity=$observedActivity; PolicyName=$observedPolicyName; PolicyId=$observedPolicyId; RuleName=$observedRuleName; RuleId=$observedRuleId"
                )

                if ([string]::IsNullOrWhiteSpace($policyName) -or
                    -not (Test-ContainsAny -Value $policyName -SearchText $policyFragments)) {
                    continue
                }

                $null = $matchedPolicyNames.Add($policyName)

                $recordIdentity = ConvertTo-FlatString (
                    Get-FirstPropertyValue -InputObject $record -PropertyName @('RecordIdentity')
                )

                if (-not [string]::IsNullOrWhiteSpace($recordIdentity) -and
                    -not $seenRecordIds.Add($recordIdentity)) {
                    continue
                }

                $user = ConvertTo-FlatString (
                    Get-ActivityPropertyValue -Record $record -PropertyName @('User')
                )
                $eventTime = Get-ActivityPropertyValue -Record $record -PropertyName @(
                    'Happened',
                    'CreationTime'
                )

                [pscustomobject]@{
                    PolicyName = $policyName
                    RuleName   = $ruleName
                    User       = $user
                    DateUTC    = ConvertTo-UtcIsoString -Value $eventTime
                }
            }
        )

        if ($exportRows.Count -gt 0) {
            $exportRows |
                Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation `
                    -Encoding utf8 -Append
            $totalExported += $exportRows.Count
        }

        Write-Host (
            '  Page {0}: received {1}, exported {2}, running total {3}' -f
            $pageNumber,
            $records.Count,
            $exportRows.Count,
            $totalExported
        )

        if ($records.Count -gt 0 -and $exportRows.Count -eq 0) {
            Write-Warning (
                'Records were returned, but none matched {0}. Values observed:' -f
                ($policyFragments -join ' OR ')
            )
            $observedOnPage |
                Sort-Object |
                ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
        }

        $lastPageValue = Get-FirstPropertyValue -InputObject $response -PropertyName @(
            'LastPage'
        )
        $lastPage = [System.Convert]::ToBoolean($lastPageValue)

        if ($lastPage) {
            break
        }

        $pageCookie = ConvertTo-FlatString (
            Get-FirstPropertyValue -InputObject $response -PropertyName @(
                'Watermark',
                'WaterMark'
            )
        )

        if ([string]::IsNullOrWhiteSpace($pageCookie)) {
            throw "Page $pageNumber was not the last page, but no watermark was returned."
        }

        # The next request is intentionally immediate: the page cookie expires
        # 120 seconds after it is issued.
    } while ($true)

    $sliceStartUtc = $nextBoundaryUtc
}

Write-Host "Export complete: $totalExported rows" -ForegroundColor Green
Write-Host "CSV: $resolvedOutputPath" -ForegroundColor Green

[pscustomobject]@{
    OutputPath     = $resolvedOutputPath
    RowsExported   = $totalExported
    StartTimeUtc   = $startUtc
    EndTimeUtc     = $endUtc
    MatchedPolicies = @($matchedPolicyNames | Sort-Object)
}
