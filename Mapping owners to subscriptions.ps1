# Install ImportExcel module if missing
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Paths
$ActiveUsersCsvPath = "C:\Users\vvacarescu\Downloads\tag_scripts\EnabledUserEmails.csv"
$SubscriptionsFolder = "C:\Users\vvacarescu\Downloads\tag_scripts"

# Find latest SubscriptionsTags CSV
$subsFile = Get-ChildItem -Path $SubscriptionsFolder -Filter "SubscriptionsTags_*.csv" |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

if (-not $subsFile) {
    throw "No SubscriptionsTags CSV found in $SubscriptionsFolder"
}

Write-Host "Using active users file: $ActiveUsersCsvPath"
Write-Host "Using subscriptions file: $($subsFile.FullName)"

# Load data
$activeUsers = Import-Csv $ActiveUsersCsvPath
$subsData    = Import-Csv $subsFile.FullName

# Build set of active user emails (lowercase for matching)
$activeEmails = $activeUsers |
    Where-Object { $_.Mail -and $_.Mail.Trim() -ne "" } |
    ForEach-Object { $_.Mail.Trim().ToLower() }

# Always initialize the hashset, even if $activeEmails is empty
$activeSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($email in $activeEmails) {
    $activeSet.Add($email) | Out-Null
}

# Compare and create mapping
$results = foreach ($row in $subsData) {
    $ownerRaw = $row.OwnerTag
    $ownerEmails = @()
    if ($ownerRaw) {
        $ownerEmails = ($ownerRaw -split '[,; ]+') | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ }
    }

$isValid = $false
if ($ownerEmails.Count -gt 0) {
    foreach ($email in $ownerEmails) {
        $emailUserPart = ($email -split '@')[0]  # part before @
        foreach ($activeEmail in $activeSet) {
            if (($activeEmail -split '@')[0] -eq $emailUserPart) {
                $isValid = $true
                break
            }
        }
        if ($isValid) { break }
    }
}


    [pscustomobject]@{
        SubscriptionId   = $row.SubscriptionId
        SubscriptionName = $row.SubscriptionName
        OwnerTag         = $ownerRaw
        MatchedEmails    = ($ownerEmails -join "; ")
        Status           = if ($isValid) { "Valid" } else { "Invalid" }
    }
}

# Output Excel path
$outputExcel = Join-Path $SubscriptionsFolder "OwnerTagMapping.xlsx"

# Export the data to Excel (no conditional formatting yet)
$results | Export-Excel -Path $outputExcel -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "Owner Tag Mapping"

# Re-open the Excel file and apply conditional formatting to column E (Status)
$excelPackage = Open-ExcelPackage -Path $outputExcel
$ws = $excelPackage.Workbook.Worksheets["Owner Tag Mapping"]

# Red for Invalid first
Set-ConditionalFormatting -Address "E2:E100" -RuleType ContainsText -Text "Invalid" -BackgroundColor Red -FontColor Black

# Green for Valid second
Set-ConditionalFormatting -Address "E2:E100" -RuleType ContainsText -Text "Valid" -BackgroundColor Green -FontColor Black


Close-ExcelPackage $excelPackage

Write-Host "Excel mapping file created: $outputExcel"
