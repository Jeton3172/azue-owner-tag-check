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


# --- Auto-detect the "Status" column and highlight only Invalid ---
# Find the "Status" header in row 1 (case-insensitive)
$lastCol = $ws.Dimension.End.Column
$headerCells = $ws.Cells[1,1,1,$lastCol]
$statusCol = ($headerCells | Where-Object {
    ($_.Text   -as [string]).Trim().ToLower() -eq 'status' -or
    ($_.Value  -as [string]).Trim().ToLower() -eq 'status'
}).Start.Column

if (-not $statusCol) { throw "Couldn't find a 'Status' header in row 1." }

# Helper: convert column index -> Excel column letters (handles > Z)
function Get-ExcelColumnLetter([int]$col) {
    $s = ""
    while ($col -gt 0) {
        $col--; $s = [char](65 + ($col % 26)) + $s
        $col = [math]::Floor($col / 26)
    }
    return $s
}

$colLetter = Get-ExcelColumnLetter $statusCol
$lastRow   = $ws.Dimension.End.Row
$addr      = "{0}2:{0}{1}" -f $colLetter, $lastRow

# Only highlight "Invalid" in red
Add-ConditionalFormatting -Worksheet $ws -Address $addr -RuleType ContainsText -ConditionValue "Invalid" -BackgroundColor Red


Close-ExcelPackage $excelPackage

Write-Host "Excel mapping file created: $outputExcel"
