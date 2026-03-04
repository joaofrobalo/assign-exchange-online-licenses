# Assign-Licenses.ps1
# Reads email addresses from a CSV and assigns a Microsoft 365 license to each user.
#
# Prerequisites:
#   - Microsoft.Graph module installed: Install-Module Microsoft.Graph -Scope CurrentUser
#   - Connected to Microsoft Graph with license permissions:
#       Connect-MgGraph -Scopes "User.ReadWrite.All", "Organization.Read.All"
#
# CSV format: one column with header "EmailAddress"
# Example:
#   EmailAddress
#   user1@contoso.com
#   user2@contoso.com

param(
    [Parameter(Mandatory)]
    [string]$CsvPath
)

# --- Verify Microsoft.Graph is connected ---
try {
    $context = Get-MgContext
    if (-not $context) {
        throw "Not connected"
    }
    Write-Host "Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Cyan
} catch {
    Write-Error "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.ReadWrite.All','Organization.Read.All'"
    exit 1
}

# --- Load CSV ---
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

$csv = Import-Csv -Path $CsvPath

# Detect the email column
$emailColumn = $csv[0].PSObject.Properties.Name |
    Where-Object { $_ -in @("Email", "EmailAddress", "UserPrincipalName") } |
    Select-Object -First 1

if (-not $emailColumn) {
    $found = $csv[0].PSObject.Properties.Name -join ", "
    Write-Error "CSV must have a column named 'Email', 'EmailAddress', or 'UserPrincipalName'. Found: $found"
    exit 1
}

# Strip surrounding quotes in case values are stored as "user@domain.com"
$emails = $csv.$emailColumn |
    Where-Object { $_ -and $_.Trim() -ne "" } |
    ForEach-Object { $_.Trim().Trim('"') }

Write-Host "`nLoaded $($emails.Count) email(s) from CSV." -ForegroundColor Cyan

# --- List available SKUs in the tenant ---
Write-Host "`nFetching available licenses in your tenant..." -ForegroundColor Cyan
$skus = Get-MgSubscribedSku |
    Where-Object { $_.CapabilityStatus -eq "Enabled" } |
    Sort-Object SkuPartNumber

if (-not $skus) {
    Write-Error "No active license SKUs found in this tenant."
    exit 1
}

Write-Host "`nAvailable licenses:" -ForegroundColor Yellow
for ($i = 0; $i -lt $skus.Count; $i++) {
    $sku = $skus[$i]
    $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
    Write-Host "  [$($i + 1)] $($sku.SkuPartNumber)  |  Available: $available  |  SKU ID: $($sku.SkuId)"
}

# --- Let user pick a license ---
Write-Host ""
$selection = Read-Host "Enter the number of the license to assign"
$index = [int]$selection - 1

if ($index -lt 0 -or $index -ge $skus.Count) {
    Write-Error "Invalid selection."
    exit 1
}

$chosenSku = $skus[$index]
Write-Host "`nSelected: $($chosenSku.SkuPartNumber)" -ForegroundColor Green

$confirm = Read-Host "Assign '$($chosenSku.SkuPartNumber)' to $($emails.Count) user(s)? (y/n)"
if ($confirm -notmatch "^[Yy]") {
    Write-Host "Aborted." -ForegroundColor Yellow
    exit 0
}

# --- Assign licenses ---
$results = @()
$licenseBody = @{
    addLicenses    = @(@{ skuId = $chosenSku.SkuId })
    removeLicenses = @()
}

foreach ($email in $emails) {
    $email = $email.Trim()
    try {
        # Ensure UsageLocation is set (required before license assignment)
        $user = Get-MgUser -UserId $email -Property "UsageLocation" -ErrorAction Stop
        if (-not $user.UsageLocation) {
            Update-MgUser -UserId $email -UsageLocation "PT" -ErrorAction Stop
            Write-Host "  SET  $email - UsageLocation set to PT" -ForegroundColor DarkYellow
        }

        Set-MgUserLicense -UserId $email -BodyParameter $licenseBody -ErrorAction Stop | Out-Null
        Write-Host "  OK   $email" -ForegroundColor Green
        $results += [PSCustomObject]@{ Email = $email; Status = "Success"; Error = "" }
    } catch {
        $errMsg = $_.Exception.Message
        Write-Host "  FAIL $email - $errMsg" -ForegroundColor Red
        $results += [PSCustomObject]@{ Email = $email; Status = "Failed"; Error = $errMsg }
    }
}

# --- Summary ---
$ok   = ($results | Where-Object { $_.Status -eq "Success" }).Count
$fail = ($results | Where-Object { $_.Status -eq "Failed" }).Count
Write-Host "`nDone. Success: $ok  |  Failed: $fail" -ForegroundColor Cyan

if ($fail -gt 0) {
    Write-Host "`nFailed users:" -ForegroundColor Yellow
    $results | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
        Write-Host "  $($_.Email) - $($_.Error)" -ForegroundColor Red
    }
}

# Export results
$exportPath = [System.IO.Path]::ChangeExtension($CsvPath, "_results.csv")
$results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults saved to: $exportPath" -ForegroundColor Cyan
