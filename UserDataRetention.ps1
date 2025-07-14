<#
.SYNOPSIS
Generates a comprehensive user data availability summary for retention justification.

.DESCRIPTION
This script investigates a user alias to determine data retention status across Azure AD, Exchange Online, and OneDrive.
It combines the strengths of previous versions to provide accurate, reliable results.

Key Features:
- Checks for deleted user accounts in Azure AD.
- Correctly identifies both active and soft-deleted mailboxes and their statistics.
- Performs a robust multi-tenant check for OneDrive storage across both Novartis tenants.
- Calculates the deletion timeline and days remaining for data recovery.
- Generates a clean, modern, and easy-to-read HTML report summarizing all findings.

.NOTES
This script requires the following PowerShell modules:
- Microsoft.Graph (Connect-MgGraph, Get-MgDirectoryDeletedItemAsUser)
- ExchangeOnlineManagement (Connect-ExchangeOnline, Get-Mailbox, Get-MailboxStatistics)
- Microsoft.Online.SharePoint.PowerShell (Connect-SPOService, Get-SPOSite)

Authentication requirements:
- Microsoft Graph (for deleted user checks)
- Exchange Online (for mailbox information)
- SharePoint Online Admin for Tenant 1: https://share-admin.novartis.net
- SharePoint Online Admin for Tenant 2: https://novartisnam-admin.sharepoint.com

The script will prompt for authentication as needed.
#>

Clear-Host
# Replaced box-drawing characters with standard ASCII for maximum compatibility.
Write-Host @"
+--------------------------------------------------------------+
|             USER DATA RETENTION INVESTIGATION TOOL           |
|       Property Of Novartis- Created by MOHD AZHAR UDDIN      |
+--------------------------------------------------------------+
"@ -ForegroundColor Cyan

#region Function Definitions

# Function to check OneDrive in a specific tenant
function Check-OneDriveInTenant {
    param(
        [string]$TenantAdminUrl,
        [string]$OneDriveUrl,
        [string]$TenantName
    )
    
    Write-Host "  Checking $TenantName..." -ForegroundColor Gray
    Write-Host "    Connecting to $TenantAdminUrl" -ForegroundColor Gray
    
    try {
        # Connect to the specific tenant's admin center. This will overwrite any previous connection.
        Connect-SPOService -Url $TenantAdminUrl -ErrorAction Stop
        
        # Check for the OneDrive site
        Write-Host "    Querying for site: $OneDriveUrl" -ForegroundColor Gray
        $odSite = Get-SPOSite -Identity $OneDriveUrl -ErrorAction Stop
        
        Write-Host "    [+] OneDrive found in $TenantName" -ForegroundColor Green
        
        # Return a result object
        return @{
            Found = $true
            TenantName = $TenantName
            SiteDetails = $odSite
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*cannot get site*" -or $errorMsg -like "*not found*") {
            Write-Host "    - OneDrive not found in $TenantName. This may be expected." -ForegroundColor Yellow
        } else {
            Write-Host "    [X] Error checking OneDrive in ${TenantName}: $errorMsg" -ForegroundColor Red
        }
        return @{ Found = $false }
    }
}

#endregion

# Get user alias
Write-Host "`nEnter User Alias: " -NoNewline -ForegroundColor Yellow
$UserAlias = Read-Host
$UserAlias = $UserAlias.ToUpper().Trim()

if ([string]::IsNullOrWhiteSpace($UserAlias)) {
    Write-Host "Error: User alias cannot be empty." -ForegroundColor Red
    exit
}

Write-Host "`nInvestigating: $UserAlias" -ForegroundColor Green
Write-Host "**************************************************" -ForegroundColor Gray

# Initialize all variables in a results hashtable
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$results = @{
    UserAlias = $UserAlias
    UserPrincipalName = "$UserAlias@novartis.net"
    Timestamp = $timestamp
    
    # Account
    AccountStatus = "Not Checked"
    AccountDeletedDate = "N/A"
    PrimarySmtp = ""
    
    # Mailbox
    MailboxFound = $false
    MailboxStatus = "Not Found"
    MailboxDisplayName = "N/A"
    MailboxItemCount = 0
    MailboxSize = "0 MB"
    MailboxDatabase = "N/A"
    MailboxLitigationHold = "N/A"
    MailboxInPlaceHolds = 0
    MailboxRetentionPolicy = "N/A"
    MailboxSoftDeletedDate = "N/A"
    
    # OneDrive
    OneDriveFound = $false
    OneDriveStatus = "Not Found"
    OneDriveUrl = "N/A"
    OneDriveTenant = "N/A"
    OneDriveStorage = "0 MB"
    
    # Timeline
    DeletionDate = "N/A"
    DaysSinceDelete = "N/A"
    DaysRemaining = "N/A"
    ExpirationDate = "N/A"
    
    # Recommendation
    HasData = $false
    Recommendation = "NO DATA"
    Action = "No action required"
}

# STEP 1: Check deleted users in Azure AD
Write-Host "`nSTEP 1: Checking Account Status..." -ForegroundColor Cyan
try {
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Gray
        Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome
    }
    
    Write-Host "  Searching deleted users for alias '$UserAlias'..." -ForegroundColor Gray
    $deletedUser = Get-MgDirectoryDeletedItemAsUser -All -Property UserPrincipalName,Mail,MailNickname,DeletedDateTime | Where-Object { $_.MailNickname -eq $UserAlias }
    
    if ($deletedUser) {
        $results.AccountStatus = "Deleted"
        $results.AccountDeletedDate = $deletedUser.DeletedDateTime
        $results.PrimarySmtp = if ($deletedUser.Mail) { $deletedUser.Mail } else { "N/A" }
        Write-Host "  [+] Account Status: DELETED" -ForegroundColor Red
        Write-Host "    Deleted Date: $($results.AccountDeletedDate)" -ForegroundColor Gray
        Write-Host "    Primary SMTP: $($results.PrimarySmtp)" -ForegroundColor Gray
    } else {
        $results.AccountStatus = "Active or Not Found"
        Write-Host "  [+] Account not found in deleted items directory (may be active)." -ForegroundColor Green
    }
} catch {
    Write-Host "  [X] Error checking account status: $_" -ForegroundColor Red
    $results.AccountStatus = "Error"
}

# STEP 2: Check Mailbox
Write-Host "`nSTEP 2: Checking Mailbox..." -ForegroundColor Cyan
try {
    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        Write-Host "  Connecting to Exchange Online..." -ForegroundColor Gray
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    $mailbox = $null
    $isActive = $false
    
    # First, try to find an active mailbox
    Write-Host "  Checking for active mailbox..." -ForegroundColor Gray
    try {
        $mailbox = Get-Mailbox -Identity $UserAlias -ErrorAction Stop
        if ($mailbox) {
            $isActive = $true
            $results.MailboxFound = $true
            $results.MailboxStatus = "Active"
            Write-Host "  [+] Found ACTIVE mailbox." -ForegroundColor Green
        }
    } catch {
        # This is expected if the mailbox is not active.
    }
    
    # If not active, check for a soft-deleted mailbox
    if (-not $isActive) {
        Write-Host "  Checking for soft-deleted mailbox..." -ForegroundColor Gray
        try {
            $mailbox = Get-Mailbox -Identity $UserAlias -SoftDeletedMailbox -ErrorAction Stop
            if ($mailbox) {
                $results.MailboxFound = $true
                $results.MailboxStatus = "Soft-Deleted"
                $results.MailboxSoftDeletedDate = if ($mailbox.WhenSoftDeleted) { ($mailbox.WhenSoftDeleted | ForEach-Object { $_.ToString() }) -join ", " } else { "Unknown" }
                $results.DeletionDate = $mailbox.WhenSoftDeleted # Set this as the primary deletion date
                Write-Host "  [+] Found SOFT-DELETED mailbox." -ForegroundColor Yellow
                Write-Host "    Soft-Deleted Date: $($results.MailboxSoftDeletedDate)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  - No mailbox found (neither active nor soft-deleted)." -ForegroundColor Yellow
        }
    }
    
    # If a mailbox was found, get its details
    if ($mailbox) {
        $results.MailboxDisplayName = $mailbox.DisplayName
        $results.MailboxDatabase = $mailbox.Database
        $results.MailboxLitigationHold = if ($mailbox.LitigationHoldEnabled) { "Enabled" } else { "Disabled" }
        $results.MailboxInPlaceHolds = $mailbox.InPlaceHolds.Count
        $results.MailboxRetentionPolicy = if ($mailbox.RetentionPolicy) { ($mailbox.RetentionPolicy | ForEach-Object { $_.ToString() }) -join ", " } else { "Default Policy" }
        
        Write-Host "  Getting mailbox statistics..." -ForegroundColor Gray
        try {
            $stats = $null
            if ($isActive) {
                $stats = Get-MailboxStatistics -Identity $UserAlias -ErrorAction Stop
            } else {
                $stats = Get-MailboxStatistics -Identity $mailbox.ExchangeGuid.ToString() -IncludeSoftDeletedRecipients -ErrorAction Stop
            }
            
            if ($stats) {
                $results.MailboxItemCount = $stats.ItemCount
                $results.MailboxSize = $stats.TotalItemSize.ToString()
                Write-Host "    Items: $($stats.ItemCount)" -ForegroundColor Gray
                Write-Host "    Size: $($stats.TotalItemSize)" -ForegroundColor Gray
                
                if ($stats.ItemCount -gt 0) {
                    $results.HasData = $true
                }
            }
        } catch {
            Write-Host "    - Could not retrieve mailbox statistics. $_" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "  [X] Error checking mailbox: $_" -ForegroundColor Red
}

# STEP 3: Check OneDrive
Write-Host "`nSTEP 3: Checking OneDrive (Multi-Tenant)..." -ForegroundColor Cyan
$tenants = @(
    @{Name="Tenant 1 (my.novartis.net)"; AdminUrl="https://share-admin.novartis.net"; SiteUrl="https://my.novartis.net/personal/${UserAlias}_novartis_net"},
    @{Name="Tenant 2 (novartisnam-my.sharepoint.com)"; AdminUrl="https://novartisnam-admin.sharepoint.com"; SiteUrl="https://novartisnam-my.sharepoint.com/personal/${UserAlias}_novartis_net"}
)

foreach ($tenant in $tenants) {
    $oneDriveResult = Check-OneDriveInTenant -TenantAdminUrl $tenant.AdminUrl -OneDriveUrl $tenant.SiteUrl -TenantName $tenant.Name
    
    if ($oneDriveResult.Found) {
        $site = $oneDriveResult.SiteDetails
        $results.OneDriveFound = $true
        $results.OneDriveStatus = "Found"
        $results.OneDriveUrl = $site.Url
        $results.OneDriveTenant = $oneDriveResult.TenantName
        $results.OneDriveStorage = "$($site.StorageUsageCurrent) MB"
        
        if ($site.StorageUsageCurrent -gt 0) {
            $results.HasData = $true
        }
        break
    }
}

if (-not $results.OneDriveFound) {
    Write-Host "  - OneDrive not found in any configured tenant." -ForegroundColor Yellow
}

# Disconnect from the last SPO session
Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null

# STEP 4: Calculate Deletion Timeline
Write-Host "`nSTEP 4: Calculating Deletion Timeline..." -ForegroundColor Cyan
if ($results.DeletionDate -eq "N/A" -and $results.AccountDeletedDate -ne "N/A") {
    $results.DeletionDate = $results.AccountDeletedDate
}

if ($results.DeletionDate -ne "N/A") {
    try {
        $delDate = [DateTime]$results.DeletionDate
        $daysSince = (Get-Date) - $delDate
        $results.DaysSinceDelete = [Math]::Floor($daysSince.TotalDays)
        $results.DaysRemaining = 30 - $results.DaysSinceDelete
        $results.ExpirationDate = $delDate.AddDays(30).ToString('yyyy-MM-dd')
        
        Write-Host "  Deletion Date: $($delDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
        Write-Host "  Days Since Deletion: $($results.DaysSinceDelete)" -ForegroundColor Gray
        Write-Host "  Days Remaining for Recovery: $($results.DaysRemaining)" -ForegroundColor $(if($results.DaysRemaining -lt 7){"Red"}else{"Gray"})
    } catch {
        Write-Host "  - Could not parse deletion date to calculate timeline." -ForegroundColor Yellow
    }
} else {
    Write-Host "  - No deletion date found to calculate timeline." -ForegroundColor Gray
}

# STEP 5: Generate Final Recommendation
Write-Host "`nSTEP 5: Generating Recommendation..." -ForegroundColor Cyan
if ($results.HasData) {
    $results.Recommendation = "RETAIN - User has recoverable data"
    $results.Action = "Maintain holds and restore if needed."
    Write-Host "  [+] RETAIN - User has data in Mailbox or OneDrive." -ForegroundColor Green
} else {
    $results.Recommendation = "NO DATA - Safe to proceed"
    $results.Action = "No significant data found. Standard cleanup can proceed."
    Write-Host "  [+] NO DATA - No data found in checked locations." -ForegroundColor Yellow
}

# Generate HTML Report
Write-Host "`nGenerating HTML Report..." -ForegroundColor Cyan
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Retention Report - $($results.UserAlias)</title>
    <style>
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; 
            background-color: #f4f7f9; 
            margin: 0; 
            padding: 20px; 
            color: #333;
        }
        .container { 
            max-width: 1200px; 
            margin: 20px auto; 
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            overflow: hidden;
        }
        .header { 
            background-color: #005A9E; /* A professional blue */
            color: white;
            padding: 30px; 
            text-align: center; 
        }
        h1 { 
            margin: 0 0 8px 0; 
            font-size: 28px; 
            font-weight: 600; 
        }
        .subtitle { 
            color: #e0e0e0; 
            font-size: 16px; 
        }
        .grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); 
            gap: 24px;
            padding: 24px;
        }
        .card { 
            background: #fdfdfd; 
            padding: 24px; 
            border-radius: 10px; 
            border: 1px solid #e9ecef;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        }
        .card-header { 
            display: flex; 
            align-items: center; 
            margin-bottom: 20px; 
            padding-bottom: 16px; 
            border-bottom: 1px solid #e9ecef; 
        }
        .icon { 
            margin-right: 12px;
            width: 24px;
            height: 24px;
            stroke-width: 2;
            color: #005A9E;
        }
        .card-title { 
            font-size: 18px; 
            font-weight: 600;
            color: #212529;
        }
        .row { 
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            padding: 12px 0; 
            border-bottom: 1px solid #f8f9fa; 
            font-size: 15px; 
        }
        .row:last-child { border: none; }
        .label { color: #6c757d; }
        .value { 
            font-weight: 600; 
            text-align: right; 
            word-break: break-all; 
            max-width: 65%;
        }
        .status-badge {
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 700;
            color: white;
            text-transform: uppercase;
        }
        .badge-green { background-color: #28a745; }
        .badge-red { background-color: #dc3545; }
        .badge-yellow { background-color: #ffc107; color: #333; }
        .badge-blue { background-color: #007bff; }
        .badge-gray { background-color: #6c757d; }

        .footer { 
            text-align: center; 
            color: #999; 
            padding: 20px;
            font-size: 14px;
            background-color: #f8f9fa;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>User Data Retention Summary</h1>
            <div class="subtitle">Investigation for <strong>$($results.UserAlias)</strong> | Generated: $($results.Timestamp)</div>
        </div>
        
        <div class="grid">
            <!-- Account Status Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                    <span class="card-title">Account Status</span>
                </div>
                <div class="row"><span class="label">Status</span><span class="value"><span class="status-badge $(if($results.AccountStatus -eq 'Deleted'){'badge-red'}else{'badge-green'})">$($results.AccountStatus)</span></span></div>
                <div class="row"><span class="label">Deleted Date</span><span class="value">$($results.AccountDeletedDate)</span></div>
                <div class="row"><span class="label">Original UPN</span><span class="value">$($results.UserPrincipalName)</span></div>
                <div class="row"><span class="label">Primary SMTP</span><span class="value">$($results.PrimarySmtp)</span></div>
            </div>
            
            <!-- Retention Status Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path></svg>
                    <span class="card-title">Retention Status</span>
                </div>
                <div class="row"><span class="label">Litigation Hold</span><span class="value"><span class="status-badge $(if($results.MailboxLitigationHold -eq 'Enabled'){'badge-green'}else{'badge-gray'})">$($results.MailboxLitigationHold)</span></span></div>
                <div class="row"><span class="label">In-Place Holds</span><span class="value">$($results.MailboxInPlaceHolds)</span></div>
                <div class="row"><span class="label">Retention Policy</span><span class="value">$($results.MailboxRetentionPolicy)</span></div>
            </div>
            
            <!-- Mailbox Information Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>
                    <span class="card-title">Mailbox Information</span>
                </div>
                <div class="row"><span class="label">Status</span><span class="value"><span class="status-badge $(if($results.MailboxStatus -eq 'Soft-Deleted'){'badge-red'}elseif($results.MailboxStatus -eq 'Active'){'badge-green'}else{'badge-yellow'})">$($results.MailboxStatus)</span></span></div>
                <div class="row"><span class="label">Display Name</span><span class="value">$($results.MailboxDisplayName)</span></div>
                <div class="row"><span class="label">Item Count</span><span class="value">$($results.MailboxItemCount)</span></div>
                <div class="row"><span class="label">Total Size</span><span class="value">$($results.MailboxSize)</span></div>
                $(if($results.MailboxSoftDeletedDate -ne 'N/A'){'<div class="row"><span class="label">When Soft Deleted</span><span class="value">' + $results.MailboxSoftDeletedDate + '</span></div>'})
            </div>
            
            <!-- Deletion Timeline Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                    <span class="card-title">Deletion Timeline</span>
                </div>
                <div class="row"><span class="label">Deletion Date</span><span class="value">$($results.DeletionDate)</span></div>
                <div class="row"><span class="label">Days Since Delete</span><span class="value">$($results.DaysSinceDelete)</span></div>
                <div class="row"><span class="label">Days Remaining</span><span class="value"><span class="status-badge $(if($results.DaysRemaining -ne 'N/A' -and $results.DaysRemaining -lt 7){'badge-red'}elseif($results.DaysRemaining -ne 'N/A' -and $results.DaysRemaining -lt 14){'badge-yellow'}else{'badge-green'})">$($results.DaysRemaining)</span></span></div>
                <div class="row"><span class="label">Expiration Date</span><span class="value">$($results.ExpirationDate)</span></div>
            </div>
            
            <!-- OneDrive Information Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M18 10a4 4 0 0 0-3.46-3.97A5 5 0 0 0 8 9a4 4 0 0 0-3.97 3.46A4 4 0 0 0 8 18h8a4 4 0 0 0 3.97-4.54A4 4 0 0 0 18 10z"></path></svg>
                    <span class="card-title">OneDrive Information</span>
                </div>
                <div class="row"><span class="label">Status</span><span class="value"><span class="status-badge $(if($results.OneDriveStatus -eq 'Found'){'badge-blue'}else{'badge-gray'})">$($results.OneDriveStatus)</span></span></div>
                <div class="row"><span class="label">Found In Tenant</span><span class="value">$($results.OneDriveTenant)</span></div>
                <div class="row"><span class="label">Used Space</span><span class="value">$($results.OneDriveStorage)</span></div>
                <div class="row"><span class="label">OneDrive URL</span><span class="value">$($results.OneDriveUrl)</span></div>
            </div>
            
            <!-- Recommendation Card -->
            <div class="card">
                <div class="card-header">
                    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                    <span class="card-title">Recommendation</span>
                </div>
                <div class="row"><span class="label">Data Assessment</span><span class="value">$(if($results.HasData){'Data Present'}else{'No Data Found'})</span></div>
                <div class="row"><span class="label">Recommendation</span><span class="value"><span class="status-badge $(if($results.HasData){'badge-green'}else{'badge-yellow'})">$($results.Recommendation)</span></span></div>
                <div class="row"><span class="label">Action</span><span class="value">$($results.Action)</span></div>
            </div>
        </div>
        
        <div class="footer">Report generated by MOHD AZHAR UDDIN</div>
    </div>
</body>
</html>
"@

# Save and open the report
$reportFile = "RetentionReport_$($results.UserAlias)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath = Join-Path (Get-Location) $reportFile

try {
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "[+] Report saved to: $reportPath" -ForegroundColor Green
    Start-Process $reportPath
} catch {
    Write-Host "[X] Error saving report: $_" -ForegroundColor Red
}

# Final Console Summary
Write-Host "`n**************************************************" -ForegroundColor Cyan
Write-Host "INVESTIGATION COMPLETE" -ForegroundColor Cyan
Write-Host "**************************************************" -ForegroundColor Cyan
Write-Host "User:           $($results.UserAlias)"
Write-Host "Account Status: $($results.AccountStatus)"
Write-Host "Mailbox Status: $($results.MailboxStatus)"
Write-Host "OneDrive Found: $($results.OneDriveFound)"
Write-Host "Has Data:       $(if($results.HasData){'YES'}else{'NO'})"
Write-Host "Recommendation: $($results.Recommendation)"
Write-Host "**************************************************" -ForegroundColor Cyan
