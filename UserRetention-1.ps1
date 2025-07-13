<#
.SYNOPSIS
User Data Retention Investigation - Final Merged Version

.DESCRIPTION
Combines working components from both scripts:
- Account status check that properly captures deletion dates
- Mailbox check that properly handles soft-deleted mailboxes  
- OneDrive multi-tenant check that properly handles authentication
#>

Clear-Host
Write-Host @"
+---------------------------------------------------------------+
¦             USER DATA RETENTION INVESTIGATION TOOL             ¦
¦                       Final Version 5.0                        ¦
+---------------------------------------------------------------+
"@ -ForegroundColor Cyan

# Get user alias
Write-Host "`nEnter User Alias: " -NoNewline -ForegroundColor Yellow
$UserAlias = Read-Host
$UserAlias = $UserAlias.ToUpper().Trim()

if ([string]::IsNullOrWhiteSpace($UserAlias)) {
    Write-Host "Error: User alias required" -ForegroundColor Red
    exit
}

$UserPrincipalName = "$UserAlias@novartis.net"
Write-Host "`nInvestigating: $UserAlias" -ForegroundColor Green
Write-Host "????????????????????????????????????????????????????????????" -ForegroundColor Gray

# Initialize all variables
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$results = @{
    UserAlias = $UserAlias
    UserPrincipalName = $UserPrincipalName
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
    OneDriveTenant = "N/A"
    OneDriveStorage = 0
    OneDriveURL = ""
    
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

# Initialize deletion date variable to track across checks
$deletedDate = $null

# STEP 1: Check Account Status (FROM OLD SCRIPT - WORKING)
Write-Host "`nSTEP 1: Checking Account Status..." -ForegroundColor Cyan
Write-Host "[Account Status Check - Deleted Users via MgGraph]" -ForegroundColor Yellow
Write-Host "--------------------------------------------------"

try {
    # Connect to Microsoft Graph if not already connected
    if (-not (Get-MgContext)) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Gray
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
    }
    
    # Fetch all deleted users
    Write-Host "Searching for deleted users..." -ForegroundColor Gray
    $deletedUsers = Get-MgDirectoryDeletedItemAsUser -All -Property Id,UserPrincipalName,MailNickName,DeletedDateTime,Mail

    # Filter for match on alias
    $matchedUsers = $deletedUsers | Where-Object {
        $_.MailNickName -ieq $UserAlias
    }

    if ($matchedUsers) {
        Write-Host "`n--- Deleted user found ---`n" -ForegroundColor Yellow
        $matchedUsers | Select-Object UserPrincipalName, Mail, MailNickName, DeletedDateTime | Format-Table -AutoSize
        
        $results.AccountStatus = "Deleted"
        $deletedDate = $matchedUsers[0].DeletedDateTime
        $results.AccountDeletedDate = $deletedDate
        $results.PrimarySmtp = if ($matchedUsers[0].Mail) { $matchedUsers[0].Mail } else { "" }
        
        Write-Host "Account Status:     DELETED" -ForegroundColor Red
        Write-Host "Original UPN:       $UserPrincipalName"
        Write-Host "Primary SMTP:       $($results.PrimarySmtp)"
        Write-Host "Deleted Date:       $($deletedDate.ToString('dd-MMM-yyyy HH:mm'))"
    } else {
        Write-Host "`nNo deleted users found matching alias [$UserAlias]"
        $results.AccountStatus = "Active"
    }
}
catch {
    Write-Host "Error occurred during deleted user check: $_" -ForegroundColor Red
    $results.AccountStatus = "Error"
}

# STEP 2: Check Mailbox (FROM NEW SCRIPT - WORKING)
Write-Host "`nSTEP 2: Checking Mailbox..." -ForegroundColor Cyan
try {
    # Ensure connected to Exchange Online
    $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoSession) {
        Write-Host "  Connecting to Exchange Online..." -ForegroundColor Gray
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Try active mailbox first
    Write-Host "  Checking for active mailbox..." -ForegroundColor Gray
    $mailbox = $null
    $isActive = $false
    
    try {
        $mailbox = Get-Mailbox -Identity $UserAlias -ErrorAction Stop
        if ($mailbox) {
            $isActive = $true
            $results.MailboxFound = $true
            $results.MailboxStatus = "Active"
            Write-Host "  ? Found ACTIVE mailbox" -ForegroundColor Green
        }
    } catch {
        # Not active, this is expected
    }
    
    # If not active, check soft-deleted
    if (-not $isActive) {
        Write-Host "  Checking for soft-deleted mailbox..." -ForegroundColor Gray
        try {
            $mailbox = Get-Mailbox -Identity $UserAlias -SoftDeletedMailbox -ErrorAction Stop
            if ($mailbox) {
                $results.MailboxFound = $true
                $results.MailboxStatus = "Soft-Deleted"
                $results.MailboxSoftDeletedDate = if ($mailbox.WhenSoftDeleted) { $mailbox.WhenSoftDeleted.ToString() } else { "Unknown" }
                Write-Host "  ? Found SOFT-DELETED mailbox" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "  ? No mailbox found (active or soft-deleted)" -ForegroundColor Red
        }
    }
    
    # If we found a mailbox, get details
    if ($mailbox) {
        $results.MailboxDisplayName = $mailbox.DisplayName
        $results.MailboxDatabase = $mailbox.Database
        $results.MailboxLitigationHold = if ($mailbox.LitigationHoldEnabled) { "Enabled" } else { "Disabled" }
        $results.MailboxInPlaceHolds = $mailbox.InPlaceHolds.Count
        $results.MailboxRetentionPolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy.ToString() } else { "Default" }
        
        Write-Host "    Display Name: $($results.MailboxDisplayName)" -ForegroundColor Gray
        Write-Host "    Database: $($results.MailboxDatabase)" -ForegroundColor Gray
        Write-Host "    Litigation Hold: $($results.MailboxLitigationHold)" -ForegroundColor Gray
        
        # Try to get statistics
        Write-Host "  Getting mailbox statistics..." -ForegroundColor Gray
        try {
            $stats = $null
            if ($isActive) {
                $stats = Get-MailboxStatistics -Identity $UserAlias -ErrorAction Stop
            } else {
                # For soft-deleted, try different methods
                try {
                    $stats = Get-MailboxStatistics -Identity $mailbox.ExchangeGuid.ToString() -IncludeSoftDeletedRecipients -ErrorAction Stop
                } catch {
                    try {
                        $stats = Get-MailboxStatistics -Identity $UserAlias -Database $mailbox.Database -IncludeSoftDeletedRecipients -ErrorAction Stop
                    } catch {
                        Write-Host "    Could not retrieve statistics" -ForegroundColor Gray
                    }
                }
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
            Write-Host "    Statistics unavailable" -ForegroundColor Gray
        }
        
        # Update deletion date if soft-deleted and we don't have one yet
        if ($mailbox.WhenSoftDeleted -and -not $deletedDate) {
            $deletedDate = $mailbox.WhenSoftDeleted
        }
    }
} catch {
    Write-Host "  ? Error checking mailbox: $_" -ForegroundColor Red
}

# STEP 3: Check OneDrive (FROM OLD SCRIPT - WORKING)
Write-Host "`nSTEP 3: Checking OneDrive (Multi-Tenant)..." -ForegroundColor Cyan
Write-Host "[OneDrive Information - Multi-Tenant Check]" -ForegroundColor Yellow

# OneDrive URLs for both tenants
$OneDriveURL_Tenant1 = "https://my.novartis.net/personal/${UserAlias}_novartis_net"
$OneDriveURL_Tenant2 = "https://novartisnam-my.sharepoint.com/personal/${UserAlias}_novartis_net"

# Function to check OneDrive in a specific tenant
function Check-OneDriveInTenant {
    param(
        [string]$TenantUrl,
        [string]$OneDriveUrl,
        [string]$TenantName
    )
    
    Write-Host "`nChecking $TenantName..." -ForegroundColor Cyan
    
    try {
        # Disconnect any existing connection
        try { 
            Disconnect-SPOService -ErrorAction SilentlyContinue 
            Write-Host "Disconnected from previous SPO session" -ForegroundColor Gray
        } catch {}
        
        # Connect to the specific tenant
        Write-Host "Connecting to $TenantUrl..."
        Write-Host "Please authenticate when prompted (you may use a different account for SharePoint)" -ForegroundColor Yellow
        Connect-SPOService -Url $TenantUrl -ErrorAction Stop
        
        # Check OneDrive
        Write-Host "Checking for OneDrive at: $OneDriveUrl" -ForegroundColor Gray
        $od = Get-SPOSite -Identity $OneDriveUrl -ErrorAction Stop
        
        Write-Host "OneDrive Status:    $($od.Status)" -ForegroundColor Green
        Write-Host "OneDrive URL:       $($od.Url)"
        Write-Host "Owner:              $($od.Owner)"
        Write-Host "Used Space:         $($od.StorageUsageCurrent) MB"
        Write-Host "Storage Quota:      $($od.StorageQuota) MB"
        Write-Host "Last Modified:      $($od.LastContentModifiedDate)"
        
        # Calculate storage percentage
        $storagePercent = 0
        if ($od.StorageQuota -gt 0) {
            $storagePercent = [math]::Round(($od.StorageUsageCurrent / $od.StorageQuota) * 100, 2)
            Write-Host "Storage Used:       $storagePercent%"
        }
        
        Write-Host "Successfully found OneDrive in $TenantName" -ForegroundColor Green
        
        return @{
            Found = $true
            Tenant = $TenantName
            Details = $od
            StoragePercent = $storagePercent
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*401*" -or $errorMsg -like "*unauthorized*") {
            Write-Host "Authentication failed for $TenantName" -ForegroundColor Red
            Write-Host "Please check your credentials and permissions" -ForegroundColor Yellow
        } elseif ($errorMsg -like "*cannot get site*" -or $errorMsg -like "*not found*") {
            Write-Host "OneDrive not found in $TenantName" -ForegroundColor Yellow
            Write-Host "This is normal if the user's OneDrive is in a different tenant" -ForegroundColor Gray
        } else {
            Write-Host "Error checking OneDrive in $TenantName" -ForegroundColor Red
            Write-Host "Error details: $errorMsg" -ForegroundColor Gray
        }
        return @{ Found = $false; Tenant = $TenantName; Error = $errorMsg }
    }
}

# Show multi-tenant notice
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Multi-Tenant OneDrive Check" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Note: You may be prompted to authenticate for each tenant" -ForegroundColor Yellow
Write-Host "This is normal if you use different accounts for SharePoint admin" -ForegroundColor Yellow

# Check Tenant 1
$tenant1Result = Check-OneDriveInTenant -TenantUrl "https://share-admin.novartis.net" -OneDriveUrl $OneDriveURL_Tenant1 -TenantName "Tenant 1 (my.novartis.net)"

# Check Tenant 2
$tenant2Result = Check-OneDriveInTenant -TenantUrl "https://novartisnam-admin.sharepoint.com" -OneDriveUrl $OneDriveURL_Tenant2 -TenantName "Tenant 2 (novartisnam-my.sharepoint.com)"

# Determine which OneDrive to use for reporting
if ($tenant1Result.Found) {
    $results.OneDriveFound = $true
    $results.OneDriveStatus = "Found in Tenant 1"
    $results.OneDriveTenant = "Tenant 1 (my.novartis.net)"
    $results.OneDriveURL = $OneDriveURL_Tenant1
    $results.OneDriveStorage = $tenant1Result.Details.StorageUsageCurrent
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "OneDrive Summary: Found in Tenant 1 (my.novartis.net)" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    if ($tenant1Result.Details.StorageUsageCurrent -gt 0) {
        $results.HasData = $true
    }
} elseif ($tenant2Result.Found) {
    $results.OneDriveFound = $true
    $results.OneDriveStatus = "Found in Tenant 2"
    $results.OneDriveTenant = "Tenant 2 (novartisnam-my.sharepoint.com)"
    $results.OneDriveURL = $OneDriveURL_Tenant2
    $results.OneDriveStorage = $tenant2Result.Details.StorageUsageCurrent
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "OneDrive Summary: Found in Tenant 2 (novartisnam-my.sharepoint.com)" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    if ($tenant2Result.Details.StorageUsageCurrent -gt 0) {
        $results.HasData = $true
    }
} else {
    $results.OneDriveStatus = "Not Found"
    $results.OneDriveURL = "Checked both tenants"
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "OneDrive Summary: Not found in either tenant" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
}

# Disconnect from SharePoint after checks are complete
try { 
    Disconnect-SPOService -ErrorAction SilentlyContinue 
    Write-Host "`nDisconnected from SharePoint Online" -ForegroundColor Gray
} catch {}

# STEP 4: Calculate timeline
Write-Host "`nSTEP 4: Calculating Timeline..." -ForegroundColor Cyan

# Use deletion date from account or mailbox
if ($deletedDate) {
    $results.DeletionDate = $deletedDate
    
    try {
        $daysSince = (Get-Date) - $deletedDate
        $results.DaysSinceDelete = [Math]::Floor($daysSince.TotalDays)
        $results.DaysRemaining = 30 - $results.DaysSinceDelete
        $results.ExpirationDate = $deletedDate.AddDays(30)
        
        Write-Host "  Deletion Date: $($deletedDate.ToString('yyyy-MM-dd HH:mm'))"
        Write-Host "  Days since deletion: $($results.DaysSinceDelete)" -ForegroundColor Gray
        Write-Host "  Days remaining: $($results.DaysRemaining)" -ForegroundColor $(if($results.DaysRemaining -lt 7){"Red"}elseif($results.DaysRemaining -lt 14){"Yellow"}else{"Gray"})
        Write-Host "  Expiration: $($results.ExpirationDate.ToString('yyyy-MM-dd'))"
    } catch {
        Write-Host "  Error calculating timeline" -ForegroundColor Red
    }
} else {
    Write-Host "  No deletion date found" -ForegroundColor Gray
}

# STEP 5: Generate recommendation
Write-Host "`nSTEP 5: Recommendation..." -ForegroundColor Cyan
if ($results.HasData) {
    $results.Recommendation = "RETAIN - User has recoverable data"
    $results.Action = "Maintain holds and restore if needed"
    Write-Host "  ? RETAIN - User has data" -ForegroundColor Green
} else {
    $results.Recommendation = "NO DATA - Safe to proceed"
    $results.Action = "No significant data found"
    Write-Host "  ? NO DATA - Safe to proceed" -ForegroundColor Yellow
}

# Generate HTML
Write-Host "`nGenerating Report..." -ForegroundColor Cyan
$html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Retention Report - $($results.UserAlias)</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; background: #f0f2f5; margin: 0; padding: 20px; color: #1a1a1a; }
        .container { max-width: 1200px; margin: 0 auto; }
        .header { background: white; padding: 30px; text-align: center; border-radius: 12px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        h1 { margin: 0 0 8px 0; font-size: 28px; font-weight: 600; }
        .subtitle { color: #666; font-size: 16px; }
        .search-box { background: white; padding: 24px; text-align: center; border-radius: 12px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        input { padding: 12px 24px; font-size: 18px; border: 2px solid #e1e4e8; border-radius: 8px; background: #f6f8fa; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 20px; }
        .card { background: white; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .card-header { display: flex; align-items: center; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 2px solid #f0f2f5; }
        .icon { font-size: 24px; margin-right: 12px; }
        .card-title { font-size: 18px; font-weight: 600; }
        .row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #f6f8fa; }
        .row:last-child { border: none; }
        .label { color: #586069; font-weight: 500; }
        .value { font-weight: 600; text-align: right; max-width: 60%; word-break: break-word; }
        .active { color: #28a745; }
        .deleted { color: #dc3545; }
        .warning { color: #ffc107; }
        .footer { text-align: center; color: #999; margin-top: 40px; font-size: 14px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>User Data Retention Summary</h1>
            <div class="subtitle">Investigation for user: <strong>$($results.UserAlias)</strong> | Generated: $($results.Timestamp)</div>
        </div>
        
        <div class="search-box">
            <input type="text" value="$($results.UserAlias)" readonly>
        </div>
        
        <div class="grid">
            <div class="card">
                <div class="card-header">
                    <span class="icon">??</span>
                    <span class="card-title">Account Status</span>
                </div>
                <div class="row">
                    <span class="label">Account Status</span>
                    <span class="value $(if($results.AccountStatus -eq 'Deleted'){'deleted'}else{'active'})">$($results.AccountStatus)</span>
                </div>
                <div class="row">
                    <span class="label">Deleted Date</span>
                    <span class="value">$(if($results.AccountDeletedDate -ne 'N/A'){$results.AccountDeletedDate.ToString('yyyy-MM-dd HH:mm:ss')}else{'N/A'})</span>
                </div>
                <div class="row">
                    <span class="label">Original UPN</span>
                    <span class="value">$($results.UserPrincipalName)</span>
                </div>
                $(if($results.PrimarySmtp){
                '<div class="row">
                    <span class="label">Primary SMTP</span>
                    <span class="value">' + $results.PrimarySmtp + '</span>
                </div>'})
            </div>
            
            <div class="card">
                <div class="card-header">
                    <span class="icon">??</span>
                    <span class="card-title">Retention Status</span>
                </div>
                <div class="row">
                    <span class="label">Litigation Hold</span>
                    <span class="value">$($results.MailboxLitigationHold)</span>
                </div>
                <div class="row">
                    <span class="label">In-Place Holds</span>
                    <span class="value">$($results.MailboxInPlaceHolds)</span>
                </div>
                <div class="row">
                    <span class="label">Retention Policy</span>
                    <span class="value">$($results.MailboxRetentionPolicy)</span>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <span class="icon">??</span>
                    <span class="card-title">Mailbox Information</span>
                </div>
                <div class="row">
                    <span class="label">Mailbox Status</span>
                    <span class="value $(if($results.MailboxStatus -eq 'Soft-Deleted'){'deleted'}elseif($results.MailboxStatus -eq 'Active'){'active'}else{'warning'})">$($results.MailboxStatus)</span>
                </div>
                <div class="row">
                    <span class="label">Display Name</span>
                    <span class="value">$($results.MailboxDisplayName)</span>
                </div>
                <div class="row">
                    <span class="label">Item Count</span>
                    <span class="value">$($results.MailboxItemCount)</span>
                </div>
                <div class="row">
                    <span class="label">Total Size</span>
                    <span class="value">$($results.MailboxSize)</span>
                </div>
                $(if($results.MailboxSoftDeletedDate -ne 'N/A'){
                '<div class="row">
                    <span class="label">When Soft Deleted</span>
                    <span class="value">' + $results.MailboxSoftDeletedDate + '</span>
                </div>'})
            </div>
            
            <div class="card">
                <div class="card-header">
                    <span class="icon">?</span>
                    <span class="card-title">Deletion Timeline</span>
                </div>
                <div class="row">
                    <span class="label">Deletion Type</span>
                    <span class="value">$(if($results.MailboxSoftDeletedDate -ne 'N/A'){'Soft Delete'}elseif($results.AccountDeletedDate -ne 'N/A'){'Account Deletion'}else{'N/A'})</span>
                </div>
                <div class="row">
                    <span class="label">Deletion Date</span>
                    <span class="value">$(if($results.DeletionDate -ne 'N/A'){$results.DeletionDate.ToString('yyyy-MM-dd HH:mm:ss')}else{'N/A'})</span>
                </div>
                <div class="row">
                    <span class="label">Days Since Delete</span>
                    <span class="value">$($results.DaysSinceDelete)</span>
                </div>
                <div class="row">
                    <span class="label">Days Remaining</span>
                    <span class="value $(if($results.DaysRemaining -ne 'N/A' -and $results.DaysRemaining -lt 7){'deleted'}elseif($results.DaysRemaining -ne 'N/A' -and $results.DaysRemaining -lt 14){'warning'})">$($results.DaysRemaining)</span>
                </div>
                <div class="row">
                    <span class="label">Expiration Date</span>
                    <span class="value">$(if($results.ExpirationDate -ne 'N/A'){$results.ExpirationDate.ToString('yyyy-MM-dd')}else{'N/A'})</span>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <span class="icon">??</span>
                    <span class="card-title">OneDrive Information</span>
                </div>
                <div class="row">
                    <span class="label">OneDrive Status</span>
                    <span class="value $(if($results.OneDriveStatus -like '*Found*'){'active'}else{'warning'})">$($results.OneDriveStatus)</span>
                </div>
                <div class="row">
                    <span class="label">Checked URLs</span>
                    <span class="value">$($results.OneDriveTenant)</span>
                </div>
                <div class="row">
                    <span class="label">Used Space</span>
                    <span class="value">$($results.OneDriveStorage) MB</span>
                </div>
                <div class="row">
                    <span class="label">Error</span>
                    <span class="value">$(if($results.OneDriveStatus -eq 'Not Found'){'Not Found / Inaccessible'}else{'N/A'})</span>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <span class="icon">?</span>
                    <span class="card-title">Recommendation</span>
                </div>
                <div class="row">
                    <span class="label">Data Assessment</span>
                    <span class="value">$(if($results.HasData){'Data Present'}else{'No Data Found'})</span>
                </div>
                <div class="row">
                    <span class="label">Recommendation</span>
                    <span class="value $(if($results.HasData){'active'}else{'warning'})">$($results.Recommendation)</span>
                </div>
                <div class="row">
                    <span class="label">Action</span>
                    <span class="value">$($results.Action)</span>
                </div>
            </div>
        </div>
        
        <div class="footer">
            Report generated on $($results.Timestamp) | PowerShell Automation
        </div>
    </div>
</body>
</html>
"@

# Save report
$reportFile = "RetentionReport_$($results.UserAlias)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath = Join-Path (Get-Location) $reportFile

try {
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "? Report saved: $reportPath" -ForegroundColor Green
    Start-Process $reportPath
} catch {
    Write-Host "? Error saving report: $_" -ForegroundColor Red
}

# Summary
Write-Host "`n????????????????????????????????????????????????????????????" -ForegroundColor Cyan
Write-Host "INVESTIGATION COMPLETE" -ForegroundColor Cyan
Write-Host "????????????????????????????????????????????????????????????" -ForegroundColor Cyan
Write-Host "User: $($results.UserAlias)"
Write-Host "Status: $($results.AccountStatus)"
Write-Host "Mailbox: $($results.MailboxStatus)"
Write-Host "OneDrive: $($results.OneDriveStatus)"
Write-Host "Has Data: $(if($results.HasData){'YES'}else{'NO'})"
Write-Host "Recommendation: $($results.Recommendation)"
Write-Host "????????????????????????????????????????????????????????????" -ForegroundColor Cyan