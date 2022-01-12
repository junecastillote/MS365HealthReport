$AppID = ''
$TenantID = '<org>.onmicrosoft.com'
$ClientSecret = ''
# $thumbprint = ''
# $certificate = get-item -path "Cert:\currentuser\my\$thumbprint"

Remove-Module MS365HealthReport -ErrorAction SilentlyContinue
Import-Module MS365HealthReport

$reportSplat = @{
    OrganizationName = 'Organization Name Here'
    ClientID = $AppID
    ClientSecret = $ClientSecret
    # ClientCertificate = $certificate
    # ClientCertificateThumbprint = $thumbprint
    TenantID = $TenantID
    SendEmail = $true
    From = 'sender@domain.com'
    To = @('to@domain.com')
    # CC = @()
    # BCC = @()
    # Workload = @('SharePoint','Exchange')
    Status = 'Ongoing'
    # StartFromLastRun = $true
    LastUpdatedTime = (Get-Date).AddDays(-10)
    # Verbose = $true
    # Consolidate = $false
}
New-MS365IncidentReport @reportSplat