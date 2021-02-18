$AppID = ''
$TenantID = '<org>.onmicrosoft.com'
$ClientSecret = ''
# $thumbprint = ''
# $certificate = get-item -path "Cert:\currentuser\my\$thumbprint"

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
    # StartFromLastRun = $true
    LastUpdatedTime = (Get-Date).AddDays(-10)
    # Verbose = $true
    # Consolidate = $true
}
New-MS365IncidentReport @reportSplat