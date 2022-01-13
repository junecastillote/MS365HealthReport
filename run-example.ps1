$reportSplat = @{
    OrganizationName = 'Organization Name Here'
    ClientID = 'Client ID here'
    ClientSecret = 'Client Secret here'
    # ClientCertificate = get-item -path "Cert:\currentuser\my\$thumbprint"
    TenantID = 'Tenant ID here'
    SendEmail = $true
    From = 'sender@domain.com'
    To = @('to@domain.com')
    LastUpdatedTime = (Get-Date).AddDays(-10)
    # Workload = @('Exchange Online','SharePoint Online')
    Status = 'Ongoing'
}

New-MS365IncidentReport @reportSplat