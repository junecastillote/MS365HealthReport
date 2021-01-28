# Function to List all Office 365 subscribed services in the Tenant
Function Get-MS365SubscribedServices {
    [cmdletbinding()]
    param(
        [parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        $Token
    )
    $tenantID = ($token | Get-JWTDetails).tid
    $header = @{'Authorization' = "Bearer $($Token)" }
    $result = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantID)/ServiceComms/Services" -Headers $header -Method Get -ContentType 'application/json')
    return ($result.value | Sort-Object DisplayName | Select-Object ID, DisplayName)
}